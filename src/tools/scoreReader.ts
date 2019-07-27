import { capitalize as cap, countBy, range, trim } from 'lodash'
import { utils, WorkSheet } from 'xlsx'
// tslint:disable-next-line: max-line-length
import { CellAddressDict, cellsForRow, cellVal, forEachPeriodTeam, getAddressOfCol, initializeFirstRow, jamNumberValidator, mySPCheck, spCheck, teams } from './utils'

const pointsAndNPCheck = /(\d)\+NP/i
const ippRe = /(\d)\+(\d)/

export class ScoreReader {
    private sbData: DerbyJson.IGame
    private sbTemplate: IStatsbookTemplate
    private sbErrors: IErrorSummary
    private warningData: IWarningData

    private starPasses: Array<{ period: string; jam: number; }> = []

    private maxJams: number

    constructor(sbData: DerbyJson.IGame,
                sbTemplate: IStatsbookTemplate,
                sbErrors: IErrorSummary,
                warningData: IWarningData) {

        this.sbData = sbData
        this.sbTemplate = sbTemplate
        this.sbErrors = sbErrors
        this.warningData = warningData

        this.maxJams = this.sbTemplate.score.maxJams
    }

    public parseSheet(sheet: WorkSheet) {
        forEachPeriodTeam((period, team) => {
            const cells = this.buildFirstRow(period, team)
            const maxTrips = cells.lastTrip.c - cells.firstTrip.c

            let skaterRef: string = ''
            let jam: DerbyJson.IJam = null
            let jamIdx = 0
            let tripCount = 1
            let starPass = false

            range(0, this.maxJams).forEach((rowIdx) => {

                const rowCells = cellsForRow(rowIdx, cells)
                const jamNumber = trim(cellVal(sheet, rowCells.jamNumber))
                const skaterNumber = trim(cellVal(sheet, rowCells.jammerNumber))
                const lead = cellVal(sheet, rowCells.lead) !== undefined
                const lost = cellVal(sheet, rowCells.lost) !== undefined
                const call = cellVal(sheet, rowCells.call) !== undefined
                const inj = cellVal(sheet, rowCells.inj) !== undefined
                const initialCompleted = cellVal(sheet, rowCells.np) === undefined
                const rowDescription =
                `Team: ${cap(team)}, Period: ${period}, Jam: ${jamNumber}, Jammer: ${skaterNumber || ''}`

                if (!jamNumber) {
                    return
                }

                if (!jamNumberValidator.test(jamNumber)) {
                    // it is impossible to process an invalid jam number, so don't try
                    throw new Error(`Invalid Jam Number in cell ${rowCells.jamNumber}: ${jamNumber}`)
                }

                if (spCheck.test(jamNumber)) {
                    if (rowIdx === 0) {
                        // It is impossible to process a SP on the first row, so don't try
                        throw new Error(`SP or SP* Cannot appear on first line`)
                    }

                    starPass = true
                    if (mySPCheck.test(jamNumber)) {
                        // this pushes an event for the prior jammer
                        jam.events.push(
                            {
                                event: 'star pass',
                                skater: skaterRef,
                            })
                    }
                    this.starPasses.push({ period, jam: jamIdx })
                } else {
                    jamIdx = parseInt(jamNumber) - 1
                    tripCount = 1
                    starPass = false
                    if (parseInt(jamNumber) !== (jamIdx + 1)) {
                        this.sbErrors.scores.badJamNumber.events.push(rowDescription)

                        // Add jam objects for missing jams.
                        for (let j = jamIdx + 1; j < parseInt(jamNumber); j++) {
                            this.sbData.periods[period].jams[j - 1] = { number: j, events: [] }
                        }
                    }
                }

                jam = this.sbData.periods[period].jams[jamIdx]

                if (!jam) {
                    jam = {number: jamIdx + 1, events: []}
                    this.sbData.periods[period].jams[jamIdx] = jam
                }

                if (skaterNumber &&
                    !this.sbData.teams[team].persons.find((x) => x.number === skaterNumber)) {
                    // This SHOULD be caught by conditional formatting in Excel, but there
                    // are reports of that breaking sometimes.
                    this.sbErrors.scores.scoresNotOnIGRF.events.push(
                        `${rowDescription}, Skater: ${skaterNumber || ''} `,
                    )
                }

                // push an initial trip object if:
                // 1. not a star pass OR
                // 2. THIS team passed the star, and it's still the initial trip when they did
                if (!starPass || (mySPCheck.test(jamNumber) && tripCount === 1)) {
                    skaterRef = team + ':' + skaterNumber
                    jam.events.push(
                        {
                            event: 'pass',
                            number: 1,
                            score: 0,
                            skater: skaterRef,
                            team,
                            completed: initialCompleted,
                        },
                    )
                } else {
                    if (skaterNumber) {
                        this.sbErrors.scores.spStarWithJammer.events.push(rowDescription)
                    }
                }

                let blankTrip = false
                range(1, maxTrips).forEach((colIdx) => {
                    const tripAddress = getAddressOfCol(colIdx - 1 , utils.decode_cell(rowCells.trip))
                    const tripScore = cellVal(sheet, tripAddress)

                    if (tripScore === undefined) {
                        // ERROR CHECK - no trip score, initial pass completed
                        if (initialCompleted && colIdx === 1 && !starPass) {
                            const nextRow = cellsForRow(jamIdx + 1, cells)
                            const nextJamNumber = trim(cellVal(sheet, nextRow.jamNumber))
                            if (mySPCheck.test(nextJamNumber)) {
                                this.sbErrors.warnings.SPNoPointsNoNI.events.push(rowDescription)
                            } else {
                                this.sbErrors.scores.noPointsNoNI.events.push(rowDescription)
                            }
                        }
                        blankTrip = true
                        return
                    }

                    // ERROR CHECK - points entered for a trip that's already been completed.
                    if (colIdx + 1 < tripCount) {
                        this.sbErrors.scores.spPointsBothJammers.events.push(rowDescription)
                    }

                    // ERROR CHECK - skipped column in a non star pass line
                    if (blankTrip && !starPass) {
                        blankTrip = false
                        this.sbErrors.scores.blankTrip.events.push(rowDescription)
                    }

                    const reResult = pointsAndNPCheck.exec(tripScore)
                 //   const ippResult = ippRe.exec(tripScore)
                    let points = 0

                    if (reResult !== null) {
                        points = parseInt(reResult[1])
                        const trip = jam
                        .events
                        .find((e) =>  e.event === 'pass' && e.skater === skaterRef && e.number === 1)

                        if (trip) {
                            trip.score = points
                        }
                        // TODO: add support for x + y parsing
                    } else {
                        points = parseInt(tripScore)
                        if (!starPass) {
                            tripCount = tripCount + 1
                        }

                        jam.events.push(
                            {
                                event: 'pass',
                                number: tripCount,
                                score: points,
                                skater: skaterRef,
                                team,
                            },
                        )
                    }

                    if (!initialCompleted && !reResult) {
                        this.sbErrors.scores.npPoints.events.push(rowDescription)
                    }

                })

                if (lead) {
                    jam.events.push({
                        event: 'lead',
                        skater: skaterRef,
                    })
                }

                if (lost) {
                    jam.events.push({
                        event: 'lost',
                        skater: skaterRef,
                    })

                    this.warningData.lost.push({
                        period,
                        team,
                        jam: jam.number,
                        skater: skaterRef,
                    })
                }

                if (call) {
                    jam.events.push({
                        event: 'call',
                        skater: skaterRef,
                    })
                }

                if (inj) {
                    this.warningData.jamsCalledInjury.push(
                        {
                            team,
                            period,
                            jam: jam.number,
                        },
                    )
                }

                // Error check - SP and lead without lost
                if (mySPCheck.test(jamNumber) && lead && !lost) {
                    this.sbErrors.scores.spLeadNoLost.events.push(rowDescription)
                }
            })

            if (team === 'away') {
                this.sbData.periods[period].jams
                .forEach((j) => {
                    const jamDescription =  `Period: ${period}, Jam: ${j.number}`
                    const counts = countBy(j.events, (ev) => ev.event)
                    if (counts.lead > 1) {
                        this.sbErrors.scores.tooManyLead.events.push(jamDescription)
                    }

                    if (counts.call > 1) {
                        this.sbErrors.scores.tooManyCall.events.push(jamDescription)
                    }

                    const injuryCount = this.warningData.jamsCalledInjury
                                        .filter((ev) => ev.period === period && ev.jam === j.number)
                                        .length

                    // ERROR CHECK: Injury box checked for only one team in a jam.
                    if (injuryCount === 1)  {
                        this.sbErrors.scores.injuryOnlyOnce.events.push(jamDescription)
                    }

                    // Push injury event here, so that only one is pushed per jam instead of two
                    if (injuryCount > 1) {
                        j.events.push({
                            event: 'injury',
                        })
                    }

                    if (counts.lead === 0) {
                        teams.forEach((teamname) => {
                            // tslint:disable-next-line: max-line-length
                            const lost = !!j.events.find((ev) => ev.event === 'lost' && ev.skater.startsWith(teamname))
                            // tslint:disable-next-line: max-line-length
                            const points = !!j.events.find((ev) => ev.event === 'pass' && ev.team === teamname && ev.number > 1)

                            if (points && !lost) {
                                this.sbErrors.scores.pointsNoLeadNoLost.events.push(
                                    `Team: ${cap(teamname)}, ${jamDescription}`,
                                )
                            }
                        })
                    }

                })
            }
        })

        // Error check: Star pass marked for only one team in a jam.
        const spCount = countBy(this.starPasses, (sp) => `Period: ${sp.period} Jam: ${sp.jam}`)

        Object.keys(spCount)
            .filter((sp) => spCount[sp] === 1)
            .forEach((spDesc) => {
                this.sbErrors.scores.onlyOneStarPass.events.push(spDesc)
            })
    }

    private buildFirstRow(period: period, team: team): CellAddressDict {
        const fields = ['firstJamNumber', 'firstJammerNumber', 'firstLost', 'firstLead',
            'firstCall', 'firstInj', 'firstNp', 'firstTrip', 'lastTrip']

        return initializeFirstRow(this.sbTemplate, 'score', team, period, fields)
    }

}

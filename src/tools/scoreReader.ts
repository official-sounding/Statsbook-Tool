import { capitalize as cap, get, range, trim } from 'lodash'
import { WorkSheet } from 'xlsx/types'
// tslint:disable-next-line: max-line-length
import { CellAddressDict, cellsForRow, cellVal, forEachPeriodTeam, getAddressOfRow, getAddressOfTrip, initializeFirstRow } from './utils'

const jamNumberValidator = /^(\d+|SP|SP\*)$/i
const spCheck = /^SP\*?$/i
const mySPCheck = /^SP$/i

export class ScoreReader {
    private sbData: DerbyJson.IGame
    private sbTemplate: IStatsbookTemplate
    private sbErrors: IErrorSummary

    private starPasses: Array<{ period: string; jam: number; }> = []

    private maxJams: number
    private tab = 'score'

    constructor(sbData, sbTemplate, sbErrors) {
        this.sbData = sbData
        this.sbTemplate = sbTemplate
        this.sbErrors = sbErrors

        this.maxJams = this.sbTemplate.score.maxJams
    }

    public parseScoreSheet(sheet: WorkSheet) {
        forEachPeriodTeam((period, team) => {
            const cells = this.buildFirstRow(period, team)
            const maxTrips = cells.lastTrip.c - cells.firstTrip.c

            let skaterRef: string = ''
            const priorJam: DerbyJson.IJam = null
            let jamIdx = 0
            const tripCount = 1
            let starPass = false

            range(0, this.maxJams).forEach((rowIdx) => {
                const isLost = false
                const isLead = false

                const rowCells = cellsForRow(rowIdx, cells)
                const jamNumber = trim(cellVal(sheet, rowCells.jamNumber))
                const skaterNumber = trim(cellVal(sheet, rowCells.jammerNumber))
                const initialCompleted = cellVal(sheet, rowCells.np) !== undefined
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
                        priorJam.events.push(
                            {
                                event: 'star pass',
                                skater: skaterRef,
                            })
                    }
                    this.starPasses.push({ period, jam: jamIdx })
                } else {
                    if (parseInt(jamNumber) !== (jamIdx + 1)) {
                        this.sbErrors.scores.badJamNumber.events.push(rowDescription)

                        // Add jam objects for missing jams.
                        for (let j = jamIdx + 1; j < parseInt(jamNumber); j++) {
                            this.sbData.periods[period].jams[j - 1] = { number: j, events: [] }
                        }
                    }

                    jamIdx = parseInt(jamNumber) - 1
                }

                let jam = this.sbData.periods[period].jams[jamIdx]

                if (!jam) {
                    jam = {number: jamIdx + 1, events: []}
                    this.sbData.periods[period].jams[jamIdx] = jam
                }

                if (skaterNumber &&
                    this.sbData.teams[team].persons.find((x) => x.number === skaterNumber)) {
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
                range(1, maxTrips).forEach((tripIdx) => {
                    const tripAddress = getAddressOfTrip(tripIdx, rowCells.firstTrip)
                    const tripScore = cellVal(sheet, tripAddress)

                    if (!tripScore) {
                        // ERROR CHECK - no trip score, initial pass completed
                        if (initialCompleted && tripIdx === 1 && !starPass) {
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

                })
            })
        })
    }

    private buildFirstRow(period: string, team: string): CellAddressDict {
        const fields = ['firstJamNumber', 'firstJammerNumber', 'firstLost', 'firstLead',
            'firstCall', 'firstInj', 'firstNp', 'firstTrip', 'lastTrip']

        return initializeFirstRow(this.sbTemplate, 'score', team, period, fields)
    }

}

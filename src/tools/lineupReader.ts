import { capitalize as cap, countBy, last, range } from 'lodash'
import { utils, WorkSheet } from 'xlsx'
// tslint:disable-next-line: max-line-length
import { CellAddressDict, cellsForRow, cellVal, forEachPeriodTeam, getAddressOfCol, initializeFirstRow, jamNumberValidator, mySPCheck, spCheck } from './utils'

const positions = ['jammer', 'pivot', 'blocker', 'blocker', 'blocker']
const positionCount = positions.length

function buildPenaltyFinder(team: team) {
    return (x: DerbyJson.IEvent): boolean => (x.event === 'penalty' && x.skater.startsWith(team))
}

export class LineupReader {
    private sbData: DerbyJson.IGame
    private sbTemplate: IStatsbookTemplate
    private sbErrors: IErrorSummary
    private warningData: IWarningData
    private boxTripReader: IBoxTripReader

    constructor(sbData: DerbyJson.IGame,
                sbTemplate: IStatsbookTemplate,
                sbErrors: IErrorSummary,
                warningData: IWarningData,
                boxTripReader: IBoxTripReader) {

        this.sbData = sbData
        this.sbTemplate = sbTemplate
        this.sbErrors = sbErrors
        this.warningData = warningData
        this.boxTripReader = boxTripReader
    }

    public parseSheet(sheet: WorkSheet): void {
        const boxCodeCount = this.sbTemplate.lineups.boxCodes
        const maxJams = this.sbTemplate.lineups.maxJams

        forEachPeriodTeam((period, team) => {
            let jamIdx = 0
            let starPass = false
            let skaterNumbers: string[] = []

            const firstCells = this.buildFirstRow(period, team)
            const sectionDesc = `Team: ${cap(team)}, Period: ${period}`

            range(0, maxJams).forEach((rowIdx) => {
                const rowCells = cellsForRow(rowIdx, firstCells)

                const jamNumber = cellVal(sheet, rowCells.jamNumber)
                const noPivot = cellVal(sheet, rowCells.noPivot)

                if (!jamNumberValidator.test(jamNumber)) {
                    return
                }

                if (spCheck.test(jamNumber)) {
                    starPass = true
                    if (!mySPCheck.test(jamNumber)) {
                        const spStarSkater = range(0, positionCount).some((colIdx) => {
                            const skaterAddr = getAddressOfCol(colIdx, utils.decode_cell(rowCells.jammer))
                            const skaterNum = cellVal(sheet, skaterAddr)
                            return !!skaterNum
                        })

                        if (spStarSkater) {
                            this.sbErrors.lineups.spStarSkater.events
                                .push(`${sectionDesc}, Jam: ${jamIdx}`)
                        }
                        return
                    } else if (!noPivot) {
                        this.sbErrors.lineups.starPassNoPivot.events
                        .push(`${sectionDesc}, Jam: ${jamIdx}`)
                    }
                } else {
                    starPass = false
                    jamIdx = parseInt(jamNumber)
                    skaterNumbers = []
                }

                // Retrieve penalties from this jam and prior jam for
                // error checking later
                // tslint:disable-next-line: max-line-length
                const penaltyFinder = buildPenaltyFinder(team)
                const jam = this.sbData.periods[period].jams[jamIdx - 1]

                const thisJamPenalties = jam.events.filter(penaltyFinder)
                let priorJamPenalties = []

                if (jamIdx !== 1) {
                    const lastJam = this.sbData.periods[period].jams[jamIdx - 2]

                    priorJamPenalties = lastJam.events
                    .filter(penaltyFinder)
                } else if (period === '2') {
                    const lastJam = last(this.sbData.periods['1'].jams)

                    priorJamPenalties = lastJam.events
                    .filter(penaltyFinder)
                }

                range(0, positionCount).forEach((colIdx) => {
                    const skaterAddr = getAddressOfCol(colIdx * (boxCodeCount + 1), utils.decode_cell(rowCells.jammer))
                    const skaterNumCell = sheet[skaterAddr]
                    const skaterNumber = cellVal(sheet, skaterAddr)
                    let position = ''

                    if (!skaterNumber
                        || skaterNumber === '?'
                        || skaterNumber.toLowerCase() === 'n/a'
                        ) {
                        // WARNING: Empty box on Lineups without comment
                        if (!skaterNumber && skaterNumCell.c === undefined) {
                            this.sbErrors.warnings.emptyLineupNoComment.events
                            .push(`${sectionDesc}, Jam: ${jamIdx}, Column: ${colIdx + 1}`)
                        }

                        return
                    }

                    const skater = `${team}:${skaterNumber}`

                    // ERROR CHECK: Skater on lineups not on IGRF
                    if (!this.sbData.teams[team].persons.find((x) => x.number === skaterNumber)) {
                        // If the skater is not on the IGRF, record an error.
                        // This SHOULD be caught by conditional formatting in Excel, but there
                        // are reports of that breaking sometimes.
                        this.sbErrors.lineups.lineupsNotOnIGRF.events
                        .push(`${sectionDesc}, Jam: ${jamIdx}, Skater: ${skaterNumber}`)
                    }

                    // ERROR CHECK: Same skater entered more than once per jam
                    if (skaterNumbers.includes(skater) && !starPass) {
                        this.sbErrors.lineups.samePlayerTwice.events
                        .push(`${sectionDesc}, Jam: ${jamIdx}, Skater: ${skaterNumber}`)
                    }

                    if (!starPass) {
                        skaterNumbers.push(skater)
                    }

                    if (colIdx === 1 && !!noPivot) {
                        position = 'blocker'
                    } else {
                        position = positions[colIdx]
                    }

                    if (!starPass) {
                        // Unless this is a star pass, add a
                        // "lineup" event for that skater with the position
                        jam.events.push(
                            {
                                event: 'lineup',
                                skater,
                                position,
                            })
                    }

                    function hasPenalty(btwnJams: boolean): boolean {
                        if (btwnJams) {
                            return !!thisJamPenalties.find((x) => x.skater === skater)
                                || !!priorJamPenalties.find((x) => x.skater === skater)
                        } else {
                            return !!thisJamPenalties.find((x) => x.skater === skater)
                        }
                    }

                    const priorFoulout = !!this.warningData.foulouts.find((x) =>
                                    (x.period === period && x.jam < jamIdx && x.skater === skater) ||
                                    (x.period < period && x.skater === skater))

                    let noCodes = true

                    range(0, boxCodeCount).forEach((boxColIdx) => {
                        const boxCol = colIdx * (boxCodeCount + 1) + boxColIdx
                        const boxAddr = getAddressOfCol(boxCol, utils.decode_cell(rowCells.jammer))
                        const boxCode = cellVal(sheet, boxAddr)

                        if (!boxCode) {
                            return
                        } else {
                            noCodes = false
                        }

                        const event = this.boxTripReader.parseGlyph(boxCode, team, skater)

                        switch (event.eventType) {
                            case 'enter':
                                this.enterBox(jam, skater, event.note)
                                if (!hasPenalty(event.betweenJams)) {
                                    const errKey = event.betweenJams ?
                                                        this.boxTripReader.badStartErrorKey :
                                                        this.boxTripReader.badBtwnJamErrorKey
                                    this.sbErrors
                                        .lineups[errKey]
                                        .events.push(`${sectionDesc}, Jam: ${jamIdx}, Skater: ${skaterNumber}`)

                                    this.warningData.badStarts.push({
                                        period,
                                        team,
                                        skater,
                                        jam: jamIdx,
                                    })
                                }
                                break
                            case 'exit':
                                this.exitBox(jam, skater, event.note)
                                break
                            case 'enterExit':
                                this.enterBox(jam, skater, event.note)
                                this.exitBox(jam, skater, event.note)
                                if (!hasPenalty(event.betweenJams)) {
                                    const errKey = event.betweenJams ?
                                                        this.boxTripReader.badCompleteErrorKey :
                                                        this.boxTripReader.badBtwnJamCompleteErrorKey

                                    this.sbErrors
                                        .lineups[errKey]
                                        .events.push(`${sectionDesc}, Jam: ${jamIdx}, Skater: ${skaterNumber}`)

                                    this.warningData.badContinues.push({
                                        skater,
                                        team,
                                        period,
                                        jam: jamIdx,
                                    })
                                }
                                break
                            case 'badContinue':
                                if (priorFoulout) {
                                    this.sbErrors.lineups.foInBox.events
                                        .push(`${sectionDesc}, Jam: ${jamIdx}, Skater: ${skaterNumber}`)
                                } else {
                                    const errKey = this.boxTripReader.badContinueErrorKey

                                    this.sbErrors
                                        .lineups[errKey]
                                        .events.push(`${sectionDesc}, Jam: ${jamIdx}, Skater: ${skaterNumber}`)
                                }
                                this.warningData.badContinues.push({
                                    skater,
                                    team,
                                    period,
                                    jam: jamIdx,
                                })
                                break
                            case 'injury':
                                this.warningData.lineupThree.push({
                                    skater,
                                    team,
                                    period,
                                    jam: jamIdx,
                                })
                                break
                            case 'error':
                                this.sbErrors
                                    .lineups[event.errorKey]
                                    .events.push(`${sectionDesc}, Jam: ${jamIdx}, Skater: ${skaterNumber}`)
                                break
                        }
                    })

                    if (this.boxTripReader.stillInBox(team, skater) && noCodes) {
                        this.sbErrors.lineups.seatedNoCode
                        .events.push(`${sectionDesc}, Jam: ${jamIdx}, Skater: ${skaterNumber}`)
                    }
                })

                // Remove fouled out skaters from the box
                this.warningData.foulouts
                    .filter((x) => x.period === period && x.team === team && x.jam === jamIdx)
                    .forEach((fo) => this.boxTripReader.removeFromBox(fo.team, fo.skater))

                // Remove Expelled skaters from the box
                this.warningData.expulsions
                    .filter((x) => x.period === period && x.team === team && x.jam === jamIdx)
                    .forEach((fo) => this.boxTripReader.removeFromBox(fo.team, fo.skater))

                // Skaters in the box not listed on the lineup for this jam
                this.boxTripReader.missingSkaters(team, skaterNumbers)
                    .forEach((missing) => {
                        const missingNumber = missing.slice(5)

                        this.sbErrors.lineups.seatedNotLinedUp.events
                        .push(`${sectionDesc}, Jam: ${jamNumber}, Skater: ${missingNumber}`)

                        this.warningData.noExits.push({
                            team,
                            period,
                            jam: jamIdx,
                            skater: missing,
                        })
                    })

                thisJamPenalties
                    .filter((p) => !skaterNumbers.includes(p.skater))
                    .forEach((missing) => {
                        const missingNumber = missing.skater.slice(5)

                        this.sbErrors.penalties.penaltyNoLineup.events
                        .push(`${sectionDesc}, Jam: ${jamNumber}, Skater: ${missingNumber}`)
                    })

            })
        })

    }

    private enterBox(jam: DerbyJson.IJam, skater: string, note: string = null) {
        const event: DerbyJson.IEvent = {
            event: 'enter box',
            skater,
        }

        if (note) {
            event.notes = [{ note }]
        }

        jam.events.push(event)
    }

    private exitBox(jam: DerbyJson.IJam, skater: string, note: string = null) {
        const event: DerbyJson.IEvent = {
            event: 'exit box',
            skater,
        }

        if (note) {
            event.notes = [{ note }]
        }

        jam.events.push(event)
    }

    private buildFirstRow(period: period, team: team): CellAddressDict {
        const fields = ['firstJamNumber', 'firstNoPivot', 'firstJammer']

        return initializeFirstRow(this.sbTemplate, 'score', team, period, fields)
    }
}

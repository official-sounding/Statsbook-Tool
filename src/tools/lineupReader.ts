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

    constructor(sbData: DerbyJson.IGame,
                sbTemplate: IStatsbookTemplate,
                sbErrors: IErrorSummary,
                warningData: IWarningData) {

        this.sbData = sbData
        this.sbTemplate = sbTemplate
        this.sbErrors = sbErrors
        this.warningData = warningData
    }

    public parseSheet(sheet: WorkSheet): void {
        const boxCodeCount = this.sbTemplate.lineups.boxCodes
        const maxJams = this.sbTemplate.lineups.maxJams

        forEachPeriodTeam((period, team) => {
            let jamIdx = 0
            let starPass = false
            let skaterNumbers = []

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

                })

            })
        })

    }

    private buildFirstRow(period: period, team: team): CellAddressDict {
        const fields = ['firstJamNumber', 'firstNoPivot', 'firstJammer']

        return initializeFirstRow(this.sbTemplate, 'score', team, period, fields)
    }
}

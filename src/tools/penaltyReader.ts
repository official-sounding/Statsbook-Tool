import { capitalize as cap, countBy, inRange, range, trim } from 'lodash'
import { utils, WorkSheet } from 'xlsx'
// tslint:disable-next-line: max-line-length
import { CellAddressDict, cellsForRow, cellVal, forEachPeriodTeam, getAddressOfCol, initializeFirstRow } from './utils'

export class PenaltyReader {
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
        const maxPen = this.sbTemplate.penalties.maxPenalties
        const foulOuts: { [skater: string]: boolean } = {}
        const skaterPenaltyMap: { [skater: string]: number } = {}

        forEachPeriodTeam((period, team) => {
            const sectionDesc = `Team: ${cap(team)}, Period: ${period}`
            const maxNum = this.sbTemplate.teams[team].maxNum
            const maxJam = this.sbData.periods[period].jams.length
            const firstRow = this.buildFirstRow(period, team)

            range(0, maxNum).forEach((rowIdx) => {
                const row = cellsForRow(rowIdx * 2, firstRow)
                const skaterNumber = cellVal(sheet, row.number)

                if (skaterNumber === undefined) {
                    return
                }

                // ERROR CHECK: skater on penalty sheet not on the IGRF
                if (this.sbData.teams[team].persons.find((x) => x.number === skaterNumber)) {
                    this.sbErrors.penalties.penaltiesNotOnIGRF.events
                        .push(`${sectionDesc}, Skater: ${skaterNumber}`)
                }

                const skater = `${team}:${skaterNumber}`

                if (!skaterPenaltyMap[skater]) {
                    skaterPenaltyMap[skater] = 0
                }

                range(0, maxPen).forEach((colIdx) => {
                    const penAddr = getAddressOfCol(colIdx, utils.decode_cell(row.penalty))
                    const jamAddr = getAddressOfCol(colIdx, utils.decode_cell(row.jam))

                    const code = cellVal(sheet, penAddr)
                    const jam = cellVal(sheet, jamAddr)

                    // if both empty, skip
                    if (code === undefined && jam === undefined) {
                        return
                    } else if (code === undefined || jam === undefined) {
                        this.sbErrors.penalties.codeNoJam.events
                            .push(`${sectionDesc}, Skater: ${skaterNumber}`)
                    }

                    const jnParsed = parseInt(jam)
                    if (isNaN(jnParsed) || !inRange(jnParsed, 1, maxJam)) {
                        this.sbErrors.penalties.penaltyBadJam.events
                            .push(`${sectionDesc}, Skater: ${skaterNumber}, Recorded Jam: ${jam}`)
                        return
                    }

                    this.sbData.periods[period].jams[jnParsed - 1]
                        .events.push({
                            event: 'penalty',
                            skater,
                            penalty: code,
                        })
                    skaterPenaltyMap[skater]++

                })

                const foCode = cellVal(sheet, row.FO)
                const foJam = cellVal(sheet, row.FOJam)

                if (foCode === undefined || foJam === undefined) {

                    if (foCode !== undefined || foJam !== undefined) {
                        this.sbErrors.penalties.codeNoJam.events
                            .push(`${sectionDesc}, Skater: ${skaterNumber}`)
                    } else if (period === '2' &&
                           !foulOuts[skater] &&
                           skaterPenaltyMap[skater] !== undefined &&
                           skaterPenaltyMap[skater] >= 7) {

                        this.sbErrors.penalties.sevenWithoutFO.events
                            .push(`Team: ${cap(team)}, Skater: ${skaterNumber}`)
                    }

                    return
                }

                const foJnParsed = parseInt(foJam)
                if (isNaN(foJnParsed) || !inRange(foJnParsed, 1, maxJam)) {
                    this.sbErrors.penalties.foBadJam.events
                        .push(`${sectionDesc}, Skater: ${skaterNumber}, Jam: ${foJam}`)
                    return
                }

                // If there is expulsion, add an event.
                // Note that derbyJSON doesn't actually record foul-outs,
                // so only expulsions are recorded.
                if (foCode !== 'FO') {
                    const jam: DerbyJson.IJam = this.sbData.periods[period].jams[foJnParsed - 1]
                    jam.events.push(
                        {
                            event: 'expulsion',
                            skater,
                            notes: [
                                {note: `Penalty: ${foCode}` },
                                {note: `Jam: ${foJam}` },
                            ],
                        },
                    )
                    this.warningData.expulsions.push(
                        {
                            skater,
                            team,
                            period,
                            jam: foJnParsed,
                        },
                    )

                    // ERROR CHECK: Expulsion code for a jam with no penalty
                    if (jam.events.find((x) => x.event === 'penalty' && x.skater === skater)) {
                        this.sbErrors.penalties.expulsionNoPenalty.events
                            .push(`${sectionDesc}, Skater: ${skaterNumber}, Jam: ${foJam}`)
                    }

                } else {
                    foulOuts[skater] = true
                    this.warningData.foulouts.push(
                        {
                            skater,
                            team,
                            period,
                            jam: foJnParsed,
                        },
                    )

                    if (skaterPenaltyMap[skater] < 7) {
                        this.sbErrors.penalties.foUnder7.events.push(
                            `${sectionDesc}, Skater: ${skaterNumber}`,
                        )
                    }
                }
            })

            range(0, 1).forEach((colIdx) => {
                const benchExpCodeAddr = getAddressOfCol(colIdx, firstRow.benchExpCode)
                const benchExpJamAddr = getAddressOfCol(colIdx, firstRow.benchExpJam)

                const benchExpCode = cellVal(sheet, benchExpCodeAddr)
                const benchExpJam = cellVal(sheet, benchExpJamAddr)

                if (benchExpCode === undefined || benchExpJam === undefined) {
                    return
                }

                const bExpJnParsed = parseInt(benchExpJam)
                if (isNaN(bExpJnParsed) || !inRange(bExpJnParsed, 1, maxJam)) {
                    // TODO: warn on this thing that will likely never happen
                    return
                }

                this.sbData.periods[period].jams[bExpJnParsed - 1].events.push(
                    {
                        event: 'expulsion',
                        notes: [
                            {note: `Bench Staff Expulsion: ${benchExpCode}`},
                            {note: `Jam: ${benchExpJam}` },
                        ],
                    },
                )
            })
        })
    }

    private buildFirstRow(period: string, team: string): CellAddressDict {
        const fields = ['firstnumber', 'firstpenalty', 'firstjam',
        'firstFO', 'firstFOJam', 'benchExpCode', 'benchExpJam']

        return initializeFirstRow(this.sbTemplate, 'penalties', team, period, fields)
    }
}

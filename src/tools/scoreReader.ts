import { capitalize as cap, get, range, trim } from 'lodash'
import { WorkSheet } from 'xlsx/types'
import { CellAddressDict, cellsForRow, cellVal, forEachPeriodTeam, getAddressOfRow, initializeFirstRow } from './utils';

const jamNumberValidator = /^(\d+|SP|SP\*)$/i
const spCheck = /^SP\*?$/i
const mySPCheck = /^SP$/i

export class ScoreReader {
    private sbData: DerbyJson.IGame;
    private sbTemplate: IStatsbookTemplate;
    private sbErrors: IErrorSummary;

    private starPasses: Array<{ period: string; jam: number; }> = []

    private maxJams: number;
    private tab = 'score';

    constructor(sbData, sbTemplate, sbErrors) {
        this.sbData = sbData;
        this.sbTemplate = sbTemplate;
        this.sbErrors = sbErrors;

        this.maxJams = this.sbTemplate.score.maxJams
    }

    public parseScoreSheet(sheet: WorkSheet) {
        forEachPeriodTeam((period, team) => {
            const cells = this.buildFirstRow(period, team);
            const maxTrips = cells.lastTrip.c - cells.firstTrip.c;

            const skaterRef: string = ''
            const jamIdx = 0
            const priorJam: DerbyJson.IJam = null

            range(0, this.maxJams).forEach((rowIdx) => {
                let starPass = false
                const blankTrip = false
                const isLost = false
                const isLead = false

                const rowCells = cellsForRow(rowIdx, cells);
                const jamNumber = trim(cellVal(sheet, rowCells.jamNumber));
                const skaterNumber = trim(cellVal(sheet, rowCells.jammerNumber));

                if (!jamNumber) {
                    return;
                }

                if (!jamNumberValidator.test(jamNumber)) {
                    throw new Error(`Invalid Jam Number in cell ${rowCells.jamNumber}: ${jamNumber}`)
                }

                if (spCheck.test(jamNumber)) {
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
                        this.sbErrors.scores.badJamNumber.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${jamNumber}`,
                        )

                        // Add jam objects for missing jams.
                        for (let j = jamIdx + 1; j < parseInt(jamNumber); j++) {
                            this.sbData.periods[period].jams[j - 1] = { number: j, events: [] }
                        }
                    }
                }
            });
        });
    }

    private buildFirstRow(period: string, team: string): CellAddressDict {
        const fields = ['firstJamNumber', 'firstJammerNumber', 'firstLost', 'firstLead',
            'firstCall', 'firstInj', 'firstNp', 'firstTrip', 'lastTrip'];

        return initializeFirstRow(this.sbTemplate, 'score', team, period, fields);
    }

}

import { WorkBook, WorkSheet, CellAddress, utils } from "xlsx";
import { IStatsbookTemplate, IErrorSummary, IStatsbookSummary } from "../types";
import { IDerbyJsonData } from "../derbyJson.types";

import { get, capitalize as cap, range, trim } from 'lodash';

const template2018: IStatsbookTemplate = require('~/assets/2018statsbook.json')
const template2017: IStatsbookTemplate = require('~/assets/2017statsbook.json')
const errorTemplate: IErrorSummary = require('~/assets/sberrors.json')

const teams = ['home', 'away'];
const periods = ['1','2'];

const jamNumberValidator = /^(\d+|SP|SP\*)$/i
const spCheck = /^SP\*?$/i
const mySPCheck = /^SP$/i

type CellAddressDict = { [key:string]: CellAddress };

export class WorkbookReader {
    static defaultVersion: string = '2018';
    static currentVersion: string = '2019';

    private workbook: WorkBook;
    private sbVersion: string;
    private sbFilename: string;
    private sbTemplate: IStatsbookTemplate;
    private sbErrors: IErrorSummary;
    private sbData: IDerbyJsonData;
    private penalties: { [playerId:string]: any[] };
    private starPasses: { period: string; jam: number; }[] = []

    constructor(workbook: WorkBook, filename: string) {
        this.workbook = workbook;
        this.sbErrors = JSON.parse(JSON.stringify(errorTemplate))
        this.sbFilename = filename
        this.penalties = {};
        this.parseFile();
    }

    get summary(): IStatsbookSummary {
        return {
            version: this.sbVersion,
            filename: this.sbFilename,
        }
    }

    get errors(): IErrorSummary {
        return JSON.parse(JSON.stringify(this.sbErrors))
    }

    get data(): IDerbyJsonData {
        return JSON.parse(JSON.stringify(this.sbData))
    }

    private parseFile(): void {
        this.sbVersion = this.getVersion();
        this.sbTemplate = this.getTemplate();

        this.sbData = {
            date: null,
            time: null,
            venue: {
                name: null,
                city: null,
                state: null
            },
            periods: {
                "1": { jams: []},
                "2": { jams: []}
            },
            teams: {
                home: null,
                away: null,
                officials: { persons: [] }
            }
        }

        this.getIGRF();
        this.getTeams();
    }

    private getVersion(): string {
        const sheet = this.workbook.Sheets['Read Me']
        const versionText = (sheet ? sheet['A3'].v : WorkbookReader.defaultVersion)
        const versionRe = /(\d){4}/
        return versionRe.exec(versionText)[0];
    }

    private getTemplate(): IStatsbookTemplate {
        let result: IStatsbookTemplate;

        if (this.sbVersion != WorkbookReader.currentVersion) {
            this.sbErrors.warnings.oldStatsbookVersion.events.push(
                `This File: ${this.sbVersion}  Current Version: ${WorkbookReader.currentVersion} `
            )
        }

        switch (this.sbVersion) {
            case '2019':
            case '2018':
                result = template2018
                break
            case '2017':
                result = template2017
                break
            default:
                throw `Unable to Load Template for year ${this.sbVersion}`
        }

        return result;
    }

    private getIGRF() {
        const sheet = this.workbook.Sheets[this.sbTemplate.mainSheet]
        const venue = this.sbData.venue;

        venue.name = this.getExpectedValue(sheet, 'venue.name', 'Venue Name')
        venue.city = this.getExpectedValue(sheet, 'venue.city', 'Venue City')
        venue.state = this.getExpectedValue(sheet, 'venue.state', 'Venue State')


        const excelDate = this.getExpectedValue(sheet, 'date', 'Date')
        const excelTime = this.getExpectedValue(sheet, 'time', 'Time')

        this.sbData.date = getJsDateFromExcel(excelDate)
        this.sbData.time = getJsTimeFromExcel(excelTime)
    }

    private getTeams() {
        teams.forEach(team => {
            const teamTemplate = this.sbTemplate.teams[team];
            const sheet = this.workbook.Sheets[teamTemplate.sheetName]

            this.sbData.teams[team] = {
                leauge: cellVal(sheet, teamTemplate.league),
                name: cellVal(sheet, teamTemplate.name),
                color: cellVal(sheet, teamTemplate.color),
                persons: []
            };

            const teamData = this.sbData.teams[team];



            if (!teamData.color) {
                this.sbErrors.warnings.missingData.events.push(
                    `Missing color for ${cap(team)} team.`
                )
            }

            // Extract skater data
            const firstNameRC = utils.decode_cell(teamTemplate.firstName)
            const firstNumRC = utils.decode_cell(teamTemplate.firstNumber)

            range(0, teamTemplate.maxNum).forEach(i => {
                const nameAddr = getAddressOfRow(i, firstNameRC)
                const numAddr = getAddressOfRow(i, firstNumRC)

                const skaterNumber = cellVal(sheet, numAddr)
                const skaterName = cellVal(sheet, nameAddr)

                if(!skaterNumber) {
                    return
                }

                teamData.persons.push({
                    name: skaterName,
                    number: skaterNumber
                })

                this.penalties[`${team}:${skaterNumber}`] = []

            });
        });

        this.getOfficials();
    }

    private getOfficials() {

        const sheet = this.workbook.Sheets[this.sbTemplate.teams.officials.sheetName];
        const template = this.sbTemplate.teams.officials;

        const cells: { [index:string]: CellAddress } = { 
            firstName: utils.decode_cell(template.firstName),
            firstRole: utils.decode_cell(template.firstRole),
            firstLeague: utils.decode_cell(template.firstLeague),
            firstCert: utils.decode_cell(template.firstCert)
        };

        range(0, template.maxNum).forEach(idx => {
            const nameAddr = getAddressOfRow(idx, cells.firstName)
            const roleAddr = getAddressOfRow(idx, cells.firstRole)
            const leagueAddr = getAddressOfRow(idx, cells.firstLeague)
            const certAddr = getAddressOfRow(idx, cells.firstCert)

            const name = cellVal(sheet, nameAddr)
            const role = cellVal(sheet, roleAddr)
            const league = cellVal(sheet, leagueAddr)
            const cert = cellVal(sheet, certAddr)

            if(role === undefined) { return }

            const person = {
                name,
                roles: [ role ],
                league,
                certifications: []
            };

            if(cert !== undefined) {
                person.certifications.push({ level: cert });
            }

            this.sbData.teams.officials.persons.push(person);

        })
    }

    private getScores() {
        const maxJams = this.sbTemplate.score.maxJams
        const sheet = this.workbook.Sheets[this.sbTemplate.score.sheetName]
        const tab = 'score'

        const fields = ['firstJamNumber','firstJammerNumber','firstLost','firstLead',
        'firstCall','firstInj','firstNp','firstTrip','lastTrip'];

        periods.forEach(period => {
            teams.forEach(team => {
                const cells = initializeFirstRow(this.sbTemplate, tab, team, period, fields)
                const maxTrips = cells.lastTrip.c - cells.firstTrip.c
                
                let skaterRef:string = '';

                range(0, maxJams).forEach(jam => {
                    let starPass = false
                    let blankTrip = false
                    let isLost = false
                    let isLead = false

                    const rowCells = cellsForRow(jam, cells);

                    const jamNumber = trim(cellVal(sheet,rowCells.jamNumber))
                    const skaterNumber = trim(cellVal(sheet, rowCells.jammerNumber));

                    
                    if(!jamNumber) {
                        return;
                    }

                    if(!jamNumberValidator.test(jamNumber)) {
                        throw `Invalid Jam Number in cell ${rowCells.jamNumber}: ${jamNumber}`
                    }

                    if(spCheck.test(jamNumber)) {
                        starPass = true
                        if (mySPCheck.test(jamNumber)) {
                            // this pushes an event for the prior jammer
                            this.sbData.periods[period].jams[jam -1].events.push(
                            {
                                event: 'star pass',
                                skater: skaterRef
                            })
                        }
                        this.starPasses.push({period: period, jam: jam})
                    }



                });
            })
        })
    }


    private getExpectedValue(sheet: WorkSheet, field: string, longName: string = null) {
        const address: string = get(this.sbTemplate, field);
        const value = cellVal(sheet, address);
        if (!value && longName) {
            this.sbErrors.warnings.missingData.events.push(longName);
        }

        return value;
    }


}

function initializeFirstRow(template:IStatsbookTemplate, tab: string, team: string, period: string, fields: string[]): CellAddressDict
{
    const result:CellAddressDict = {};

    fields.reduce((prev, curr) => {
        prev[curr] = utils.decode_cell(template[tab][team][period][curr])
        return prev;
    }, result);

    return result;
}

function cellsForRow(idx: number, firstCells: CellAddressDict): { [key:string]: string }
{
    const result: { [key:string]: string } = {};
    Object.keys(firstCells).reduce((prev, curr) => {
        const key = curr.replace('first','');
        prev[key] = getAddressOfRow(idx, firstCells[curr])
        return prev
    }, result)

    return result
}

function getAddressOfRow(idx: number, firstCell: CellAddress): string {
    const rcAddr = Object.assign({}, firstCell);
    rcAddr.r = rcAddr.r + idx;

    return utils.encode_cell(rcAddr);
}

function getJsDateFromExcel(excelDate) {
    // Helper function to convert Excel date to JS format
    if (!excelDate) { return undefined }

    return new Date((excelDate - (25567 + 1)) * 86400 * 1000)
}



function getJsTimeFromExcel(excelTime) {
    // Helper function to convert Excel time to JS format
    if (!excelTime) { return undefined }

    const secondsAfterMid = excelTime * 86400
    const hours = Math.floor(secondsAfterMid / 3600)
    const remainder = secondsAfterMid % 3600
    const minutes = Math.floor(remainder / 60)
    const seconds = remainder % 60

    return `${hours.toString().padStart(2, '0')
        }:${minutes.toString().padStart(2, '0')
        }:${seconds.toString().padStart(2, '0')}`
}



function cellVal(sheet: WorkSheet, address: string) {
    // Given a worksheet and a cell address, return the value
    // in the cell if present, and undefined if not.
    if (sheet[address] && sheet[address].v) {
        return sheet[address].v
    } else {
        return undefined
    }
}
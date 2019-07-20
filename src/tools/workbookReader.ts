import { capitalize as cap, get, range, trim } from 'lodash'
import { CellAddress, utils, WorkBook, WorkSheet } from 'xlsx'
import { ScoreReader } from './scoreReader'
import { CellAddressDict, cellsForRow, cellVal, getAddressOfRow, periods, teams } from './utils'

import template2017 from '../../assets/2017statsbook.json'
import template2018 from '../../assets/2018statsbook.json'
import errorTemplate from '../../assets/sberrors.json'

export class WorkbookReader {
    public static defaultVersion: string = '2018'
    public static currentVersion: string = '2019'

    private workbook: WorkBook
    private sbVersion: string
    private sbFilename: string
    private sbTemplate: IStatsbookTemplate
    private sbErrors: IErrorSummary
    private sbData: DerbyJson.IGame
    private warningData: IWarningData
    private penalties: { [playerId: string]: any[] }

    constructor(workbook: WorkBook, filename: string) {
        this.workbook = workbook
        this.sbErrors = JSON.parse(JSON.stringify(errorTemplate))
        this.sbFilename = filename
        this.penalties = {}
        this.parseFile()
    }

    get summary(): IStatsbookSummary {
        return {
            filename: this.sbFilename,
            version: this.sbVersion,
        }
    }

    get errors(): IErrorSummary {
        return JSON.parse(JSON.stringify(this.sbErrors))
    }

    get data(): DerbyJson.IGame {
        return JSON.parse(JSON.stringify(this.sbData))
    }

    get warnings(): IWarningData {
        return JSON.parse(JSON.stringify(this.warningData))
    }

    private parseFile(): void {
        this.sbVersion = this.getVersion()
        this.sbTemplate = this.getTemplate()
        this.warningData = {
            lost: [],
            badStarts: [],
            badContinues: [],
            jamsCalledInjury: [],
            noEntries: [],
            noExits: [],
            foulouts: [],
            expulsions: [],
            lineupThree: [],
        }

        this.sbData = {
            version: 'v0.3',
            type: 'game',
            metadata: {
                producer: 'Statsbook-Tool',
                date: new Date(),
            },
            date: null,
            time: null,
            venue: {
                name: null,
                city: null,
                state: null,
            },
            periods: {
                1: { jams: []},
                2: { jams: []},
            },
            teams: {
                home: null,
                away: null,
                officials: { persons: [] },
            },
        }

        this.getIGRF()
        this.getTeams()
        this.getScores()

    }

    private getVersion(): string {
        const sheet = this.workbook.Sheets['Read Me']
        const versionText = (sheet ? sheet.A3.v : WorkbookReader.defaultVersion)
        const versionRe = /(\d){4}/
        return versionRe.exec(versionText)[0]
    }

    private getTemplate(): IStatsbookTemplate {
        let result: IStatsbookTemplate

        if (this.sbVersion !== WorkbookReader.currentVersion) {
            this.sbErrors.warnings.oldStatsbookVersion.events.push(
                `This File: ${this.sbVersion}  Current Version: ${WorkbookReader.currentVersion} `,
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
                throw new Error(`Unable to Load Template for year ${this.sbVersion}`)
        }

        return result
    }

    private getIGRF() {
        const sheet = this.workbook.Sheets[this.sbTemplate.mainSheet]
        const venue = this.sbData.venue

        venue.name = this.getExpectedValue(sheet, 'venue.name', 'Venue Name')
        venue.city = this.getExpectedValue(sheet, 'venue.city', 'Venue City')
        venue.state = this.getExpectedValue(sheet, 'venue.state', 'Venue State')

        const excelDate = this.getExpectedValue(sheet, 'date', 'Date')
        const excelTime = this.getExpectedValue(sheet, 'time', 'Time')

        this.sbData.date = getJsDateFromExcel(excelDate)
        this.sbData.time = getJsTimeFromExcel(excelTime)
    }

    private getTeams() {
        teams.forEach((team) => {
            const teamTemplate = this.sbTemplate.teams[team]
            const sheet = this.workbook.Sheets[teamTemplate.sheetName]

            this.sbData.teams[team] = {
                leauge: cellVal(sheet, teamTemplate.league),
                name: cellVal(sheet, teamTemplate.name),
                color: cellVal(sheet, teamTemplate.color),
                persons: [],
            }

            const teamData = this.sbData.teams[team]

            if (!teamData.color) {
                this.sbErrors.warnings.missingData.events.push(
                    `Missing color for ${cap(team)} team.`,
                )
            }

            // Extract skater data
            const firstNameRC = utils.decode_cell(teamTemplate.firstName)
            const firstNumRC = utils.decode_cell(teamTemplate.firstNumber)

            range(0, teamTemplate.maxNum).forEach((i) => {
                const nameAddr = getAddressOfRow(i, firstNameRC)
                const numAddr = getAddressOfRow(i, firstNumRC)

                const skaterNumber = cellVal(sheet, numAddr)
                const skaterName = cellVal(sheet, nameAddr)

                if (!skaterNumber) {
                    return
                }

                teamData.persons.push({
                    name: skaterName,
                    number: skaterNumber,
                })

                this.penalties[`${team}:${skaterNumber}`] = []

            })
        })

        this.getOfficials()
    }

    private getOfficials() {

        const sheet = this.workbook.Sheets[this.sbTemplate.teams.officials.sheetName]
        const template = this.sbTemplate.teams.officials

        const cells: { [index: string]: CellAddress } = {
            firstName: utils.decode_cell(template.firstName),
            firstRole: utils.decode_cell(template.firstRole),
            firstLeague: utils.decode_cell(template.firstLeague),
            firstCert: utils.decode_cell(template.firstCert),
        }

        range(0, template.maxNum).forEach((idx) => {
            const nameAddr = getAddressOfRow(idx, cells.firstName)
            const roleAddr = getAddressOfRow(idx, cells.firstRole)
            const leagueAddr = getAddressOfRow(idx, cells.firstLeague)
            const certAddr = getAddressOfRow(idx, cells.firstCert)

            const name = cellVal(sheet, nameAddr)
            const role = cellVal(sheet, roleAddr)
            const league = cellVal(sheet, leagueAddr)
            const cert = cellVal(sheet, certAddr)

            if (role === undefined) { return }

            const person = {
                name,
                roles: [ role ],
                league,
                certifications: [],
            }

            if (cert !== undefined) {
                person.certifications.push({ level: cert })
            }

            this.sbData.teams.officials.persons.push(person)

        })
    }

    private getScores() {
        const sheet = this.workbook.Sheets[this.sbTemplate.score.sheetName]
        const scoreReader = new ScoreReader(this.sbData, this.sbTemplate, this.sbErrors, this.warningData)

        scoreReader.parseScoreSheet(sheet)
    }

    private getExpectedValue(sheet: WorkSheet, field: string, longName: string = null) {
        const address: string = get(this.sbTemplate, field)
        const value = cellVal(sheet, address)
        if (!value && longName) {
            this.sbErrors.warnings.missingData.events.push(longName)
        }

        return value
    }

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

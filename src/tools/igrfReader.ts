import { capitalize as cap, get, range } from 'lodash'
import { utils, WorkSheet } from 'xlsx'
import { CellAddressDict, cellsForRow, cellVal, getAddressOfRow, teams } from './utils'

export class IgrfReader {
    private sbData: DerbyJson.IGame
    private sbTemplate: IStatsbookTemplate
    private sbErrors: IErrorSummary

    constructor(sbData: DerbyJson.IGame,
                sbTemplate: IStatsbookTemplate,
                sbErrors: IErrorSummary) {

        this.sbData = sbData
        this.sbTemplate = sbTemplate
        this.sbErrors = sbErrors
    }

    public parseSheet(sheet: WorkSheet): void {
        this.gameDetails(sheet)
        this.teams(sheet)
        this.officials(sheet)
    }

    private gameDetails(sheet: WorkSheet): void {
        const venue = this.sbData.venue

        venue.name = this.getExpectedValue(sheet, 'venue.name', 'Venue Name')
        venue.city = this.getExpectedValue(sheet, 'venue.city', 'Venue City')
        venue.state = this.getExpectedValue(sheet, 'venue.state', 'Venue State')

        const excelDate = this.getExpectedValue(sheet, 'date', 'Date')
        const excelTime = this.getExpectedValue(sheet, 'time', 'Time')

        this.sbData.date = getJsDateFromExcel(excelDate)
        this.sbData.time = getJsTimeFromExcel(excelTime)
    }

    private teams(sheet: WorkSheet): void {
        teams.forEach((team) => {
            const teamTemplate = this.sbTemplate.teams[team]

            this.sbData.teams[team] = {
                league: cellVal(sheet, teamTemplate.league),
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
            })
        })

    }

    private officials(sheet: WorkSheet): void {
        const template = this.sbTemplate.teams.officials

        const firstRow: CellAddressDict = {
            firstname: utils.decode_cell(template.firstName),
            firstrole: utils.decode_cell(template.firstRole),
            firstleague: utils.decode_cell(template.firstLeague),
            firstcert: utils.decode_cell(template.firstCert),
        }

        range(0, template.maxNum).forEach((idx) => {
            const cells = cellsForRow(idx, firstRow)

            const name = cellVal(sheet, cells.name)
            const role = cellVal(sheet, cells.role)
            const league = cellVal(sheet, cells.league)
            const cert = cellVal(sheet, cells.cert)

            if (name === undefined || role === undefined) { return }

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

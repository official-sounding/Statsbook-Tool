import { WorkBook } from 'xlsx'
import template2017 from '../assets/2017statsbook.json'
import template2018 from '../assets/2018statsbook.json'
import errorTemplate from '../assets/sberrors.json'
import { BoxTripReader2018, BoxTripReader2019 } from './boxTripReader.js'
import { IgrfReader } from './igrfReader'
import { LineupReader } from './lineupReader.js'
import { PenaltyReader } from './penaltyReader'
import { ScoreReader } from './scoreReader'

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
    private boxTripReader: IBoxTripReader

    constructor(workbook: WorkBook, filename: string) {
        this.workbook = workbook
        this.sbErrors = JSON.parse(JSON.stringify(errorTemplate))
        this.sbFilename = filename
        this.parseFile()
    }

    get template(): IStatsbookTemplate {
        return JSON.parse(JSON.stringify(this.sbTemplate))
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
        this.getScores()
        this.getPenalties()
        this.getLineups()
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
                this.boxTripReader = new BoxTripReader2019()
                result = template2018
                break
            case '2018':
                this.boxTripReader = new BoxTripReader2018()
                result = template2018
                break
            case '2017':
                this.boxTripReader = new BoxTripReader2018()
                result = template2017
                break
            default:
                throw new Error(`Unable to Load Template for year ${this.sbVersion}`)
        }

        return result
    }

    private getIGRF() {
        const sheet = this.workbook.Sheets[this.sbTemplate.mainSheet]
        const reader = new IgrfReader(this.sbData, this.sbTemplate, this.sbErrors)

        reader.parseSheet(sheet)
    }

    private getScores() {
        const sheet = this.workbook.Sheets[this.sbTemplate.score.sheetName]
        const scoreReader = new ScoreReader(this.sbData, this.sbTemplate, this.sbErrors, this.warningData)

        scoreReader.parseSheet(sheet)
    }

    private getPenalties() {
        const sheet = this.workbook.Sheets[this.sbTemplate.penalties.sheetName]
        const penaltyReader = new PenaltyReader(this.sbData, this.sbTemplate, this.sbErrors, this.warningData)

        penaltyReader.parseSheet(sheet)
    }

    private getLineups() {
        const sheet = this.workbook.Sheets[this.sbTemplate.lineups.sheetName]
        const lineupReader = new LineupReader(this.sbData,
                                              this.sbTemplate,
                                              this.sbErrors,
                                              this.warningData,
                                              this.boxTripReader)

        lineupReader.parseSheet(sheet)
    }
}

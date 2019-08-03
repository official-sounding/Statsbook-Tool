
declare interface IStatsbookTemplate {
    "version": string,
    "mainSheet": string,
    "venue": {
        "name": string,
        "city": string,
        "state": string
    },
    "date": string,
    "time": string,
    "tournament": string,
    "host-league": string,
    "teams": {
        "home": {
            "sheetName": string,
            "league": string,
            "name": string,
            "color": string,
            "firstName": string,
            "firstNumber": string,
            "maxNum": number
        },
        "away": {
            "sheetName": string,
            "league": string,
            "name": string,
            "color": string,
            "firstName": string,
            "firstNumber": string,
            "maxNum": number
        },
        "officials":{
            "sheetName": string,
            "firstName": string,
            "firstRole": string,
            "firstLeague": string,
            "firstCert": string,
            "maxNum": number
                }
    },
    "score": {
        "sheetName": string,
        "maxJams": number,
        "1": {
            "home": {
                "firstJamNumber": string,
                "firstJammerNumber": string,
                "firstLost": string,
                "firstLead": string,
                "firstCall": string,
                "firstInj": string,
                "firstNp": string,
                "firstTrip": string,
                "lastTrip": string,
            },
            "away": {
                "firstJamNumber": string,
                "firstJammerNumber": string,
                "firstLost": string,
                "firstLead": string,
                "firstCall": string,
                "firstInj": string,
                "firstNp": string,
                "firstTrip": string,
                "lastTrip": string,                
            }
        },
        "2": {
            "home": {
                "firstJamNumber": string,
                "firstJammerNumber": string,
                "firstLost": string,
                "firstLead": string,
                "firstCall": string,
                "firstInj": string,
                "firstNp": string,
                "firstTrip": string,
                "lastTrip": string,
            },
            "away": {
                "firstJamNumber": string,
                "firstJammerNumber": string,
                "firstLost": string,
                "firstLead": string,
                "firstCall": string,
                "firstInj": string,
                "firstNp": string,
                "firstTrip": string,
                "lastTrip": string,                
            }
        }
    },
    "penalties": {
        "sheetName": string,
        "maxPenalties": number,
        "1": {
            "home": {
                "firstNumber": string,
                "firstPenalty": string,
                "firstJam": string,
                "firstFO": string,
                "firstFOJam": string,
                "benchExpCode": string,
                "benchExpJam": string,
                
            },
            "away": {
                "firstNumber": string,
                "firstPenalty": string,
                "firstJam": string,
                "firstFO": string,
                "firstFOJam": string, 
                "benchExpCode": string,
                "benchExpJam": string,            
            }
        },
        "2": {
            "home": {
                "firstNumber": string,
                "firstPenalty": string,
                "firstJam": string,
                "firstFO": string,
                "firstFOJam": string,
                "benchExpCode": string,
                "benchExpJam": string,
            },
            "away": {
                "firstNumber": string,
                "firstPenalty": string,
                "firstJam": string,
                "firstFO": string,
                "firstFOJam": string,
                "benchExpCode": string,
                "benchExpJam": string,
            }           
        }
    },
    "lineups": {
        "sheetName": string,
        "maxJams": number,
        "boxCodes": number,
        "1": {
            "home": {
                "firstJamNumber": string,
                "firstNoPivot": string,
                "firstJammer": string,
            },
            "away": {
                "firstJamNumber": string,
                "firstNoPivot": string,
                "firstJammer": string,                
            }
        },
        "2": {
            "home": {
                "firstJamNumber": string,
                "firstNoPivot": string,
                "firstJammer": string,
            },
            "away": {
                "firstJamNumber": string,
                "firstNoPivot": string,
                "firstJammer": string,                
            }            
        }
    },
    "clock": {
        "sheetName": string,
        "1": {
            "firstJamTime": string,
        },
        "2": {
            "firstJamTime": string,
        }
    }
}

declare interface IErrorDetails {
    "description": string,
    "events": string[],
    "long": string
}

declare type period = '1' | '2'
declare type team = 'home' | 'away'
declare type boxTripEvent = 'enter' | 'exit' | 'enterExit' | 'continue' | 'badContinue' | 'injury' | 'error'

declare interface ISimpleWarningDetails {
    period: period,
    team: team,
    jam: number
}
declare interface  IWarningDetails extends ISimpleWarningDetails {
    skater: string
}

declare interface IBoxTripReader {
    parseGlyph(glyph: string, team: team, skaterNumber: string): IBoxTrip
    stillInBox(team: team, skaterNumber: string): boolean
    removeFromBox(team: team, skaterNumber: string): boolean
    missingSkaters(team: team, skaterList: string[]): string[]

    badStartErrorKey: string
    badCompleteErrorKey: string
    badContinueErrorKey: string
    badBtwnJamErrorKey: string
    badBtwnJamCompleteErrorKey: string
}

declare interface IBoxTrip {
    eventType: boxTripEvent,
    errorKey?: string,
    note?: string,
    betweenJams?: boolean
}

declare interface IWarningData {
    lost: Array<IWarningDetails>,
    badStarts: Array<IWarningDetails>,
    noEntries: Array<IWarningDetails>,
    badContinues: Array<IWarningDetails>,
    noExits: Array<IWarningDetails>,
    foulouts: Array<IWarningDetails>,
    expulsions: Array<IWarningDetails>,
    jamsCalledInjury: Array<ISimpleWarningDetails>,
    lineupThree: Array<IWarningDetails>,
}

declare interface IErrorSummary {
    scores: { [key:string]: IErrorDetails },
    penalties: { [key:string]: IErrorDetails },
    lineups: { [key:string]: IErrorDetails },
    warnings: { [key:string]: IErrorDetails }
}

declare interface IStatsbookSummary {
    version: string,
    filename: string
}

declare interface ICrgSkater {
    id: string,
    flags: string,
    number: string,
    name: string
}

declare interface ICrgTeam {
    id: string,
    name: string,
    skaters: Array<ICrgSkater>
}
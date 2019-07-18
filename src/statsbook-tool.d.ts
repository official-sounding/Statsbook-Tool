
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
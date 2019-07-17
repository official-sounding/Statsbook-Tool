export interface IDerbyJsonData {
    date: Date
    time: string
    venue: IDerbyJsonVenue,
    teams: {
        home: IDerbyJsonTeam,
        away: IDerbyJsonTeam,
        officials: IDerbyJsonOfficialsTeam
    }
    periods: {
        "1": IDerbyJsonPeriod,
        "2": IDerbyJsonPeriod
    }
}

export interface IDerbyJsonPeriod {
    jams: IDerbyJsonJam[]
}

export interface IDerbyJsonJam {
    events: IDerbyJsonEvent[],
}

export interface IDerbyJsonEvent {
    event: string,
    skater: IDerbyJsonSkaterRef,
    position: string,
}

export interface IDerbyJsonSkaterRef {

}

export interface IDerbyJsonVenue {
    name: string
    city: string
    state: string

}

export interface IDerbyJsonTeam {
    league: string,
    name: string,
    color: string,
    persons: IDerbyJsonPerson[]
}

export interface IDerbyJsonPerson {
    name: string,
    number: string
}

export interface IDerbyJsonOfficialsTeam {
    persons: IDerbyJsonOfficialPerson[]
}

export interface IDerbyJsonOfficialPerson {
    name: string,
    league: string,
    roles: string[],
    certifications: string[]
}
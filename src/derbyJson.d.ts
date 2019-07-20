declare namespace DerbyJson {
    interface IGame {
        version: string;
        /**
         * Additional information about the file that isn't relevant to the game data itself
         */
        metadata: {
          producer: string;
          date: Date;
          author?: string;
          comments?: string;
        };
        type: "game";
        date: Date
        time: string
        venue: IVenue,
        teams: {
            home: ITeam,
            away: ITeam,
            officials: IOfficialsTeam
        }
        periods: {
            "1": IPeriod,
            "2": IPeriod
        }
    }

    interface IPeriod {
        jams: IJam[]
    }

    interface IJam {
        number: number,
        events: IEvent[],
    }

    interface IEvent {
        event: string,
        skater: ISkaterRef,
        position?: string,
        number?: number,
        score?: number,
        team?: string,
        completed?: boolean
    }

    interface ISkaterRef {

    }

    interface IVenue {
        name: string
        city: string
        state: string

    }

    interface ITeam {
        name: string;
        abbreviation?: string;
        league?: string;
        level?: "All Star" | "B" | "C" | "Rec" | "Officials" | "Home" | "Adhoc";
        date?: string;
        color?: {
          [k: string]: any;
        };
        logo?: Logo;
        persons?: IPerson[];
    }

    export interface Logo {
        url?: string;
        small?: string;
        medium?: string;
        large?: string;
        url_light?: string;
        small_light?: string;
        medium_light?: string;
        large_light?: string;
        url_dark?: string;
        small_dark?: string;
        medium_dark?: string;
        large_dark?: string;
        url_grayscale?: string;
        small_grayscale?: string;
        medium_grayscale?: string;
        large_grayscale?: string;
    }

    interface IPerson {
        name: string,
        number: string
    }

    interface IOfficialsTeam {
        persons: IOfficialPerson[]
    }

    interface IOfficialPerson {
        name: string,
        league: string,
        roles: string[],
        certifications: string[]
    }

}
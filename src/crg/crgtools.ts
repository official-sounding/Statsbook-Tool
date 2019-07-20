import uuid from 'uuid/v5'
import { teams } from '../tools/utils'

// use this guid as the namespace for the skater id
const appNamespace = 'f291bfb9-74bf-4c5b-af33-902c19a74bea'

function generateSkaterId(teamId: string, skaterName: string): string {
    const name = `${teamId}-${skaterName}`
    return uuid(name, appNamespace)
}

export function extractTeamsFromSBData(sbData: DerbyJson.IGame): ICrgTeam[] {
    return teams.map((t) => {
        let teamName: string

        if (sbData.teams[t].league) {
            teamName = `${sbData.teams[t].league} ${sbData.teams[t].name}`
        } else {
            teamName = sbData.teams[t].name
        }

        const team: ICrgTeam = {
            id: teamName,
            name: teamName,
            skaters: [],
        }

        team.skaters = sbData.teams[t].persons.map((person: DerbyJson.IPerson) => {
            const skaterName = person.name
            return {
                flags: '',
                id: generateSkaterId(teamName, skaterName),
                name: skaterName,
                number: person.number,
            }
        })

        return team
    })
}

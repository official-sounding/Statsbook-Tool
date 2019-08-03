export class BoxTripReader2018 implements IBoxTripReader {

    public badStartErrorKey = 'slashNoPenalty'
    public badCompleteErrorKey = 'xNoPenalty'
    public badContinueErrorKey = 'iNotInBox'
    public badBtwnJamErrorKey = 'sNoPenalty'
    public badBtwnJamCompleteErrorKey = 'sSlashNoPenalty'

    private box: { [t in team]: Set<string> } = { home: new Set<string>(), away: new Set<string>() }

    public stillInBox(team: team, skaterNumber: string): boolean {
        return this.box[team].has(skaterNumber)
    }

    public missingSkaters(team: team, skaterList: string[]): string[] {
        return Array.from(this.box[team])
                    .filter((box) => !skaterList.includes(box))
    }

    public removeFromBox(team: team, skaterNumber: string): boolean {
        return this.box[team].delete(skaterNumber)
    }

    // Possible codes - /, X, S, $, I or |, 3
    // Possible events - enter box, exit box, injury
    // / - Enter box
    // X - Test to see if skater is IN box
    //      Yes: exit box, No: enter box, exit box
    // S - Enter box, note: sat between jams
    // $ - Enter box, exit box, note: sat between jams
    // I or | - no event, error checking only
    // 3 - Injury object, verify not already present from score tab
    public parseGlyph(glyph: string, team: team, skaterNumber: string): IBoxTrip {
        let result: IBoxTrip

        const lower = (glyph || '').toLowerCase()

        switch (lower) {
        case '/':
            result = ({ eventType: 'enter' })
            this.box[team].add(skaterNumber)
            break
        case 'x':
            if (!this.box[team].has(skaterNumber)) {
                result = ({ eventType: 'enterExit' })
            } else {
                this.box[team].delete(skaterNumber)
                result = ({ eventType: 'exit' })
            }
            break
        case 's':
            if (!this.box[team].add(skaterNumber)) {
                result = ({ eventType: 'error', errorKey: 'startsWhileThere' })
            } else {
                result = ({ eventType: 'enter', betweenJams: true, note: 'Sat between jams' })
            }
            break
        case '$':
            if (this.box[team].has(skaterNumber)) {
                result = ({ eventType: 'error', errorKey: 'startsWhileThere' })
            } else {
                result = ({ eventType: 'enterExit', betweenJams: true, note: 'Sat between jams' })
            }
            break
        case '|':
        case 'i':
            if (!this.box[team].has(skaterNumber)) {
                result = ({ eventType: 'badContinue' })
            } else {
                result = ({ eventType: 'continue' })
            }
            break
        case '3':
            result = ({ eventType: 'injury' })
            break
        case 'ᚾ':
            result = ({ eventType: 'error', errorKey: 'runeUsed' })
            break
        default:
            result = ({ eventType: 'error', errorKey: 'badLineupCode', note: `Code: ${glyph}` })
            break
        }

        return result
    }
}

export class BoxTripReader2019 implements IBoxTripReader {

    public badStartErrorKey = 'dashNoPenalty'
    public badCompleteErrorKey = 'plusNoPenalty'
    public badContinueErrorKey = ''
    public badBtwnJamErrorKey = 'sNoPenalty'
    public badBtwnJamCompleteErrorKey = 'sSlashNoPenalty'

    private box: { [t in team]: Set<string> } = { home: new Set<string>(), away: new Set<string>() }

    public stillInBox(team: team, skaterNumber: string): boolean {
        return this.box[team].has(skaterNumber)
    }

    public missingSkaters(team: team, skaterList: string[]): string[] {
        return Array.from(this.box[team])
                    .filter((box) => !skaterList.includes(box))
    }

    public removeFromBox(team: team, skaterNumber: string): boolean {
        return this.box[team].delete(skaterNumber)
    }

        // Possible codes: -, +, S, $, 3

        // - - Enter box
        // + - Enter and exit box
        // S - Sat between jams or continued
        // $ - Sat between jams or continued with exit
        // 3 - Injury object, verify not already present from score tab

    public parseGlyph(glyph: string, team: team, skaterNumber: string): IBoxTrip {
        let result: IBoxTrip

        const lower = (glyph || '').toLowerCase()

        switch (lower) {
            case '-':
                this.box[team].add(skaterNumber)
                result = ({ eventType: 'enter' })
                break
            case '+':
                // TODO: error check if already in box
                result = ({ eventType: 'enterExit' })
                break
            case 's':
                if (this.box[team].add(skaterNumber)) {
                    result = ({ eventType: 'enter', betweenJams: true, note: 'Sat between jams' })
                }
                break
            case '$':
                if (this.box[team].has(skaterNumber)) {
                    result = ({ eventType: 'exit' })
                    this.box[team].delete(skaterNumber)
                } else {
                    result = ({ eventType: 'enterExit', betweenJams: true, note: 'Sat between jams' })
                }
                break
            case '3':
                result = ({ eventType: 'injury' })
                break
            default:
                result = ({ eventType: 'error', errorKey: 'badLineupCode', note: `Code: ${glyph}` })
                break
        }

        return result
    }
}

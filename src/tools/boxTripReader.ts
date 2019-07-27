export class BoxTripReader2018 implements IBoxTripReader {
    private box: { [t in team]: Set<string> } = { home: new Set<string>(), away: new Set<string>() }

    public parseGlyph(glyph: string, team: team, skaterNumber: string): IBoxTrip[] {
        const result: IBoxTrip[] = []

        const lower = (glyph || '').toLowerCase()

        switch (lower) {
        case '/':
            result.push({ eventType: 'enter' })
            this.box[team].add(skaterNumber)
            break
        case 'x':
            if (!this.box[team].has(skaterNumber)) {
                result.push({ eventType: 'enter' })
            } else {
                this.box[team].delete(skaterNumber)
            }
            result.push({ eventType: 'exit' })
            break
        case 's':
            if (!this.box[team].add(skaterNumber)) {
                result.push({ eventType: 'error' })
            }
        }

        return result
    }
}
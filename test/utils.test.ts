import * as utils from '../dist/src/tools/utils'

describe('Utilities', () => {

    it('has the correct teams', () => {
        expect(utils.teams.length).toBe(2)
        expect(utils.teams[0]).toEqual('home')
        expect(utils.teams[1]).toEqual('away')
    })
    it('has the correct periods', () => {
        expect(utils.periods.length).toBe(2)
        expect(utils.periods[0]).toBe('1')
        expect(utils.periods[1]).toBe('2')
    })

    it('forEachPeriodTeam', () => {
        let i = 0
        const values: any = []
        utils.forEachPeriodTeam((period, team) => {
            i++
            values.push({ period, team })
        })

        expect(i).toBe(4)
        expect(values[0]).toEqual({ period: '1', team: 'home' })
        expect(values[3]).toEqual({ period: '2', team: 'away' })
    })

    describe('cellVal', () => {
        it('extracts a value with cell', () => {
            const sheet = { A1: { v: 123 } }
            const value = utils.cellVal(sheet, 'A1')

            expect(value).toBe(123)
        })
        it('returns undefined on a cell without value', () => {
            const sheet = { A1: { v: null } }
            const value = utils.cellVal(sheet, 'A1')
            const value2 = utils.cellVal(sheet, 'A2')

            expect(value).toBe(undefined)
            expect(value2).toBe(undefined)
        })
    })
})

import { CellAddress, utils, WorkSheet } from 'xlsx';

// tslint:disable-next-line: interface-name
export interface CellAddressDict { [key: string]: CellAddress }

export const teams = ['home', 'away'];
export const periods = ['1', '2'];

export function cellVal(sheet: WorkSheet, address: string) {
    // Given a worksheet and a cell address, return the value
    // in the cell if present, and undefined if not.
    if (sheet[address] && sheet[address].v) {
        return sheet[address].v
    } else {
        return undefined
    }
}

export function forEachPeriodTeam(cb: (period: string, team: string) => void): void {
    periods.forEach((period) => teams.forEach((team) => cb(period, team)));
}

export function cellsForRow(idx: number, firstCells: CellAddressDict): { [key: string]: string } {
    const result: { [key: string]: string } = {};
    Object.keys(firstCells).reduce((prev, curr) => {
        const key = curr.replace('first', '');
        prev[key] = getAddressOfRow(idx, firstCells[curr])
        return prev
    }, result)

    return result
}

export function getAddressOfRow(idx: number, firstCell: CellAddress): string {
    const rcAddr = Object.assign({}, firstCell);
    rcAddr.r = rcAddr.r + idx;

    return utils.encode_cell(rcAddr);
}

// tslint:disable-next-line: max-line-length
export function initializeFirstRow(template: IStatsbookTemplate, tab: string, team: string, period: string, fields: string[]): CellAddressDict {
    const result: CellAddressDict = {};

    fields.reduce((prev, curr) => {
        prev[curr] = utils.decode_cell(template[tab][team][period][curr])
        return prev;
    }, result);

    return result;
}

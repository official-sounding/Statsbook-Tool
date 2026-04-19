const XLSX = require('xlsx');
const request = require('request')
const fs = require('fs')

const template2018 = require('../assets/2018statsbook.json')
const template2017 = require('../assets/2017statsbook.json')
const template2023jrda = require('../assets/2023jrda.json')

const errorManager = require('./errorManager');

const versionRe = /(\d){4}/;
const currentVersion = '2024'
const currentJRDAVersion = '2024jrda'
const defaultVersion = '2018'

const teams = ['home', 'away'];
const periods = ['1', '2'];

let initCells = (team, period, tab, props) => {
    // Given a team, period, SB section, and list of properties,
    // return an object of addresses for those properties.
    // Team should be 'home' or 'away'
    let cells = {}

    for (let i in props){
        cells[props[i]] = XLSX.utils.decode_cell(
            fileManager.template[tab][period][team][props[i]])
    }

    return cells
}


class SheetManager {
    template;
    props;

    constructor(workbook, sheetName, template, props) {
        this.sheet = workbook[sheetName];
        this.template = template;
        this.props = props;
    }

    cellVal(address) {
        // Given a worksheet and a cell address, return the value
        // in the cell if present, and undefined if not.
        if (this.sheet[address] && this.sheet[address].v){
            return this.sheet[address].v
        } else {
            return undefined
        }
    }

    rawVal(addr) {
        return this.sheet[addr];
    }

    relevantCells(team, period) {
        // Given a team, period, SB section, and list of properties,
        // return an object of addresses for those properties.
        // Team should be 'home' or 'away'
        return props.reduce((prev, next) => {
            prev[next] = XLSX.utils.decode_cell(template[period][team][props[i]]);
            return prev;
        }, {});
    }


}

class FileManager {
    workbook = {};
    template = {};
    sbVersion = '';
    sbFilename = '';
    initialized = false;
    mode = '';

    get fileForExport() {
        if(this.mode === 'file') {
            return this.sbFilename.split('.')[0];
        }

        return 'export';
     }

    loadFromGoogleSheet(sheetUrl) {
        return new Promise((resolve, reject) => {
            this.mode = 'gsheet';
            this.sbFilename = sheetUrl;
            request.get(sheetUrl, { encoding: null }, (err, res, data) => {
                if (err || res.statusCode != 200) {
                    reject(`Unable to load file. Error: ${err}`);
                }
                const buf = Buffer.from(data)
                this.workbook = XLSX.read(buf)
                this.#setVersion();
                this.initialized = true;
                resolve();
            })
        })
        
    }

    loadFromFileSystem(sbFile) {
        this.mode = 'file';
        this.sbFilename = sbFile.name
        let data = fs.readFileSync(sbFile.path)
        data = new Uint8Array(data)
        this.workbook = XLSX.read(data, {type: 'array'})
        this.#setVersion();
        this.initialized = true;
    }

    #setVersion() {
        let sheet = this.workbook.Sheets['Read Me']
        let versionText = (sheet ? sheet['A3'].v : defaultVersion)
        
        this.sbVersion = versionRe.exec(versionText)[0];

        if (versionText.toLowerCase().includes('jrda')) {
            this.sbVersion = `${this.sbVersion}jrda`;
        }

        // Warning check: outdated statsbook version
        // Note that this will ALSO fire if the statsbook version is NEWER.
        if (this.sbVersion < currentVersion) {
            errorManager.addWarning('oldStatsbookVersion', `This File: ${this.sbVersion}  Current Preferred Version: ${currentVersion}`);
        }

        // Check for NEWER statsbook version
        if (this.sbVersion > currentVersion) {
            if(this.sbVersion.includes('jrda')) {
                this.sbVersion = currentJRDAVersion;
            } else {
                this.sbVersion = currentVersion;
            }
        }

        switch (sbVersion) {
            case '2024jrda':
            case '2023jrda':
                this.template = template2023jrda
                break
            case '2024':
            case '2019':
            case '2018':
                this.template = template2018
                break
            case '2017':
                this.template = template2017
                break
            default:
                this.template = {}
        }
    }

    getSheet(sheetName) {
        return new SheetManager(this.workbook, sheetName);
    }

    getLineupSheet() {
        const template = this.template.lineups;
        const props = ['firstJamNumber','firstNoPivot','firstJammer'];

        const sheet = new SheetManager(this.workbook, template, props);


        return {
            template,
            sheet,
        }

    }

    getScoreSheet() {
        const template = this.template.score;
        const props = ['firstJamNumber','firstJammerNumber','firstLost','firstLead',
        'firstCall','firstInj','firstNp','firstTrip','lastTrip'];

        const sheet = new SheetManager(this.workbook, template, props);


        return {
            template,
            sheet,
        }
    }

    getPenaltiesSheet() {
        const template = this.template.penalties;
        const props = ['firstNumber','firstPenalty','firstJam',
            'firstFO','firstFOJam','benchExpCode','benchExpJam'];

        const sheet = new SheetManager(this.workbook, template, props);


        return {
            template,
            sheet,
        }
    }
}

module.exports = new FileManager();
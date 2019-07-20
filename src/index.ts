/* tslint:disable:triple-equals max-line-length forin one-variable-per-declaration */
import { ipcRenderer as ipc, remote } from 'electron'
import _ from 'lodash'
import { capitalize as cap } from 'lodash'
import moment from 'moment'
import mousetrap from 'mousetrap'
import { CellAddress, read, utils, WorkBook } from 'xlsx'
import { download } from './tools/download'
import { WorkbookReader } from './tools/workbookReader'
const { Menu, MenuItem } = remote

import { extractTeamsFromSBData } from './crg/crgtools'
import { exportJson as exportJsonRoster } from './crg/exportJson'
import { exportXml } from './crg/exportXml'

interface IHTMLInputEvent extends Event {
    target: HTMLInputElement & EventTarget
}

// Page Elements
const holder = document.getElementById('drag-file')
const fileSelect = document.getElementById('file-select')
const fileInfoBox = document.getElementById('file-info-box')
const outBox = document.getElementById('output-box')
const newVersionWarningBox = document.getElementById('newVersionWarning')
let refreshButton: HTMLElement

const menu = new Menu()
menu.append( new MenuItem( { role: 'copy'} ))
menu.append( new MenuItem( { role: 'selectall'} ))

window.addEventListener('contextmenu', (e) => {
    e.preventDefault()
    menu.popup( { window: remote.getCurrentWindow() })
}, false)

const rABS = true // read XLSX files as binary strings vs. array buffers
// Globals

let sbData: any = {},  // derbyJSON formatted statsbook data
    sbFile: File,
    sbReader: WorkbookReader,
    sbErrors: IErrorSummary,
    sbSummary: IStatsbookSummary,
    penalties: any = {},
    starPasses: any[] | Array<{ period: number; jam: number; }> = [],
    sbFilename = '',
    warningData: any = {}

const sbTemplate: IStatsbookTemplate = null
const teamList = ['home', 'away']
const anSP = /^sp\*?$/i
const mySP = /^sp$/i

// Check for new version
ipc.on('do-version-check', (event: any, version: any) => {
    fetch('https://api.github.com/repos/AdamSmasherDerby/Statsbook-Tool/tags')
        .then((result) =>  result.json())
        .then((data) => {
            const latestVersion = data[0]
            const currentVersion = `v${version}`
            if (latestVersion !== currentVersion) {
                newVersionWarningBox.innerHTML = `New version available: ${latestVersion} (Current Version: ${version})</BR>` +
                '<A HREF="https://github.com/AdamSmasherDerby/Statsbook-Tool/releases/" target="_system">Download Here</a>'
            }
        })
        .catch((reason) => console.error(`Cannot check for latest version ${reason}`))
})

fileSelect.onchange = (e?: IHTMLInputEvent) => {
    // Fires if a file is selected by clicking "select file."
    if (e.target.value == undefined) {
        return false
    }
    e.preventDefault()
    e.stopPropagation()

    if (e.target.files.length > 1) {
        fileInfoBox.innerHTML = 'Error: Multiple Files Selected.'
        return false
    }

    sbFile = e.target.files[0]

    makeReader(sbFile)
    return false
}

holder.ondrop = (e) => {
    // Fires if a file is dropped into the box
    holder.classList.remove('box__ondragover')
    e.preventDefault()
    e.stopPropagation()

    if (e.dataTransfer.files.length > 1) {
        fileInfoBox.innerHTML = 'Error: Multiple Files Selected.'
        return false
    }

    sbFile = e.dataTransfer.files[0]

    makeReader(sbFile)
    return false
}

const makeReader = (file: File) => {
    // Create reader object and load statsbook file
    const reader = new FileReader()
    sbFilename = file.name

    reader.onload = (e: any) => {
        readSbData(e.target.result, file.name)
    }

    // Actually load the file
    if (rABS) {
        reader.readAsBinaryString(file)
    } else {
        reader.readAsArrayBuffer(file)
    }
}

const readSbData = (data, filename) => {
    // Read in the statsbook data for an event e
    let readType: 'binary' | 'array' = 'binary'
    if (!rABS) {
        data = new Uint8Array(data)
        readType = 'array'
    }
    const workbook = read(data, { type: readType })

    // Reinitialize globals
    penalties = {}
    starPasses = []
    warningData = {
        badStarts: [],
        noEntries: [],
        badContinues: [],
        noExits: [],
        foulouts: [],
        expulsions: [],
        lost: [],
        jamsCalledInjury: [],
        lineupThree: [],
    }

    sbReader = new WorkbookReader(workbook, filename)

    sbSummary = sbReader.summary
    sbErrors = sbReader.errors
    sbData = sbReader.data

    updateFileInfo()

    sbData.periods = {1: {jams: []}, 2: {jams: []}}
    readScores(workbook)
    readPenalties(workbook)
    readLineups(workbook)
    errorCheck()
    warningCheck()

    // Display Error List
    if (outBox.lastElementChild) {
        outBox.removeChild(outBox.lastElementChild)
    }
    outBox.appendChild(sbErrorsToTable())

    // Initialize Tooltips
    $(document).ready(() => {
        $('[data-toggle="tooltip"]').tooltip()
    })

    // Update UI
    ipc.send('enable-menu-items')
    createRefreshButton()
}

function updateFileInfo(): void {
    // Update the "File Information Box"
    // Update File Information Box

    fileInfoBox.innerHTML = `<strong>Filename:</strong>  ${sbSummary.filename}<br>`
    fileInfoBox.innerHTML += `<strong>SB Version:</strong> ${sbSummary.version}<br>`
    fileInfoBox.innerHTML += `<strong>Game Date:</strong> ${moment(sbData.date).format('MMMM DD, YYYY')}<br>`
    fileInfoBox.innerHTML += `<strong>Team 1:</strong> ${sbData.teams.home.league} ${sbData.teams.home.name}<br>`
    fileInfoBox.innerHTML += `<strong>Team 2:</strong> ${sbData.teams.away.league} ${sbData.teams.away.name}<br>`
    fileInfoBox.innerHTML += `<strong>File Read:</strong> ${moment().format('HH:mm:ss MMM DD, YYYY')} `
}

const createRefreshButton = () => {

    fileInfoBox.innerHTML += '<button id="refresh" type="button" class="btn btn-secondary btn-sm">Refresh (f5)</button>'
    refreshButton = document.getElementById('refresh')

    refreshButton.onclick = () => {
        makeReader(sbFile)
    }

    mousetrap.bind('f5', () => {
        makeReader(sbFile)
    })
}

const readScores = (workbook: WorkBook) => {
// Given a workbook, extract the information from the score tab

    let cells: any = {},
    jamNumber: { v: string; },
    skater: any = {}

    const maxJams = sbTemplate.score.maxJams,
        sheet = workbook.Sheets[sbTemplate.score.sheetName],
        jamAddress: CellAddress = { c: null, r: null },
        jammerAddress: CellAddress = { c: null, r: null },
        tripAddress: CellAddress = { c: null, r: null },
        lostAddress: CellAddress = { c: null, r: null },
        leadAddress: CellAddress = { c: null, r: null },
        callAddress: CellAddress = { c: null, r: null },
        injAddress: CellAddress = { c: null, r: null },
        npAddress: CellAddress = { c: null, r: null }

    const props = ['firstJamNumber', 'firstJammerNumber', 'firstLost', 'firstLead',
        'firstCall', 'firstInj', 'firstNp', 'firstTrip', 'lastTrip']
    const tab = 'score'
    const npRe = /(\d)\+NP/
    const ippRe = /(\d)\+(\d)/
    const jamNoRe = /^(\d+|SP|SP\*)$/i

    for (let period = 1; period < 3; period ++) {
    // For each period, import data

        // Add a period object with a jams array
        const pstring = period.toString()

        for (const i in teamList) {
        // For each team

            // Setup variables.  Jam is 0 indexed (1 less than jam nubmer).  Trip is 1 indexed.
            const team = teamList[i]
            let jam = 0
            let trip = 1
            let starPass = false

            // Get an array of starting points for each type of info
            cells = initCells(team, pstring, tab, props)
            const maxTrips = cells.lastTrip.c - cells.firstTrip.c
            jamAddress.c = cells.firstJamNumber.c
            jammerAddress.c = cells.firstJammerNumber.c
            tripAddress.c = cells.firstTrip.c
            lostAddress.c = cells.firstLost.c
            leadAddress.c = cells.firstLead.c
            callAddress.c = cells.firstCall.c
            injAddress.c = cells.firstInj.c
            npAddress.c = cells.firstNp.c

            for (let l = 0; l < maxJams; l++) {
            // For each line in the scoresheet, import data.

                let blankTrip = false
                let isLost = false
                let isLead = false

                // increment addresses
                jamAddress.r = cells.firstJamNumber.r + l
                jammerAddress.r = cells.firstJammerNumber.r + l
                tripAddress.r = cells.firstTrip.r + l
                lostAddress.r = cells.firstLost.r + l
                leadAddress.r = cells.firstLead.r + l
                callAddress.r = cells.firstCall.r + l
                injAddress.r = cells.firstInj.r + l
                npAddress.r = cells.firstNp.r + l

                // determine current jam number
                jamNumber = sheet[utils.encode_cell(jamAddress)]

                // if we're out of jams, stop
                if (
                    _.get(jamNumber, 'v') == undefined ||
                    /^\s+$/.test(jamNumber.v)
                ) { break }

                // Test for invalid jam number, throw error and stop
                if (!jamNoRe.test(_.trim(jamNumber.v))) {
                    throw new Error(`Invalid Jam Number: ${jamNumber.v}`)
                }

                // handle star passes
                if (anSP.test(jamNumber.v)) {
                    starPass = true
                    if (mySP.test(jamNumber.v)) {
                        sbData.periods[pstring].jams[jam - 1].events.push(
                            {
                                event: 'star pass',
                                skater,
                            },
                        )
                    }
                    starPasses.push({period, jam})
                } else {
                    // Not a star pass

                    // Error check - is this jam number out of sequence?
                    if (parseInt(jamNumber.v) != (jam + 1)) {
                        sbErrors.scores.badJamNumber.events.push(
                            `Team: ${cap(team)}, Period: ${pstring}, Jam: ${parseInt(jamNumber.v)}`,
                        )

                        // Add jam objects for missing jams.
                        for (let j = jam + 1; j < parseInt(jamNumber.v); j++) {
                            sbData.periods[pstring].jams[j - 1] = {number: j, events: []}
                        }
                    }

                    // Update the jam, reset the trip
                    jam = parseInt(jamNumber.v)
                    trip = 1
                    starPass = false
                }

                // If there isn't currently an numbered object for this jam, create it
                // Note that while the "number" field is one indexed, the jams array itself is zero indexed
                if (!sbData.periods[pstring].jams.find((o) => o.number === jam)) {
                    sbData.periods[pstring].jams[jam - 1] = {number: jam, events: []}
                }

                // Process trips.
                // Add a "pass" object for each trip, including initial passes
                // (note that even incomplete initial passes get "pass" events.)
                let skaterNum = ' '

                // Check for no initial pass box checked
                const np = sheet[utils.encode_cell(npAddress)]
                const initialCompleted = (_.get(np, 'v') != undefined ? 'no' : 'yes')

                if (sheet[utils.encode_cell(jammerAddress)] != undefined) {
                    skaterNum = sheet[utils.encode_cell(jammerAddress)].v
                }

                // ERROR CHECK: Skater on score sheet not on the IGRF
                if (skaterNum != ' ' &&
                    sbData.teams[team].persons.findIndex((x) => x.number == skaterNum) == -1) {
                    // This SHOULD be caught by conditional formatting in Excel, but there
                    // are reports of that breaking sometimes.
                    sbErrors.scores.scoresNotOnIGRF.events.push(
                        `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}, Skater: ${skaterNum} `,
                    )
                }

                if (!starPass) {
                // If this line is not a star pass, create an intital pass object

                    skater = team + ':' + skaterNum
                    sbData.periods[period].jams[jam - 1].events.push(
                        {
                            event: 'pass',
                            number: 1,
                            score: '',
                            skater,
                            team,
                            completed: initialCompleted,
                        },
                    )
                } else if (mySP.test(jamNumber.v)) {
                    // If THIS team has a star pass, use the skater number from the sheet

                    skater = team + ':' + skaterNum

                    // If this is still the initial trip, add another initial pass object.
                    if (trip == 1) {
                        sbData.periods[period].jams[jam - 1].events.push(
                            {
                                event: 'pass',
                                number: 1,
                                score: '',
                                skater,
                                team,
                                completed: initialCompleted,
                            },
                        )
                    }

                } else {
                    // Final case - jam number is SP*.
                    if (skaterNum != ' ') {
                        sbErrors.scores.spStarWithJammer.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}`,
                        )
                    }
                }

                // Check for subsequent trips, and add additional pass objects
                for (let t = 2; t < maxTrips + 2; t++) {
                    tripAddress.c = cells.firstTrip.c + t - 2
                    const tripScore = sheet[utils.encode_cell(tripAddress)]

                    if (tripScore == undefined) {

                        // ERROR CHECK - no trip score, initial pass completed
                        if (initialCompleted == 'yes' && t == 2 && !starPass) {
                            const nextJamNumber = sheet[utils.encode_cell({
                                r: jamAddress.r + 1, c: jamAddress.c})]
                            if (_.get(nextJamNumber, 'v') == 'SP') {
                                sbErrors.warnings.SPNoPointsNoNI.events.push(
                                    `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}, Jammer: ${skaterNum}`,
                                )
                            } else {
                                sbErrors.scores.noPointsNoNI.events.push(
                                    `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}, Jammer: ${skaterNum}`,
                                )
                            }
                        }

                        // Go on to next cell
                        blankTrip = true
                        continue
                    }

                    // Error check - points entered for a trip that's already been completed.
                    if (t <= trip) {
                        sbErrors.scores.spPointsBothJammers.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}`,
                        )
                    }

                    // Error check - skipped column in a non star pass line
                    if (blankTrip && !starPass) {
                        blankTrip = false
                        sbErrors.scores.blankTrip.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}`,
                        )
                    }

                    let reResult = []
                    let ippResult = []
                    let points = 0

                    if ((reResult = npRe.exec(tripScore.v))) {
                        // If score is x + NP, extract score and update initial trip
                        points = reResult[1]
                        sbData.periods[period].jams[jam - 1].events.find(
                            (x) => x.event == 'pass' && x.number == 1 && x.skater == skater,
                        ).score = points
                    } else if (tripScore.f != undefined && (ippResult = ippRe.exec(tripScore.f))) {
                        // If score is x + x, extract scores and add points to prior AND current trip
                        if (!starPass) {trip++}
                        sbData.periods[period].jams[jam - 1].events.find(
                            (x) => x.event == 'pass' && x.number == 1 && x.skater == skater,
                        ).score = ippResult[1]
                        sbData.periods[period].jams[jam - 1].events.push(
                            {
                                event: 'pass',
                                number: t,
                                score: ippResult[2],
                                skater,
                                team,
                            },
                        )
                    } else {
                        // Normal scoring trip
                        if (!starPass) {trip++}
                        points = tripScore.v
                        sbData.periods[period].jams[jam - 1].events.push(
                            {
                                event: 'pass',
                                number: t,
                                score: points,
                                skater,
                                team,
                            },
                        )
                    }

                    // ERROR CHECK: No Initial box checked with points given.
                    if (initialCompleted == 'no' && !reResult) {
                        sbErrors.scores.npPoints.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}, Jammer: ${skaterNum} `,
                        )
                    }

                }
                // Lost Lead
                const lost = sheet[utils.encode_cell(lostAddress)]
                if (_.get(lost, 'v') != undefined) {
                    isLost = true
                    sbData.periods[period].jams[jam - 1].events.push(
                        {
                            event: 'lost',
                            skater,
                        },
                    )
                    warningData.lost.push(
                        {
                            skater,
                            team,
                            period,
                            jam,
                        },
                    )
                }
                // Lead
                const lead = sheet[utils.encode_cell(leadAddress)]
                if (_.get(lead, 'v') != undefined) {
                    isLead = true
                    sbData.periods[period].jams[jam - 1].events.push(
                        {
                            event: 'lead',
                            skater,
                        },
                    )
                }
                // Call
                const call = sheet[utils.encode_cell(callAddress)]
                if (_.get(call, 'v') != undefined) {
                    sbData.periods[period].jams[jam - 1].events.push(
                        {
                            event: 'call',
                            skater,
                        },
                    )
                }
                // Injury
                const inj = sheet[utils.encode_cell(injAddress)]
                if (_.get(inj, 'v') != undefined) {

                    warningData.jamsCalledInjury.push(
                        {
                            team,
                            period,
                            jam,
                        },
                    )
                }

                // Error check - SP and lead without lost
                if (mySP.test(jamNumber.v) && isLead && !isLost) {
                    sbErrors.scores.spLeadNoLost.events.push(
                        `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}`,
                    )
                }
            }

        }
        // End of period - check for cross team errors and process injuries

        for (const j in sbData.periods[period].jams) {
            // For each jam in the period
            const jam = parseInt(j) + 1

            const numLead = sbData.periods[period].jams[j].events.filter(
                (x) => x.event == 'lead',
            ).length

            // ERROR CHECK: Lead box checked more than once in the same jam
            if (numLead >= 2) {
                sbErrors.scores.tooManyLead.events.push(
                    `Period: ${period}, Jam: ${jam}`,
                )
            }

            // ERROR CHECK: Call box checked for both jammers in same jam
            if (sbData.periods[period].jams[j].events.filter(
                (x) => x.event == 'call',
            ).length >= 2) {
                sbErrors.scores.tooManyCall.events.push(
                    `Period: ${period}, Jam: ${jam}`,
                )
            }

            // Record one injury event for each jam with the box checked.
            const numInjuries = warningData.jamsCalledInjury.filter(
                (x) => x.period == period && x.jam == (parseInt(j) + 1),
            ).length
            if (numInjuries >= 1) {
                sbData.periods[period].jams[j].events.push(
                    {
                        event: 'injury',
                    },
                )
            }

            // ERROR CHECK: Injury box checked for only one team in a jam.
            if (numInjuries == 1) {
                sbErrors.scores.injuryOnlyOnce.events.push(
                    `Period: ${period}, Jam: ${jam}`,
                )
            }

            // ERROR Check: Points scored with:
            // Neither team decleared lead
            // Scoring team not declared lost
            if (numLead == 0) {
                for (const t in teamList) {
                    const isLost = sbData.periods[period].jams[j].events.find(
                        (x) => x.event == 'lost' && x.skater.substr(0, 4) == teamList[t],
                    )
                    const scoreTrip = sbData.periods[period].jams[j].events.find(
                        (x) => x.event == 'pass' && x.team == teamList[t] && x.number > 1,
                    )
                    if (scoreTrip != undefined && isLost == undefined) {
                        sbErrors.scores.pointsNoLeadNoLost.events.push(
                            `Team: ${cap(teamList[t])}, Period: ${period}, Jam: ${jam}`,
                        )
                    }
                }
            }
        }
    }
    // All score data read

    // Error check: Star pass marked for only one team in a jam.
    for (const sp in starPasses) {
        if (starPasses.filter(
            (x) => x.period == starPasses[sp].period && x.jam == starPasses[sp].jam,
        ).length == 1) {
            sbErrors.scores.onlyOneStarPass.events.push(
                `Period: ${starPasses[sp].period} Jam: ${starPasses[sp].jam}`,
            )
        }
    }
}

const readPenalties = (workbook) => {
// Given a workbook, extract the data from the "Penalties" tab.

    let cells: any = {}

    const numberAddress: CellAddress = { c: null, r: null },
        penaltyAddress: CellAddress = { c: null, r: null },
        jamAddress: CellAddress = { c: null, r: null },
        foAddress: CellAddress = { c: null, r: null },
        foJamAddress: CellAddress = { c: null, r: null },
        benchExpCodeAddress: CellAddress = { c: null, r: null },
        benchExpJamAddress: CellAddress = { c: null, r: null },
        foulouts = [],
        maxPenalties = sbTemplate.penalties.maxPenalties,
        sheet = workbook.Sheets[sbTemplate.penalties.sheetName]

    for (let period = 1; period < 3; period ++) {
    // For each period

        const pstring = period.toString()

        const props = ['firstNumber', 'firstPenalty', 'firstJam',
            'firstFO', 'firstFOJam', 'benchExpCode', 'benchExpJam']
        const tab = 'penalties'

        for (const i in teamList) {
        // For each team

            const team = teamList[i]

            // Maximum number of skaters per team
            const maxNum = sbTemplate.teams[team].maxNum

            // Read in starting positions for penalty parameters
            cells = initCells(team, pstring, tab, props)
            numberAddress.c = cells.firstNumber.c
            penaltyAddress.c = cells.firstPenalty.c
            jamAddress.c = cells.firstJam.c
            foAddress.c = cells.firstFO.c
            foJamAddress.c = cells.firstFOJam.c

            for (let s = 0; s < maxNum; s++) {
            // For each player

                // Advance two rows per skater - TODO make this settable?
                numberAddress.r = cells.firstNumber.r + (s * 2)
                penaltyAddress.r = cells.firstPenalty.r + (s * 2)
                jamAddress.r = cells.firstJam.r + (s * 2)
                foAddress.r = cells.firstFO.r + (s * 2)
                foJamAddress.r = cells.firstFOJam.r + (s * 2)

                const skaterNum = sheet[utils.encode_cell(numberAddress)]

                if (skaterNum == undefined || skaterNum.v == '') {continue}

                // ERROR CHECK: skater on penalty sheet not on the IGRF
                if (sbData.teams[team].persons.findIndex((x) => x.number == skaterNum.v) == -1) {
                    // This SHOULD be caught by conditional formatting in Excel, but there
                    // are reports of that breaking sometimes.
                    sbErrors.penalties.penaltiesNotOnIGRF.events.push(
                        `Team: ${cap(team)}, Period: ${period}, Skater: ${skaterNum.v} `,
                    )
                }

                const skater = team + ':' + skaterNum.v

                for (let p = 0; p < maxPenalties; p++) {
                // For each penalty space

                    penaltyAddress.c = cells.firstPenalty.c + p
                    jamAddress.c = cells.firstJam.c + p

                    // Read the penalty code and jam number
                    const codeText = sheet[utils.encode_cell(penaltyAddress)]
                    const jamText = sheet[utils.encode_cell(jamAddress)]

                    const code = _.get(codeText, 'v')
                    const jam = _.get(jamText, 'v')

                    if (code == undefined || jam == undefined) {
                        // Error Check - penalty code without jam # or vice versa

                        if (code == undefined && jam == undefined) {
                            continue
                        } else {
                            sbErrors.penalties.codeNoJam.events.push(
                                `Team: ${cap(team)}, Skater: ${skaterNum.v}, Period: ${period}.`,
                            )
                            continue
                        }
                    }

                    if (jam > sbData.periods[period].jams.length || jam - 1 < 0 || typeof(jam) != 'number') {
                        // Error Check - jam number out of range
                        sbErrors.penalties.penaltyBadJam.events.push(
                            `Team: ${cap(team)}, Skater: ${skaterNum.v}, Period: ${period}, Recorded Jam: ${jam}`,
                        )
                        continue
                    }

                    // Add a penalty event to that jam
                    sbData.periods[period].jams[jam - 1].events.push(
                        {
                            event: 'penalty',
                            skater,
                            penalty: code,
                        },
                    )
                    penalties[skater].push([jam, code])

                }

                // Check for FO or EXP, add events
                const foCode = sheet[utils.encode_cell(foAddress)]
                const foJam = sheet[utils.encode_cell(foJamAddress)]
                const code = _.get(foCode, 'v')
                const jam = _.get(foJam, 'v')

                if (foCode == undefined || foJam == undefined) {

                    // Error Check: FO or EXP code without jam, or vice versa.
                    if (foCode != undefined || foJam != undefined) {
                        sbErrors.penalties.codeNoJam.events.push(
                            `Team: ${cap(team)}, Skater: ${skaterNum.v}, Period: ${period}.`,
                        )
                    }

                    // ERROR CHECK: Seven or more penalties with NO foulout entered
                    if (foulouts.indexOf(skater) == -1
                        && penalties[skater] != undefined
                        && penalties[skater].length > 6
                        && period === 2) {
                        sbErrors.penalties.sevenWithoutFO.events.push(
                            `Team: ${cap(team)}, Skater: ${skaterNum.v}`,
                        )
                    }

                    continue
                }

                if (typeof(jam) != 'number' ||
                    jam > sbData.periods[period].jams.length ||
                    jam - 1 < 0) {
                    sbErrors.penalties.foBadJam.events.push(
                        `Team: ${cap(team)}, Skater: ${skaterNum.v}, Period: ${period}, Recorded Jam: ${jam}`,
                    )
                    continue
                }

                // If there is expulsion, add an event.
                // Note that derbyJSON doesn't actually record foul-outs,
                // so only expulsions are recorded.
                if (code != 'FO') {
                    sbData.periods[period].jams[jam - 1].events.push(
                        {
                            event: 'expulsion',
                            skater,
                            notes: [
                                {note: 'Penalty: ' + code},
                                {note: 'Jam: ' + jam},
                            ],
                        },
                    )
                    warningData.expulsions.push(
                        {skater,
                            team,
                            period,
                            jam},
                    )

                    // ERROR CHECK: Expulsion code for a jam with no penalty
                    if (sbData.periods[period].jams[foJam.v - 1].events.filter(
                        (x) => x.event == 'penalty' && x.skater == skater,
                    ).length < 1) {
                        sbErrors.penalties.expulsionNoPenalty.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${foJam.v}, Skater: ${skaterNum.v}`,
                        )
                    }

                }

                // If there is a foul-out, add an event.
                if (foCode.v == 'FO') {
                    foulouts.push(skater)
                    warningData.foulouts.push(
                        {skater,
                            team,
                            period,
                            jam: foJam.v},
                    )
                }

                // ERROR CHECK: FO entered with fewer than seven penalties
                if (foCode.v == 'FO' && penalties[skater].length < 7) {
                    sbErrors.penalties.foUnder7.events.push(
                        `Team: ${cap(team)}, Period: ${period}, Skater: ${skaterNum.v}`,
                    )
                }

            }

            // Deal with bench expulsions
            benchExpCodeAddress.r = cells.benchExpCode.r
            benchExpJamAddress.r = cells.benchExpJam.r

            for (let e = 0; e < 2; e++) {
                benchExpCodeAddress.c = cells.benchExpCode.c + e
                benchExpJamAddress.c = cells.benchExpJam.c + e

                const benchExpCode = sheet[utils.encode_cell(benchExpCodeAddress)]
                const benchExpJam = sheet[utils.encode_cell(benchExpJamAddress)]

                if (benchExpCode == undefined || benchExpJam == undefined) {
                    continue
                }
                sbData.periods[period].jams[benchExpJam.v - 1].events.push(
                    {
                        event: 'expulsion',
                        notes: [
                            {note: 'Bench Staff Expulsion - ' + benchExpCode.v},
                            {note: 'Jam: ' + benchExpJam.v},
                        ],
                    },
                )

            }
        }
    }

}

const readLineups = (workbook: WorkBook) => {
// Read in the data from the lineups tab.

    let cells: any = {},
    skaterList = []

    const
        jamNumberAddress: CellAddress = { c: null, r: null },
        noPivotAddress: CellAddress = { c: null, r: null },
        skaterAddress: CellAddress = { c: null, r: null },
        maxJams = sbTemplate.lineups.maxJams,
        boxCodes = sbTemplate.lineups.boxCodes,
        sheet = workbook.Sheets[sbTemplate.lineups.sheetName],
        positions = {0: 'jammer', 1: 'pivot', 2: 'blocker', 3: 'blocker', 4: 'blocker'},
        box = {home: [], away: []},
        tab = 'lineups',
        props = ['firstJamNumber', 'firstNoPivot', 'firstJammer']

    for (let period = 1; period < 3; period++) {
    // For each period

        const pstring = period.toString()

        for (const i in teamList) {
        // For each team
            const team = teamList[i]
            let jam = 0
            let starPass = false

            cells = initCells(team, pstring, tab, props)
            jamNumberAddress.c = cells.firstJamNumber.c
            noPivotAddress.c = cells.firstNoPivot.c
            skaterAddress.c = cells.firstJammer.c

            for (let l = 0; l < maxJams; l++) {
            // For each line

                jamNumberAddress.r = cells.firstJamNumber.r + l
                noPivotAddress.r = cells.firstNoPivot.r + l
                skaterAddress.r = cells.firstJammer.r + l

                const jamText = sheet[utils.encode_cell(jamNumberAddress)]
                const noPivot = sheet[utils.encode_cell(noPivotAddress)]

                if (jamText == undefined ||
                    jamText.v == '' ||
                    /^\s+$/.test(jamText.v)) {continue}
                // If there is no jam number, go on to the next line.
                // TODO - maybe change this to not give up if the jam # is blank?

                if (anSP.test(jamText.v)) {
                // If this is a star pass line (SP or SP*)
                    starPass = true

                    if (!mySP.test(jamText.v)) {
                    // If this is an opposing team star pass only,
                    // Check for skaters that shouldn't be here, then go on.
                        let spStarSkater = false

                        for (let s = 0; s < 5; s++) {
                            skaterAddress.c = cells.firstJammer.c + (s * (boxCodes + 1))
                            const skaterText = sheet[utils.encode_cell(skaterAddress)]
                            if (skaterText != undefined && skaterText.v != false) {
                                spStarSkater = true
                            }
                        }

                        if (spStarSkater) {
                            sbErrors.lineups.spStarSkater.events.push(
                                `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}`,
                            )
                        }

                        continue
                    }

                    if (_.get(noPivot, 'v') == undefined) {
                        // Error check: Star Pass line without "No Pivot" box checked.

                        sbErrors.lineups.starPassNoPivot.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}`,
                        )
                    }

                } else {
                    // Not a starpass line, update the jam number
                    jam = jamText.v
                    starPass = false
                    skaterList = []
                }

                // Retrieve penalties from this jam and prior jam for
                // error checking later
                const thisJamPenalties = sbData.periods[pstring].jams[jam - 1].events.filter(
                    (x) => (x.event == 'penalty' && x.skater.substr(0, 4) == team),
                )
                let priorJamPenalties = []
                if (jam != 1) {
                    priorJamPenalties = sbData.periods[pstring].jams[jam - 2].events.filter(
                        (x) => (x.event == 'penalty' && x.skater.substr(0, 4) == team),
                    )
                } else if (period == 2) {
                    priorJamPenalties = sbData.periods['1'].jams[
                        sbData.periods['1'].jams.length - 1
                    ].events.filter(
                        (x) => (x.event == 'penalty' && x.skater.substr(0, 4) == team),
                    )
                }

                for (let s = 0; s < 5; s++) {
                // For each skater
                    let position = ''

                    skaterAddress.c = cells.firstJammer.c + (s * (boxCodes + 1))
                    const skaterText = sheet[utils.encode_cell(skaterAddress)]

                    if (skaterText == undefined ||
                        (skaterText.v == undefined ||
                            skaterText.v == '?' ||
                            skaterText.v == 'n/a' ||
                            skaterText.v == 'N/A')
                    ) {
                        if (skaterText == undefined || (skaterText.v == undefined && skaterText.c == undefined)) {
                        // WARNING: Empty box on Lineups without comment
                            sbErrors.warnings.emptyLineupNoComment.events.push(
                                `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}, Column: ${s + 1}`,
                            )
                        }
                        continue
                    }

                    const skater = team + ':' + skaterText.v

                    // ERROR CHECK: Skater on lineups not on IGRF
                    if (sbData.teams[team].persons.findIndex((x) => x.number == skaterText.v) == -1) {
                    // If the skater is not on the IGRF, record an error.
                        // This SHOULD be caught by conditional formatting in Excel, but there
                        // are reports of that breaking sometimes.
                        sbErrors.lineups.lineupsNotOnIGRF.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}, Skater: ${skaterText.v} `,
                        )
                    }

                    // ERROR CHECK: Same skater entered more than once per jam
                    if (skaterList.indexOf(skater) != -1 && !starPass) {
                        sbErrors.lineups.samePlayerTwice.events.push(
                            `Team: ${cap(team)}, Period: ${period}, Jam: ${jam}, Skater: ${skaterText.v}`,
                        )
                    }

                    if (!starPass) {skaterList.push(skater)}

                    if (s == 1 && noPivot != undefined && noPivot.v != undefined) {
                        position = 'blocker'
                    } else {
                        position = positions[s]
                    }

                    if (!starPass) {
                    // Unless this is a star pass, add a
                    // "lineup" event for that skater with the position
                        sbData.periods[pstring].jams[jam - 1].events.push(
                            {
                                event: 'lineup',
                                skater,
                                position,
                            },
                        )

                    }

                    let allCodes = ''
                    // Add box codes if present
                    for (let c = 1; c <= boxCodes; c++) {
                        // for each code box

                        skaterAddress.c = cells.firstJammer.c + (s * (boxCodes + 1)) + c
                        const codeText = sheet[utils.encode_cell(skaterAddress)]

                        if (codeText == undefined) {continue}

                        allCodes += codeText.v

                        switch (sbSummary.version) {
                        case '2017':
                        case '2018':
                            // Possible codes - /, X, S, $, I or |, 3
                            // Possible events - enter box, exit box, injury
                            // / - Enter box
                            // X - Test to see if skater is IN box
                            //      Yes: exit box, No: enter box, exit box
                            // S - Enter box, note: sat between jams
                            // $ - Enter box, exit box, note: sat between jams
                            // I or | - no event, error checking only
                            // 3 - Injury object, verify not already present from score tab

                            switch (codeText.v) {
                            case '/':
                                // Add an "Enter Box" event, and push the skater onto the box list
                                enterBox(pstring, jam, skater)
                                box[team].push(skater)

                                // ERROR CHECK: Skater enters the box during the jam
                                // without a penalty in the current jam.
                                if (thisJamPenalties.find(
                                    (x) => x.skater == skater,
                                ) == undefined) {
                                    sbErrors.lineups.slashNoPenalty.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                }
                                break
                            case 'ᚾ':
                            // Error Check: Using the rune instead of an X
                                sbErrors.lineups.runeUsed.events.push(
                                    `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}`,
                                )
                                // break omitted
                            case 'X':
                            case 'x':
                                if (!box[team].includes(skater)) {
                                    // If the skater is not in the box, add an "enter box" event
                                    enterBox(pstring, jam, skater)

                                    // ERROR CHECK: Skater enters the box during the jam
                                    // without a penalty in the current jam.
                                    if (thisJamPenalties.find(
                                        (x) => x.skater == skater,
                                    ) == undefined) {
                                        sbErrors.lineups.xNoPenalty.events.push(
                                            `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                        )
                                        warningData.badContinues.push({
                                            skater,
                                            team,
                                            period,
                                            jam,
                                        })
                                    }

                                }
                                // Whether or not the skater started in the box, add an "exit box" event
                                exitBox(pstring, jam, skater)

                                // Remove the skater from the box list.
                                if (box[team].includes(skater)) {
                                    remove(box[team], skater)
                                }
                                break

                            case 'S':
                            case 's':
                                // Add a box entry, with a note that the skater sat between jams.
                                enterBox(pstring, jam, skater, 'Sat Between Jams.')

                                // ERROR CHECK: Skater starts in the box while already in the box.
                                if (box[team].includes(skater)) {
                                    sbErrors.lineups.startsWhileThere.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                } else {
                                    // Add skater to the box list.
                                    box[team].push(skater)
                                }

                                // ERROR CHECK: Skater starts in the box without a penalty
                                // in the prior or current jam.
                                if (thisJamPenalties.find((x) => x.skater == skater) == undefined
                                    && priorJamPenalties.find((x) => x.skater == skater) == undefined) {
                                    sbErrors.lineups.sNoPenalty.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                    warningData.badStarts.push({
                                        skater,
                                        team,
                                        period,
                                        jam,
                                    })
                                }
                                break

                            case '$':
                                enterBox(pstring, jam, skater, 'Sat Between Jams.')
                                exitBox(pstring, jam, skater)

                                // ERROR CHECK: Skater starts in the box while already in the box.
                                if (box[team].includes(skater)) {
                                    sbErrors.lineups.startsWhileThere.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                    remove(box[team], skater)
                                }

                                // ERROR CHECK: Skater starts in the box without a penalty
                                // in the prior or current jam.
                                if (thisJamPenalties.find((x) => x.skater == skater) == undefined
                                    && priorJamPenalties.find((x) => x.skater == skater) == undefined) {
                                    sbErrors.lineups.sSlashNoPenalty.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                    warningData.badStarts.push({
                                        skater,
                                        team,
                                        period,
                                        jam,
                                    })
                                }

                                break
                            case 'I':
                            case '|':
                                // no derbyJSON event, but use this branch for error checking
                                if (!box[team].includes(skater)) {
                                    const priorFoulout = warningData.foulouts.filter((x) =>
                                        (x.period == period && x.jam < jam && x.skater == skater) ||
                                        (x.period < period && x.skater == skater))
                                    if (priorFoulout.length > 0) {
                                        sbErrors.lineups.foInBox.events.push(
                                            `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                        )
                                    } else {
                                        sbErrors.lineups.iNotInBox.events.push(
                                            `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                        )
                                    }
                                    warningData.badContinues.push({
                                        skater,
                                        team,
                                        period,
                                        jam,
                                    })
                                }
                                break
                            case '3':
                            case 3:
                                // Since '3' does not necessarily mean the jam was called, not enough information
                                // here to conclusively record a derbyJSON injury event, which specifies that the
                                // jam was called for injury.   However, save the skater information for error
                                // checking later.
                                warningData.lineupThree.push({
                                    skater,
                                    team,
                                    period,
                                    jam,
                                })
                                break
                            default:
                            // Handle invalid lineup codes
                                sbErrors.lineups.badLineupCode.events.push(
                                    `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}, Code: ${codeText.v}`,
                                )
                                break
                            }
                            break
                        case '2019':
                            // Possible codes - -, +, S, $, 3

                            // - - Enter box
                            // + - Enter and exit box
                            // S - Sat between jams or continued
                            // $ - Sat between jams or continued with exit
                            // 3 - Injury object, verify not already present from score tab

                            switch (codeText.v) {
                            case '-':
                                // Add an "Enter Box" event, and push the skater onto the box list
                                enterBox(pstring, jam, skater)
                                box[team].push(skater)

                                // ERROR CHECK: Skater enters the box during the jam
                                // without a penalty in the current jam.
                                if (thisJamPenalties.find(
                                    (x) => x.skater == skater,
                                ) == undefined) {
                                    sbErrors.lineups.dashNoPenalty.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                }
                                break
                            case '+':

                                enterBox(pstring, jam, skater)
                                exitBox(pstring, jam, skater)

                                // ERROR CHECK: Skater enters the box during the jam
                                // without a penalty in the current jam.
                                if (thisJamPenalties.find(
                                    (x) => x.skater == skater,
                                ) == undefined) {
                                    sbErrors.lineups.plusNoPenalty.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                }

                                break

                            case 'S':
                            case 's': {

                                // ERROR CHECK: skater who has fouled out starting in box
                                const priorFoulout = warningData.foulouts.filter((x) =>
                                    (x.period == period && x.jam < jam && x.skater == skater) ||
                                    (x.period < period && x.skater == skater))
                                if (priorFoulout.length > 0) {
                                    sbErrors.lineups.foInBox.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                }

                                if (!box[team].includes(skater)) {
                                // If the skater is not already in the box:

                                    // ERROR CHECK: Skater starts in the box without a penalty
                                    // in the prior or current jam.
                                    if (thisJamPenalties.find((x) => x.skater == skater) == undefined
                                        && priorJamPenalties.find((x) => x.skater == skater) == undefined) {
                                        sbErrors.lineups.sNoPenalty.events.push(
                                            `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                        )
                                        warningData.badStarts.push({
                                            skater,
                                            team,
                                            period,
                                            jam,
                                        })
                                    }

                                    // Add a box entry, and add the skater to the box list
                                    enterBox(pstring, jam, skater, 'Sat Between Jams.')
                                    box[team].push(skater)
                                }

                                break
                            }
                            case '$': {
                                // ERROR CHECK: skater who has fouled out starting in box
                                const priorFoulout = warningData.foulouts.filter((x) =>
                                    (x.period == period && x.jam < jam && x.skater == skater) ||
                                    (x.period < period && x.skater == skater))
                                if (priorFoulout.length > 0) {
                                    sbErrors.lineups.foInBox.events.push(
                                        `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                    )
                                }

                                if (!box[team].includes(skater)) {
                                // If the skater is not already in the box:

                                    // ERROR CHECK: Skater starts in the box without a penalty
                                    // in the prior or current jam.
                                    if (thisJamPenalties.find((x) => x.skater == skater) == undefined
                                        && priorJamPenalties.find((x) => x.skater == skater) == undefined) {
                                        sbErrors.lineups.sSlashNoPenalty.events.push(
                                            `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}`,
                                        )
                                        warningData.badStarts.push({
                                            skater,
                                            team,
                                            period,
                                            jam,
                                        })
                                    }

                                    // Add a box entry, and add the skater to the box list
                                    enterBox(pstring, jam, skater, 'Sat Between Jams.')
                                    exitBox(pstring, jam, skater)
                                } else {
                                    exitBox(pstring, jam, skater)
                                    remove(box[team], skater)
                                }

                                break
                            }
                            case '3':
                            case 3:
                                // Since '3' does not necessarily mean the jam was called, not enough information
                                // here to conclusively record a derbyJSON injury event, which specifies that the
                                // jam was called for injury.   However, save the skater information for error
                                // checking later.
                                warningData.lineupThree.push({
                                    skater,
                                    team,
                                    period,
                                    jam,
                                })
                                break
                            default:
                                // Handle invalid lineup codes
                                sbErrors.lineups.badLineupCode.events.push(
                                    `Team: ${cap(team)}, Period: ${pstring}, Jam: ${jam}, Skater: ${skaterText.v}, Code: ${codeText.v}`,
                                )
                                break
                            }
                            break
                        default:
                        // Handle unrecognized  statsbook versions?
                            break
                        }
                    }
                    // Done reading all codes

                    // ERROR CHECK: is this skater still in the box without
                    // any code on the present line?
                    if (box[team].includes(skater) && !allCodes) {
                        sbErrors.lineups.seatedNoCode.events.push(
                            `Team: ${
                                cap(skater.substr(0, 4))
                            }, Period: ${pstring}, Jam: ${jam}, Skater: ${skater.slice(5)}`,
                        )
                        warningData.noExits.push({
                            skater,
                            team,
                            period,
                            jam,
                        })
                        remove(box[team], skater)
                    }
                    // Done processing skater

                }
                // Done reading line

                // Remove fouled out or expelled skaters from the box
                const fouledOutSkaters = warningData.foulouts.filter((x) => x.period == period && x.jam == jam && x.team == team)
                if (fouledOutSkaters != undefined) {
                    for (const s in fouledOutSkaters) {
                        const skater = fouledOutSkaters[s].skater
                        if (box[team].includes(skater)) {
                            remove(box[team], skater)
                        }
                    }
                }
                const expelledSkaters = warningData.expulsions.filter((x) => x.period == period && x.jam == jam && x.team == team)
                if (expelledSkaters != undefined) {
                    for (const s in expelledSkaters) {
                        const skater = expelledSkaters[s].skater
                        if (box[team].includes(skater)) {
                            remove(box[team], skater)
                        }
                    }
                }

                // Error Check: Skater still in the box not listed on lineup tab at all
                for (const s in box[team]) {
                    const skater = box[team][s]
                    if (!skaterList.includes(skater)) {
                        sbErrors.lineups.seatedNotLinedUp.events.push(
                            `Team: ${
                                cap(skater.substr(0, 4))
                            }, Period: ${pstring}, Jam: ${jam}, Skater: ${skater.slice(5)}`,
                        )
                        warningData.noExits.push({
                            jam,
                            period,
                            skater,
                            team,
                        })
                    }
                }

                // ERROR CHECK: Skaters with penalties in this jam not listed on the lineup tab
                for (const p in thisJamPenalties) {
                    if (skaterList.indexOf(thisJamPenalties[p].skater) == -1) {
                        sbErrors.penalties.penaltyNoLineup.events.push(
                            `Team: ${
                                cap(thisJamPenalties[p].skater.substr(0, 4))
                            }, Period: ${pstring}, Jam: ${jam}, Skater: ${thisJamPenalties[p].skater.slice(5)}`,
                        )
                    }
                }
            }

        }
    }

}

const errorCheck = () => {
    // Run error checks that occur after all data has been read

    let jams = 0,
        events = [],
        pstring = ''

    for (let period = 1; period <= Object.keys(sbData.periods).length; period++) {

        pstring = period.toString()
        jams  = sbData.periods[pstring].jams.length

        for (let jam = 1; jam <= jams; jam++) {
            events = sbData.periods[pstring].jams[jam - 1].events

            // Get the list of Penalties in this jam
            const thisJamPenalties = events.filter(
                (x) => x.event == 'penalty',
            )

            // Get lead jammer if present (will only catch FIRST if two are marked)
            let leadJammer = ''
            const leadEvent = events.filter((x) => x.event == 'lead')
            if (leadEvent.length != 0) {
                leadJammer = leadEvent[0].skater
            }

            // Get the list of box entires in the current jam and the next one
            const thisJamEntries = events.filter(
                (x) => x.event == 'enter box',
            )
            let nextJamEntries = []
            if (period == 1 && jam == (jams)) {
                // If this is the last jam of the 1st period, get period 2, jam 1
                try {
                    nextJamEntries = sbData.periods['2'].jams[0].events.filter(
                        (x) => x.event == 'enter box',
                    )
                } catch (e) {
                    nextJamEntries = []
                }
            } else if (jam != (jams)) {
                // Otherwise, just grab the next jam (don't forget 0 indexing)
                nextJamEntries = sbData.periods[pstring].jams[jam].events.filter(
                    (x) => x.event == 'enter box',
                )
            }   // Last jam of the 2nd period gets ignored.

            // ERROR CHECK: Penalty without box entry in this jam
            // or the following jam.
            for (const pen in thisJamPenalties) {
                if (thisJamEntries.filter(
                    (x) => x.skater == thisJamPenalties[pen].skater,
                ).length == 0 && nextJamEntries.filter(
                    (x) => x.skater == thisJamPenalties[pen].skater,
                ).length == 0) {
                    if (!(jam == jams && period == 2)) {
                        sbErrors.penalties.penaltyNoEntry.events.push(
                            `Team: ${
                                cap(thisJamPenalties[pen].skater.substr(0, 4))
                            }, Period: ${period}, Jam: ${jam}, Skater: ${thisJamPenalties[pen].skater.slice(5)}`,
                        )
                    } else {
                        sbErrors.warnings.lastJamNoEntry.events.push(
                            `Team: ${
                                cap(thisJamPenalties[pen].skater.substr(0, 4))
                            }, Period: 2, Jam: ${jam}, Skater: ${thisJamPenalties[pen].skater.slice(5)}`,
                        )
                    }
                    warningData.noEntries.push({
                        skater: thisJamPenalties[pen].skater,
                        team: thisJamPenalties[pen].skater.substr(0, 4),
                        period,
                        jam,
                    })
                }
            }

            // Warning check: Jammer with lead and penalty, but not lost
            if (leadJammer != ''
                && thisJamPenalties.filter((x) => x.skater == leadJammer).length != 0
                && events.filter((x) => x.event == 'lost' && x.skater == leadJammer).length == 0
            ) {
                sbErrors.warnings.leadPenaltyNotLost.events.push(
                    `Team: ${
                        cap(leadJammer.substr(0, 4))
                    }, Period: ${period}, Jam: ${jam}, Jammer: ${leadJammer.slice(5)}`,
                )
            }
        }
    }
}

const warningCheck = () => {
    // Run checks for things that should throw warnings but not errors.

    // Warning check: Possible substitution.
    // For each skater who has a $ or S without a corresponding penalty,
    // check to see if a different skater on the same team has
    // a penalty without a subsequent box exit.
    for (const event in warningData.badStarts) {
        const bs = warningData.badStarts[event]
        if (warningData.noEntries.filter(
            (ne) => (ne.team == bs.team &&
                (
                    (ne.period == bs.period && ne.jam == (bs.jam - 1)) ||
                    (ne.period == (bs.period - 1) && bs.jam == 1)
                )
            )).length >= 1) {
            if (bs.jam != 1) {
                sbErrors.warnings.possibleSub.events.push(
                    `Team: ${cap(bs.team)}, Period: ${bs.period
                    }, Jams: ${bs.jam - 1} & ${bs.jam}`,
                )
            } else {
                sbErrors.warnings.possibleSub.events.push(
                    `Team: ${cap(bs.team)}, Period: 1, Jam: ${sbData.periods['1'].jams.length
                    } & Period: 2, Jam: ${bs.jam}`,
                )
            }
        }
    }

    // Warning check: Possible substitution.
    // For each skater who has a I, |, X or x without a corresponding penalty,
    // check to see if a different skater on the same team has
    // a penalty without a subsequent box exit.

    for (const event in warningData.badContinues) {
        // For each skater who is on the "continued without entry" list
        const bc = warningData.badContinues[event]

        // If there's a corresponding entry on the "never exited the box list", issue a warning
        if (warningData.noExits.filter(
            (ne) => (ne.team == bc.team &&
                (
                    (ne.period == bc.period && ne.jam == bc.jam)
                )
            )).length >= 1) {
            if (bc.jam != 1) {
                sbErrors.warnings.possibleSub.events.push(
                    `Team: ${cap(bc.team)}, Period: ${bc.period
                    }, Jams: ${bc.jam - 1} & ${bc.jam}`,
                )
            } else {
                sbErrors.warnings.possibleSub.events.push(
                    `Team: ${cap(bc.team)}, Period: 1, Jam: ${sbData.periods['1'].jams.length
                    } & Period: 2, Jam: ${bc.jam}`,
                )
            }
        }

        // If there's a skater in the prior jam with a foulout, issue a warning as well
        if (warningData.foulouts.filter(
            (fo) => fo.team == bc.team &&
            (
                (fo.period == bc.period && fo.jam == bc.jam - 1) ||
                (bc.period == 2 && bc.jam == 1 && fo.jam == sbData.periods['1'].jams.length)
            ),
        ).length > 0) {
            if (bc.jam != 1) {
                sbErrors.warnings.possibleSub.events.push(
                    `Team: ${cap(bc.team)}, Period: ${bc.period
                    }, Jams: ${bc.jam - 1} & ${bc.jam}`,
                )
            } else {
                sbErrors.warnings.possibleSub.events.push(
                    `Team: ${cap(bc.team)}, Period: 1, Jam: ${sbData.periods['1'].jams.length
                    } & Period: 2, Jam: ${bc.jam}`,
                )
            }
        }

        // If there's a skater in the prior jam with an expulsion, issue a warning as well
        if (warningData.expulsions.filter(
            (exp) => exp.team == bc.team &&
            (
                (exp.period == bc.period && exp.jam == bc.jam - 1) ||
                (bc.period == 2 && bc.jam == 1 && exp.jam == sbData.periods['1'].jams.length)
            ),
        ).length > 0) {
            if (bc.jam != 1) {
                sbErrors.warnings.possibleSub.events.push(
                    `Team: ${cap(bc.team)}, Period: ${bc.period
                    }, Jams: ${bc.jam - 1} & ${bc.jam}`,
                )
            } else {
                sbErrors.warnings.possibleSub.events.push(
                    `Team: ${cap(bc.team)}, Period: 1, Jam: ${sbData.periods['1'].jams.length
                    } & Period: 2, Jam: ${bc.jam}`,
                )
            }
        }

    }

    // Warning Check - lost lead without a penalty
    for (const s in warningData.lost) {
        const lost = warningData.lost[s]
        if (sbData.periods[lost.period].jams[lost.jam - 1].events.filter(
            (event) => event.skater == lost.skater && event.event == 'penalty',
        ).length == 0) {
            sbErrors.warnings.lostNoPenalty.events.push(
                `Team: ${cap(lost.team)}, Period: ${lost.period
                }, Jam: ${lost.jam}, Skater: ${lost.skater.slice(5)}`,
            )
        }
    }

    // Warning Check - jam called for injury without a skater marked with a "3"
    // Note: filtered for home team to prevent duplicate errors.
    for (const j in warningData.jamsCalledInjury) {
        const injJam = warningData.jamsCalledInjury[j]
        if (!warningData.lineupThree.find((x) => x.jam == injJam.jam)) {
            sbErrors.warnings.injNoThree.events.push(
                `Period: ${injJam.period}, Jam: ${injJam.jam}`,
            )
        }
    }
    // Remove duplicates
    sbErrors.warnings.injNoThree.events = sbErrors.warnings.injNoThree.events.filter(
        (v, i, a) => a.indexOf(v) === i,
    )

}

const sbErrorsToTable = () => {
    // Build error report

    const errorTypes = ['scores', 'lineups', 'penalties', 'warnings']
    const typeHeaders = ['Scores', 'Lineups', 'Penalties', 'Warnings - These should be checked, but may be OK']
    const table = document.createElement('table')
    table.setAttribute('class', 'table')

    for (const t in errorTypes) {
        // For each of the three types of errors

        const section = errorTypes[t]

        const secHead = document.createElement('tr')
        const secCell = document.createElement('th')
        secCell.appendChild(document.createTextNode(typeHeaders[t]))
        secHead.appendChild(secCell)
        secHead.setAttribute('class', 'thead-dark')

        table.appendChild(secHead)
        let noErrors = true
        // TODO - Add Tip Icon

        for (const e in sbErrors[errorTypes[t]]) {
            // For each error in the type

            if (sbErrors[errorTypes[t]][e].events.length == 0) {
                continue
            }
            noErrors = false
            const descRow = document.createElement('tr')
            const descCell = document.createElement('th')
            descCell.appendChild(document.createTextNode(
                `${sbErrors[section][e].description} `,
            ))
            descCell.setAttribute('data-toggle', 'tooltip')
            descCell.setAttribute('title', sbErrors[section][e].long)
            const icon = document.createElement('i')
            icon.setAttribute('class', 'fa fa-info-circle')
            descCell.appendChild(icon)
            descRow.appendChild(descCell)
            descRow.setAttribute('class', 'thead-light')

            table.appendChild(descRow)

            for (const v in sbErrors[errorTypes[t]][e].events) {
                const evRow = document.createElement('tr')
                const evCell = document.createElement('td')
                evCell.appendChild(document.createTextNode(
                    sbErrors[section][e].events[v],
                ))
                evRow.appendChild(evCell)

                table.appendChild(evRow)
            }

        }
        if (noErrors) {secHead.remove()}
    }

    if (table.rows.length == 0) {
        const secHead = document.createElement('tr')
        const secCell = document.createElement('th')
        secCell.appendChild(document.createTextNode('No Errors Found!'))
        secHead.appendChild(secCell)
        table.appendChild(secHead)
    }

    return table
}

const cellVal = (sheet, address) => {
    // Given a worksheet and a cell address, return the value
    // in the cell if present, and undefined if not.
    if (sheet[address] && sheet[address].v) {
        return sheet[address].v
    } else {
        return undefined
    }
}

const initCells = (team, period, tab, props) => {
    // Given a team, period, SB section, and list of properties,
    // return an object of addresses for those properties.
    // Team should be 'home' or 'away'
    const cells = {}

    for (const i in props) {
        cells[props[i]] = utils.decode_cell(
            sbTemplate[tab][period][team][props[i]])
    }

    return cells
}

const remove = (array, element) => {
    // Lifted from https://blog.mariusschulz.com/
    // Removes an element from an arry
    const index = array.indexOf(element)

    if (index !== -1) {
        array.splice(index, 1)
    }
}

const enterBox = (pstring, jam, skater, note = undefined) => {
// Add an 'enter box' event
    const event: any = {
        event: 'enter box',
        skater,
    }

    if (note !== undefined) {
        event.note = note
    }

    sbData.periods[pstring].jams[jam - 1].events.push(event)

}

const exitBox = (pstring, jam, skater) => {
// Add an 'exit box' event
    sbData.periods[pstring].jams[jam - 1].events.push(
        {
            event: 'exit box',
            skater,
        },
    )
}

// Change appearance of input box on file dragover
holder.ondragover = () => {
    holder.classList.add('box__ondragover')
    return false
}

holder.ondragleave = () => {
    holder.classList.remove('box__ondragover')
    return false
}

holder.ondragend = () => {
    return false
}

ipc.on('save-derby-json', () => {
    // Saves statsbook data to a JSON file

    const blob = new Blob( [ JSON.stringify(sbData, null, ' ') ], { type: 'application/json' })

    download(blob, sbFilename.split('.')[0] + '.json')
})

ipc.on('export-crg-roster', () => {
    // Exports statsbook rosters in CRG Scoreboard XML Format

    const teams = extractTeamsFromSBData(sbData)
    const xml = exportXml(teams)

    const data = encode(xml.end({pretty: true}))
    const blob = new Blob( [data], { type: 'text/xml'})
    download(blob, sbFilename.split('.')[0] + '.xml')
})

ipc.on('export-crg-roster-json', () => {
    // Exports statsbook rosters in CRG Scoreboard's Beta JSON Format
    const teams = extractTeamsFromSBData(sbData)
    const json = exportJsonRoster(teams)

    const blob = new Blob( [JSON.stringify(json, null, ' ')], { type: 'application/json' })
    download(blob, sbFilename.split('.')[0] + '.json')
})

const encode = (s) => {
    const out = []
    for ( let i = 0; i < s.length; i++ ) {
        out[i] = s.charCodeAt(i)
    }
    return new Uint16Array( out )
}

window.onerror = (msg, url, lineNo, columnNo) => {
    ipc.send('error-thrown', msg, url, lineNo, columnNo)
    return false
}

/*
List of error checks.

Check while reading:

Just Scores
1. NI checked with points.*
2. No points (including a zero) entered, but NI *not* checked.*
3. "Lead" checked for both jammers.*
4. "Call" checked for both jammers.*
5. "Injury" checked on one team but not the other.*
6. Star pass for only one team.*
7. Jam Number out of sequence
8. Points given to more than one jammer in the same trip during a star pass.
9. Skipped column on score sheet.
10. SP* with jammer number entered.

Just Penalties
1. "FO" entered for skater with fewer than 7 penalties.*
2. Seven or more penalties without "FO" or expulsion code entered.*
3. Expulsion code entered for jam with no penalty.*
4. Penalty code without jam number, of jam number without penalty.

Just Lineups
1. Players listed more than once in the same jam on the lineup tab.*
2. "I" or "|" in lineups without the player being in the box already.*
3. Skater previously seated in the box with no code on the present line.
4. A player seated in a prior jam who has no marked exit from the box.*
5. "S" or "$" entered for skater already seated in the box.
6. "No Pivot" box not checked after star pass.
7. "ᚾ" used in place of "X".

Lineups + Penalties (Check while reading lineups)
1. Penalties on skaters not listed on the lineup for that jam.*
2. "X" in lineups without a matching penalty.*
3. "/" in lineups without a matching penalty.*
4. "S" or "$" in lineups without a matching penalty.*

Check after all data read:
Lineups + Penalties:
1. Penalty recorded without a "X", "/", "S", or "$". *

Scores + Penalties
1. Jammers with lead and a penalty, but not marked "lost."*
2. Penalties with jam numbers marked that are not on the score sheet.

List of error checks I'm not going to BOTHER implementing, because
the statsbook now has conditional formatting to flag them
1. Skater numbers on any sheet not on the IGRF.
2. Jammers that don't match between lineups and scores.
3. SP matching between lineups and scores.
*/

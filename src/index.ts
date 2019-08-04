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
    sbFilename = '',
    warningData: any = {}

let sbTemplate: IStatsbookTemplate = null
const teamList = ['home', 'away']
const anSP = /^sp\*?$/i
const mySP = /^sp$/i

// Check for new version
ipc.on('do-version-check', (event: any, version: any) => {
    fetch('https://api.github.com/repos/AdamSmasherDerby/Statsbook-Tool/tags')
        .then((result) =>  result.json())
        .then((data) => {
            const latestVersion = data[0].name
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

    sbReader = new WorkbookReader(workbook, filename)

    sbSummary = sbReader.summary
    sbErrors = sbReader.errors
    sbData = sbReader.data
    warningData = sbReader.warnings
    sbTemplate = sbReader.template

    updateFileInfo()

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

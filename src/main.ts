import {app, BrowserWindow, dialog, ipcMain as ipc, Menu } from 'electron'
import isDev from 'electron-is-dev'
import { join } from 'path'
import { format as formaturl } from 'url'

let menu: Menu
let win: BrowserWindow
let helpWin: BrowserWindow
let aboutWin: BrowserWindow

const createWindow = () => {
    win = new BrowserWindow({
        title: 'Statsbook Tool',
        icon: join(__dirname, '../build/flamingo-white.png'),
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true,
        },
    })

    win.loadURL(formaturl({
        pathname: join(__dirname, '../src/index.html'),
        protocol: 'file',
        slashes: true,
    }))

    if (isDev) {
        win.webContents.openDevTools()
        require('devtron').install()
    }

    // Prevent files dropped outside of the drop zone from doing anything.
    win.webContents.on('will-navigate', (event) => event.preventDefault())

    win.on('closed', () => {
        win = null
    })

    win.webContents.on('crashed', () => {
        dialog.showMessageBox(win, {
            type: 'error',
            title: 'Statsbook Tool',
            message: 'Statsbook Tool has crashed.  This should probably not surprise you.',
        })
    })

    win.on('unresponsive', () => {
        dialog.showMessageBox(win, {
            type: 'error',
            title: 'Statsbook Tool',
            // tslint:disable-next-line: max-line-length
            message: 'Statsbook Tool has become unresponsive.  You should probably have been more emotionally supportive.',
        })
    })

    menu = Menu.buildFromTemplate([
        {
            label: 'File',
            submenu: [
                {
                    id: 'exportXML',
                    label: 'Export Roster to CRG XML',
                    click() {
                        win.webContents.send('export-crg-roster')
                    },
                    enabled: false,
                },
                {
                    id: 'exportJSON',
                    label: 'Export Roster to CRG JSON (beta)',
                    click() {
                        win.webContents.send('export-crg-roster-json')
                    },
                    enabled: false,
                },

                {
                    id: 'exportDerbyJSON',
                    label: 'Save DerbyJSON',
                    click() {
                        win.webContents.send('save-derby-json')
                    },
                    enabled: false,
                },
                {
                    label: 'Exit',
                    accelerator:  'CmdOrCtrl+Q',
                    click() {
                        app.quit()
                    },
                },
            ],
        },
        {
            label: 'Edit',
            submenu: [
                {
                    label: 'Copy',
                    accelerator: 'CmdOrCtrl+C',
                },
                {
                    label: 'Paste',
                    accelerator: 'CmdOrCtrl+V',
                },
                {
                    label: 'Select All',
                    accelerator: 'CmdOrCtrl+A',
                },
            ],
        },
        {
            label: 'Help',
            submenu: [
                {
                    label: 'Error Descriptions',
                    click() {
                        openHelp()
                    },
                },
                {
                    label: 'About',
                    click() {
                        openAbout()
                    },
                },
            ],
        },
    ])
    Menu.setApplicationMenu(menu)

    // Do version check
    win.webContents.on('did-finish-load', () => {
        win.webContents.send('do-version-check', app.getVersion())
    })

    win.webContents.on('new-window', (e, url) => {
        e.preventDefault()
        require('electron').shell.openExternal(url)
    })
}

const openAbout = () => {
    aboutWin = new BrowserWindow({
        parent: win,
        title: 'StatsBook Tool',
        icon: join(__dirname, '../build/flamingo-white.png'),
        width: 300,
        height: 300,
        x: win.getPosition()[0] + 250,
        y: win.getPosition()[1] + 150,
    })

    aboutWin.setMenu(null)

    aboutWin.loadURL(formaturl({
        pathname: join(__dirname, '../src/aboutst.html'),
        protocol: 'file',
        slashes: true,
    }))

    aboutWin.webContents.on('new-window', (e, url) => {
        e.preventDefault()
        require('electron').shell.openExternal(url)
    })

    aboutWin.on('closed', () => {
        aboutWin = null
    })

    aboutWin.webContents.on('did-finish-load', () => {
        aboutWin.webContents.send('set-version', app.getVersion())
    })

}

const openHelp = () => {
    helpWin = new BrowserWindow({
        parent: win,
        title: 'Error Descriptions',
        icon: join(__dirname, '../build/flamingo-white.png'),
        width: 800,
        height: 600,
        x: win.getPosition()[0] + 20,
        y: win.getPosition()[1] + 20,
    })

    helpWin.loadURL(formaturl({
        pathname: join(__dirname, '../src/help.html'),
        protocol: 'file',
        slashes: true,
    }))

    helpWin.setMenu(null)

    helpWin.on('closed', () => {
        helpWin = null
    })

}

app.on('ready', createWindow)

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit()
    }
})

app.on('activate', () => {
    if (win == null) {
        createWindow()
    }
})

ipc.on('enable-menu-items', () => {

    menu.getMenuItemById('exportXML').enabled = true
    menu.getMenuItemById('exportJSON').enabled = true
    menu.getMenuItemById('exportDerbyJSON').enabled = true
})

ipc.on('error-thrown', (event: any, msg: any, url: any, lineNo: any, columnNo: any) => {
    dialog.showMessageBox(win, {
        type: 'error',
        title: 'Statsbook Tool',
        message: `Statsbook Tool has encountered an error.
        Here's some details:
        Message: ${msg}
        URL: ${url}
        Line Number: ${lineNo}
        Column Number: ${columnNo}
        Does this help?  It probably doesn't help.`,
    })
})

process.on('uncaughtException', (err) => {
    dialog.showMessageBox(win, {
        type: 'error',
        title: 'Statsbook Tool',
        // tslint:disable-next-line: max-line-length
        message: `Statsbook Tool has had an uncaught exception in main.js.  Does this help? (Note: will probably not help.) ${err}`,
    })
})

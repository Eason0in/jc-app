const { app, BrowserWindow, ipcMain, dialog, Menu, MenuItem } = require('electron')
const path = require('path')
const { autoUpdater } = require('electron-updater')
const beamReadFile = require('./beamReadFile')
const columnReadFile = require('./columnReadFile')
const boardReadFile = require('./boardReadFile')
const wallReadFile = require('./wallReadFile')
const isDev = require('electron-is-dev')

// const ExpirationDate = '2022/11/01'

ipcMain.handle('beam-read-file', beamReadFile)
ipcMain.handle('column-read-file', columnReadFile)
ipcMain.handle('board-read-file', boardReadFile)
ipcMain.handle('wall-read-file', wallReadFile)

const createWindow = () => {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, './preload.js'),
    },
    icon: path.join(__dirname, './icon.png'),
  })

  if (isDev) {
    // 開發階段直接與 React 連線
    win.loadURL('http://localhost:3000/')
    // 開啟 DevTools.
    win.webContents.openDevTools()
  } else {
    // 產品階段直接讀取 React 打包好的
    win.loadFile('./build/index.html')
    autoUpdater.checkForUpdates()
  }

  // win.loadFile('index.html')
  // win.once('ready-to-show', win.show)

  const menu = new Menu()
  menu.append(
    new MenuItem({
      label: '版本',
      click: () => {
        const dialogOpts = {
          type: 'info',
          detail: `此版本為 ${app.getVersion()}`,
        }
        dialog.showMessageBox(dialogOpts)
      },
    })
  )
  // menu.append(
  //   new MenuItem({
  //     label: '使用到期日',
  //     click: () => {
  //       const dialogOpts = {
  //         type: 'info',
  //         detail: `使用到期日為 ${ExpirationDate}`,
  //       }
  //       dialog.showMessageBox(dialogOpts)
  //     },
  //   })
  // )
  menu.append(
    new MenuItem({
      label: 'Toggle Developer Tools',
      accelerator: 'Ctrl++Shift+I',
      click: () => {
        win.webContents.toggleDevTools()
      },
    })
  )

  Menu.setApplicationMenu(menu)
}

// const isValidDate = () => {
//   if (isDev) return true

//   // 如果是正式環境，檢查到期日
//   const today = new Date(new Date().toDateString())
//   return new Date(ExpirationDate) >= today
// }

app.whenReady().then(() => {
  // if (!isValidDate()) {
  //   dialog.showErrorBox('軟體已過期', '軟體已過期，請洽管理員')
  //   app.quit()
  // }

  createWindow()
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow()
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

autoUpdater.on('update-downloaded', (_event, releaseNotes, releaseName) => {
  const dialogOpts = {
    type: 'info',
    buttons: ['立即更新', '稍後更新'],
    title: '軟體更新',
    message: process.platform === 'win32' ? releaseNotes : releaseName,
    detail: '新版已下載完畢，請立即更新',
  }
  dialog.showMessageBox(dialogOpts).then((returnValue) => {
    if (returnValue.response === 0) autoUpdater.quitAndInstall()
  })
})

const { app, BrowserWindow, ipcMain, dialog, Menu, MenuItem } = require('electron')
const path = require('path')
const Excel = require('exceljs')
const { autoUpdater } = require('electron-updater')
const readFile = require('./readFile')
const isDev = require('electron-is-dev')

const ExpirationDate = '2023/07/01'

ipcMain.handle('read-file', readFile)

ipcMain.handle('insert-file', async (e, filePath) => {
  const workbook = new Excel.Workbook()
  await workbook.xlsx.readFile(filePath)

  const row = workbook.getWorksheet('鋼筋型號').getRow(4) // 30.75
  const cell = row.getCell(2) // 15

  workbook.addWorksheet('abc', { properties: { defaultRowHeight: 30, defaultColWidth: 19.625 } })
  const abcWs = workbook.getWorksheet('abc')

  // console.log(sheet.name)

  // console.log(
  //   'a',
  //   sheet.columns.forEach((item) => console.log(item.width))
  // )

  const imageArr = ['test.png']

  imageArr.forEach((item, i) => {
    const imagePath = path.join(__dirname, '../public/images', item)
    const imageId1 = workbook.addImage({
      filename: imagePath,
      extension: 'png',
    })
    abcWs.addImage(imageId1, `A1:A1`)
  })

  workbook.xlsx.writeFile('abc.xlsx').then(() => console.log('finished'))
  // dialog
  //   .showSaveDialog({
  //     defaultPath: path.join(__dirname, 'public', '料單.xlsx'),
  //     buttonLabel: '存檔',
  //     filters: [{ name: 'Excel 活頁簿', extensions: ['xlsx'] }],
  //   })
  //   .then((resolve) => {
  //     const { canceled, filePath } = resolve
  //     if (!canceled) {
  //       workbook.xlsx.writeFile(filePath)
  //     }
  //   })
})

ipcMain.handle('create-file', async (e) => {
  const webContents = e.sender
  const win = BrowserWindow.fromWebContents(webContents)

  const workbook = new Excel.Workbook()

  const sheet = workbook.addWorksheet('My Sheet')

  sheet.columns = [
    { header: '中文', key: 'id', width: 10 },
    { header: '姓名', key: 'name', width: 32 },
    { header: '你好.', key: 'DOB', width: 10 },
    { header: '圖片', key: 'image', width: 10 },
  ]

  sheet.addRows([
    { id: '', name: '', DOB: '' },
    { id: 'zxc', name: 'asdasdasd', DOB: 'asc' },
    { id: '', name: '', DOB: '' },
    { id: 'B3F樑-車牙', name: '123', DOB: '456' },
  ])

  // sheet.addRow({ id: 1, name: 'John Doe', DOB: new Date(1970, 1, 1) })
  // sheet.addTable({
  //   name: 'MyTable',
  //   ref: 'A1',
  //   columns: [
  //     { name: 'Date', totalsRowLabel: 'Totals:' },
  //     { name: 'Amount', totalsRowFormula: 'SUMIF($C5:$C10,H11,$I5:$I10)', totalsRowResult: 10 },
  //   ],
  //   rows: [
  //     [new Date('2019-07-20'), 70.1],
  //     [new Date('2019-07-21'), 70.6],
  //     [new Date('2019-07-22'), 70.1],
  //   ],
  // })

  // const imageArr = ['A.PNG', 'B.PNG']

  // imageArr.forEach((item, i) => {
  //   const imagePath = path.join(__dirname, 'public\\images', item)

  //   const imageId1 = workbook.addImage({
  //     filename: imagePath,
  //     extension: 'png',
  //   })

  //   sheet.addImage(imageId1, {
  //     tl: { col: 4.2, row: i + 6 },
  //     br: { col: 4.7, row: i + 6.9 },
  //     editAs: 'absolute',
  //   })

  //   // sheet.addImage(imageId1, `D${i + 2}:D${i + 2}`)
  // })

  // dialog
  //   .showSaveDialog({
  //     defaultPath: path.join(__dirname, '測試檔案.xlsx'),
  //     buttonLabel: '存檔',
  //     filters: [{ name: 'Excel 活頁簿', extensions: ['xlsx'] }],
  //   })
  //   .then((resolve) => {
  //     const { canceled, filePath } = resolve
  //     if (!canceled) {
  //       workbook.xlsx.writeFile(filePath)
  //     }
  //   })

  workbook.xlsx.writeBuffer().then((content) => {
    win.webContents.send('a-tag-content', content)
  })
})

const createWindow = () => {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, '../src/preload.js'),
    },
  })
  win.loadFile('index.html')
  // win.once('ready-to-show', win.show)

  if (!isDev) {
    autoUpdater.checkForUpdates()
  } else {
    win.webContents.openDevTools()
  }
}

const menu = new Menu()
menu.append(
  new MenuItem({
    label: '查看版本',
    click: () => {
      const dialogOpts = {
        type: 'info',
        detail: `此版本為 ${app.getVersion()}`,
      }
      dialog.showMessageBox(dialogOpts)
    },
  })
)

Menu.setApplicationMenu(menu)

const isValidDate = () => {
  if (isDev) return true

  // 如果是正式環境，檢查到期日
  const today = new Date(new Date().toDateString())
  return new Date(ExpirationDate) >= today
}

app.whenReady().then(() => {
  if (!isValidDate()) {
    dialog.showErrorBox('軟體已過期', '軟體已過期，請洽管理員')
    app.quit()
  }

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

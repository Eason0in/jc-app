const { dialog, BrowserWindow } = require('electron')
const path = require('path')
const Excel = require('exceljs')

const {
  numMap,
  rowInit,
  carTeethHasLenA,
  carTeethMap,
  COF,
  lineNightObj,
  lineTwenSixObj,
  lineTwenSevenObj,
} = require('../src/data')
const {
  createOuterBorder,
  cellCenterStyle,
  getStatsObj,
  handleStringSum,
  getSumRow,
  handleSort,
  handleOthersSort,
  commaStyle,
  othersFormula,
} = require('../src/util')

const sheetNameObj = {
  car: '車牙',
  others: '箍筋',
}

module.exports = async (e, data) => {
  try {
    const webContents = e.sender
    const win = BrowserWindow.fromWebContents(webContents)
    const { filePath } = data
    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(filePath)
    const buildName = 'test'
    // const floorName = workbook.getWorksheet('統計').getCell('D14').value
    // const lineNightObj = getStatsObj(workbook.getWorksheet('統計'), 4, 9);
    // const lineTwenSixObj = getStatsObj(workbook.getWorksheet('統計'), 25, 26);
    // const lineTwenSevenObj = getStatsObj(workbook.getWorksheet('統計'), 25, 27);

    //#region 資料集
    let ws = ''

    // car 車牙料 others 箍筋
    const sheetArr = { car: [], others: [] }
    const sheetObj = { car: {}, others: {} }

    //#endregion

    // 將資料塞到 sheetObj
    const handleSheet = () => {
      const values = ws.getColumn(3).values
      const regex = /([A-Z]{1,2})(#?[0-9]{1,2})-(\d+)(\*\d+)?(\*\d+)?=(\d+([\+|x]\d+)?)/gi
      const { car, others } = sheetObj
      values.forEach((value) => {
        value.replace(regex, (match, tNo, num, p1, p2 = '', p3 = '', count) => {
          const obj = {
            tNo,
            num,
            count: Function('return ' + count.replace(/x/i, '*'))(),
            lenB: '',
            lenA: '',
            lenC: '',
          }
          if (p3) {
            // 70*250*80 A*B*C
            obj.lenA = p1
            obj.lenB = p2.replace('*', '')
            obj.lenC = p3.replace('*', '')
          } else if (p2) {
            // 70*250 A*B
            obj.lenA = p1
            obj.lenB = p2.replace('*', '')
          } else {
            // 80 B
            obj.lenB = p1
          }

          const { lenB, lenA, lenC } = obj

          const key = `${num}_${tNo}_${lenA}_${lenB}_${lenC}`

          if (numMap.get(tNo) === '車牙料') {
            // 先補 CD CE FC 有長度A 讀統計 sheet line 9 對應 #X
            if (carTeethHasLenA.includes(tNo)) {
              obj.lenA = lineNightObj[num]
            }
            // 計算總長度 長度B+長度A+ 有5有10
            obj.tLen = handleStringSum(lenB, lenA, carTeethMap.get(tNo)) || 0
            if (car[key]) {
              car[key].count += obj.count
            } else {
              car[key] = obj
            }
          } else {
            // 箍筋
            obj.tLen = othersFormula(tNo, lenA, lenB, lenC, num)

            if (others[key]) {
              others[key].count += obj.count
            } else {
              others[key] = obj
            }
          }
        })
      })
    }

    const setSheetToWB = (workbook) => {
      Object.entries(sheetArr).forEach(async ([key, dataArr]) => {
        const sheet = workbook.addWorksheet(`柱-${sheetNameObj[key]}`) //在檔案中新增工作表

        //#region step1 定義欄位
        sheet.columns = [
          { header: '編號', key: 'no', width: 5.625, style: cellCenterStyle },
          { header: '組編號', key: 'tNo', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '號數', key: 'num', width: 5.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '長A', key: 'lenA', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
          {
            header: '型狀/長度B',
            key: 'lenB',
            width: 15.625,
            style: { ...cellCenterStyle, ...commaStyle },
          },
          { header: '長C', key: 'lenC', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '總長度', key: 'tLen', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '支數', key: 'count', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '重量', key: 'weight', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '備註', key: 'remark', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
        ]

        //#endregion

        //#region step2 把內容資料先放進去並設定 border
        // 將rows 加入 sheet
        sheet.addRows(dataArr)

        // 設定 border
        sheet.eachRow(function (row, rowNumber) {
          if (rowNumber === 1) {
            // 第一列 編號 組編號... 上下都要有
            row.border = {
              top: { style: 'thin' },
              bottom: { style: 'thin' },
            }
          } else if (rowNumber % 2) {
            row.border = {
              bottom: { style: 'thin' },
            }
            row.height = 25.2 // 有圖片那行 20.1
          }
          // 內容左右欄位要線
          row.border = {
            ...row.border,
            left: { style: 'thin' },
            right: { style: 'thin' },
          }
        })

        //#endregion

        //#region step3 insert header
        const headerRows = [
          { ...rowInit, lenA: buildName },
          rowInit,
          { ...rowInit, no: `料單內容 ： 柱-${sheetNameObj[key]}` },
        ]

        sheet.insertRows(1, headerRows)

        sheet.mergeCells('D1', 'H1')
        sheet.mergeCells('A3', 'D3')
        sheet.getRow(1).height = 25.5

        //#endregion

        //#region step4 把內容的 image 補進去
        dataArr.forEach(({ imageName }, i) => {
          if (i % 2) {
            const imagePath = path.join(__dirname, '../public/images', imageName)
            const image = workbook.addImage({
              filename: imagePath,
              extension: 'png',
            })

            sheet.addImage(image, `E${i + 5}:E${i + 5}`)
          }
        })
        //#endregion

        //#region  step5 add footer
        const sumRow = getSumRow(dataArr)
        const footerRows = [
          rowInit,
          rowInit,
          {
            no: '材質',
            tNo: '#3',
            num: '#4',
            lenA: '#5',
            lenB: '#6',
            lenC: '#7',
            tLen: '#8',
            count: '#10',
            weight: '#11',
            remark: '合計KG',
          },
          sumRow,
        ]

        // 將rows 加入 sheet
        sheet.addRows(footerRows)

        //#endregion

        //#region step6  整張外框 + footer border
        const lastRow = sheet.lastRow._number
        const lastColumn = sheet.lastColumn._number

        // 6-1 設定 footer border
        const lastSecRow = lastRow - 1
        sheet.getRow(lastSecRow).border = {
          top: { style: 'medium' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
          left: { style: 'thin' },
        }

        // 6-1-1 設定 footer 最後一行左右 border
        sheet.getRow(lastRow).border = {
          right: { style: 'thin' },
          left: { style: 'thin' },
        }

        // 6-2 設定整張外框 border
        createOuterBorder(sheet, { row: 1, col: 1 }, { row: lastRow, col: lastColumn })

        //#endregion

        // 分頁符號
        // sheet.lastRow.addPageBreak()

        // // 列印範圍
        // sheet.pageSetup.printArea = 'A1:J12'
      })
    }

    const handleSortSheetObjToSheetArr = () => {
      Object.entries(sheetObj).forEach(([key, arr]) => {
        const fillWeightArr = Object.values(arr).map((value) => {
          const { num, count, tLen } = value
          const weight = Math.round(COF[num] * count * tLen)
          return { ...value, tLen, weight }
        })

        handleSort(fillWeightArr).forEach((value, i) => {
          const { num, tNo, count, remark = '', lenC = '', lenA = '', tLen, weight, lenB } = value
          sheetArr[key].push(
            { ...rowInit, lenB },
            {
              no: i + 1,
              tNo: numMap.get(tNo),
              num,
              lenA,
              lenB: '',
              lenC,
              tLen,
              count,
              weight,
              remark,
              imageName: `${tNo}.png`,
            }
          )
        })
      })
    }

    const handleWrite = async () => {
      const workbook = new Excel.Workbook() // 創建試算表檔案

      setSheetToWB(workbook)

      //#region 產生檔案
      workbook.xlsx.writeBuffer().then((content) => {
        win.webContents.send('column-material-file', content)
      })

      //#endregion
    }

    workbook.eachSheet((sheet, id) => {
      ws = sheet
      handleSheet()
    })

    handleSortSheetObjToSheetArr() // 排序 並放進 sheetArr
    handleWrite()
  } catch (error) {
    dialog.showErrorBox('錯誤', error.stack)
  }
}

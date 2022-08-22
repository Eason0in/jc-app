const { dialog, BrowserWindow } = require('electron')
const path = require('path')
const Excel = require('exceljs')
const isDev = require('electron-is-dev')

const { numMap, rowInit, carTeethHasLenA, carTeethMap, COF, lineNightObj } = require('./data')
const {
  createOuterBorder,
  cellCenterStyle,
  handleStringSum,
  getSumRow,
  handleColumnOthersSort,
  commaStyle,
} = require('./util')

const sheetNameObj = {
  all: '除馬椅外',
  others: '馬椅',
}

module.exports = async (e, data) => {
  try {
    const webContents = e.sender
    const win = BrowserWindow.fromWebContents(webContents)
    const { filePath, range, isNeedTidy } = data
    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(filePath)
    const buildName = 'test'

    //#region 資料集
    let ws = ''

    // all 除了箍筋之外 others 箍筋
    const sheetArr = []
    const sheetObj = { all: {}, others: {} } // 歸整前
    // const sheetTidyArr = { all: [], others: [] }

    // 歸整後
    const tidiedObj = {
      all: [],
      others: [],
    }

    //#endregion

    // 將資料塞到 sheetObj
    const handleSheet = () => {
      const rowDatas = ws.getColumn(3).values
      const rowCount = ws.getColumn(1).values
      const regex = /([A-Z]{1,2})(#?\d*~?#?\d)(@\d{1,3})?-(\d+)(\*\d+)?(\*\d+)?=(\d+([\+|x]\d+)*)([\w|\W]*)/gi
      const { all, others } = sheetObj
      rowDatas.forEach((rowData, i) => {
        const { richText } = rowData
        if (richText) {
          // 有 remark 中文字時會發生
          rowData = richText.map(({ text }) => text).join('')
        }
        // p3:@120 p8:*2 不用
        rowData.replace(regex, (match, tNo, num, p3, p4, p5, p6, count, p8, remark) => {
          let obj = {
            tNo,
            num,
            count: Function(`return  ${count.replaceAll(/x/gi, '*')} * ${rowCount[i]}`)(),
            lenB: p4,
            lenA: '',
            lenC: '',
            remark: remark && remark.match(/\(([\w|\W]*)\)/)[1],
          }

          if (p4 && p5) {
            //A*B*C
            obj.lenA = p4
            obj.lenB = p5.replace('*', '')
            obj.lenC = p6.replace('*', '')
          }

          const { lenB, lenA, lenC } = obj
          if (tNo === 'I') {
            // 箍筋 I

            const key = `${num}_${tNo}_${lenA}_${lenB}_${lenC}`
            if (others[key]) {
              others[key].count += obj.count
            } else {
              others[key] = obj
            }
          } else {
            // 先補 CD CE FC 有長度A 讀統計 sheet line 9 對應 #X
            if (carTeethHasLenA.includes(tNo)) {
              obj.lenA = lineNightObj[num]
            }

            const key = `${num}_${tNo}_${obj.lenA}_${lenB}_${lenC}`

            if (all[key]) {
              all[key].count += obj.count
            } else {
              all[key] = obj
            }
          }
        })
      })
    }

    const setSheetToWB = (workbook) => {
      const sheet = workbook.addWorksheet('板') //在檔案中新增工作表
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
      sheet.addRows(sheetArr)

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
      const headerRows = [{ ...rowInit, lenA: buildName }, rowInit, { ...rowInit, no: `料單內容 ： 板` }]

      sheet.insertRows(1, headerRows)

      sheet.mergeCells('D1', 'H1')
      sheet.mergeCells('A3', 'D3')
      sheet.getRow(1).height = 25.5

      //#endregion

      //#region step4 把內容的 image 補進去
      sheetArr.forEach(({ imageName }, i) => {
        if (i % 2) {
          const imagePath = isDev
            ? path.join(__dirname, '../public/images', imageName)
            : path.join(__dirname, './images', imageName)

          const image = workbook.addImage({
            filename: imagePath,
            extension: 'png',
          })

          sheet.addImage(image, `E${i + 5}:E${i + 5}`)
        }
      })
      //#endregion

      //#region  step5 add footer
      const sumRow = getSumRow(sheetArr)
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
    }

    const handleSortTidiedObjToSheetArr = () => {
      const togetherObj = Object.values(tidiedObj)
        .flat()
        .reduce((obj, arr) => {
          const { num, tNo, newLenB, count, lenB } = arr
          const key = `${num}_${tNo}_${tNo === 'I' ? lenB : newLenB}`
          if (obj[key]) {
            obj[key].count += count
          } else {
            obj[key] = arr
          }

          return obj
        }, {})

      const fillWeightArr = Object.values(togetherObj).map((value) => {
        const { num, newLenB, count, lenC = '', lenA = '', tNo, lenB } = value
        let tLen = handleStringSum(newLenB, lenA, lenC) || 0
        if (numMap.get(tNo) === '車牙料') {
          tLen = handleStringSum(newLenB, lenA, carTeethMap.get(tNo)) || 0
        } else if (tNo === 'I') {
          // 計算總長度
          tLen = handleStringSum(lenA * 2, lenB, lenC * 2)
        }

        const weight = Math.round(COF[num] * count * tLen)
        return { ...value, weight, tLen }
      })

      handleColumnOthersSort(fillWeightArr).forEach((value, i) => {
        const { num, tNo, count, remark = '', lenC = '', lenA = '', tLen, weight, newLenB, lenB } = value
        if (tNo === 'I') {
          sheetArr.push({ ...rowInit, lenA: '高', lenC: '腳', lenB })
        } else {
          sheetArr.push({ ...rowInit, lenB: newLenB })
        }

        sheetArr.push({
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
        })
      })
    }
    const handleSortSheetObjToSheetArr = () => {
      const fillWeightArr = Object.values(sheetObj)
        .flatMap((item) => Object.values(item))
        .map((value) => {
          const { num, count, lenC = '', lenA = '', tNo, lenB } = value
          let tLen = handleStringSum(lenB, lenA, lenC) || 0
          if (numMap.get(tNo) === '車牙料') {
            tLen = handleStringSum(lenB, lenA, carTeethMap.get(tNo)) || 0
          } else if (tNo === 'I') {
            // 計算總長度
            tLen = handleStringSum(lenA * 2, lenB, lenC * 2)
          }

          const weight = Math.round(COF[num] * count * tLen)
          return { ...value, weight, tLen }
        })

      handleColumnOthersSort(fillWeightArr).forEach((value, i) => {
        const { num, tNo, count, remark = '', lenC = '', lenA = '', tLen, weight, lenB } = value
        if (tNo === 'I') {
          sheetArr.push({ ...rowInit, lenA: '高', lenC: '腳', lenB })
        } else {
          sheetArr.push({ ...rowInit, lenB })
        }

        sheetArr.push({
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
        })
      })
    }

    const handleWrite = async () => {
      const workbook = new Excel.Workbook() // 創建試算表檔案

      setSheetToWB(workbook)

      //#region 產生檔案
      workbook.xlsx.writeBuffer().then((content) => {
        win.webContents.send('board-material-file', content)
      })

      //#endregion
    }

    const handleTidy = () => {
      Object.entries(sheetObj).forEach(([key, arr]) => {
        if (key === 'all') {
          // 用號數分類，並做排序
          const arrangeObj = Object.values(arr)
            .sort((pre, next) => next.lenB - pre.lenB)
            .sort((pre, next) => +next.num.replace('#', '') - +pre.num.replace('#', ''))
            .sort((pre, next) => next.tNo > pre.tNo)
            .reduce((obj, item) => {
              const { num, tNo } = item
              const key = `${tNo}_${num}`
              return {
                ...obj,
                [key]: obj[key] ? [...obj[key], item] : [item],
              }
            }, {})

          const tidiedArr = Object.values(arrangeObj).flatMap((subArr) => {
            let currentMax = Math.ceil(subArr[0].lenB / 10) * 10
            let nextMax = currentMax - range

            for (let i = 0; i < subArr.length; i++) {
              while (subArr[i].lenB <= currentMax) {
                if (subArr[i].lenB > nextMax) {
                  subArr[i].newLenB = currentMax
                  break
                } else {
                  currentMax = nextMax
                  nextMax -= range
                }
              }
            }
            return subArr
          })

          tidiedObj[key] = tidiedArr
        } else {
          tidiedObj[key] = Object.values(arr)
        }
      })
    }

    const handleTidyWrite = async () => {
      const workbook = new Excel.Workbook() // 創建試算表檔案
      const sheet = workbook.addWorksheet('板')
      const dataArr = Object.values(tidiedObj).flat()

      //#region step1 定義欄位
      sheet.columns = [
        { header: '代號', key: 'tNo', width: 15, style: cellCenterStyle },
        { header: '號數', key: 'num', width: 9, style: cellCenterStyle },
        { header: '原長度', key: 'lenB', width: 9, style: cellCenterStyle },
        { header: '歸整後長度', key: 'newLenB', width: 15, style: cellCenterStyle },
        { header: '支數', key: 'count', width: 15, style: cellCenterStyle },
      ]

      //#endregion

      //#region step2 把內容資料先放進去並設定 border
      // 將rows 加入 sheet
      sheet.addRows(dataArr)

      // 設定 border
      sheet.eachRow(function (row) {
        row.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        }
      })

      //#endregion

      //#region 將檔案轉成 buffer 丟到前面
      workbook.xlsx.writeBuffer().then((content) => {
        win.webContents.send('board-tidy-file', content)
      })

      //#endregion
    }

    workbook.eachSheet((sheet, id) => {
      ws = sheet
      handleSheet()
    })

    if (isNeedTidy) {
      handleTidy() // 歸整
      handleTidyWrite() // 寫歸整檔案
      handleSortTidiedObjToSheetArr() // 排序 並放進 sheetArr
    } else {
      handleSortSheetObjToSheetArr() // 排序 並放進 sheetArr
    }

    handleWrite()
  } catch (error) {
    dialog.showErrorBox('錯誤', error.stack)
  }
}

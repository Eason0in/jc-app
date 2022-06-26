const { dialog } = require('electron')
const path = require('path')
const Excel = require('exceljs')

const { numMap, obj, rowInit, sheetNameObj, carTeethHasLenA, carTeethMap, COF, AandACandCCMap } = require('../src/data')
const { createOuterBorder, cellCenterStyle, getStatsObj, handleStringSum, getSumRow } = require('../src/util')

module.exports = async (e, filePath) => {
  try {
    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(filePath)
    const buildName = workbook.getWorksheet('統計').getCell('D13').value
    const floorName = workbook.getWorksheet('統計').getCell('D14').value
    const lineNightObj = getStatsObj(workbook.getWorksheet('統計'), 4, 9)
    const lineTwenSixObj = getStatsObj(workbook.getWorksheet('統計'), 25, 26)
    const lineTwenSevenObj = getStatsObj(workbook.getWorksheet('統計'), 25, 27)

    //#region dataS
    let ws = ''

    const sheetObj = { a: [], car: [], stirrups: [] }
    const aObj = {} // 直料的鋼筋 放進去資料格式 { num, tNo, lenB, count, lenA}
    const carObj = {} // 車牙料 放進去資料格式 { num, tNo, lenB, count, lenA}
    let stirrupsObj = {} // 箍筋 -彎料 放進去資料格式 { num, tNo, lenB, lenA, count, lenC, tLen }

    //#endregion

    const handleBars = ({ s, l }) => {
      // 抓主筋 (沒有 #x 要補的數字)
      const mainBar = `#${ws.getCell('D4').value}`

      for (let i = s; i < s + l; i++) {
        const row = ws.getRow(i)
        row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
          const first = cell.value ? typeof cell.value === 'string' && cell.value.toUpperCase() : undefined
          const condF = first && numMap.has(first)

          let second = ''
          let condS = false
          // 馬椅 第二個條件要判斷分解後的 lenB lenA
          if (first === 'ID') {
            second = row.getCell(colNumber + 1).value
            const [lenB, lenA] = second.split('X')
            condS = second && lenB && lenA
          } else {
            second = row.getCell(colNumber + 1).result
            condS = second && typeof second === 'number'
          }

          const thrid = row.getCell(colNumber + 2).result || row.getCell(colNumber + 2).value
          const regex = new RegExp('x', 'gi')
          const condT = thrid && regex.test(thrid)

          if (condF && condS && condT) {
            const zero = row.getCell(colNumber - 1).result || row.getCell(colNumber - 1).value

            // const arr = [first, second, thrid.replace(/X|x/gi, '')]

            const obj = { num: mainBar, tNo: first, lenB: second, count: +thrid.replace(/X|x/gi, ''), lenA: '' }

            if (zero && ~zero.indexOf('#')) {
              // 如果前一個有 #4 就抓 沒有就補
              obj.num = zero
            }

            const key = `${obj.num}_${first}_${second}`
            if (AandACandCCMap.has(first)) {
              const isNeedRemark = i === 32 // 讀到 32 腰筋搭接 備註加 腰筋
              if (isNeedRemark) obj.remark = '腰筋'

              // AC 有長度C 讀統計 sheet line 9 對應 #X
              // CC 有長度A、C 讀統計 sheet line 9 對應 #X
              if (first === 'AC') {
                obj.lenC = lineNightObj[obj.num]
              } else if (first === 'CC') {
                obj.lenA = lineNightObj[obj.num]
                obj.lenC = lineNightObj[obj.num]
              }

              if (aObj[key]) {
                aObj[key].count += obj.count
              } else {
                aObj[key] = obj
              }
            } else if (numMap.get(first) === '車牙料') {
              // CD CE FC 有長度A 讀統計 sheet line 9 對應 #X
              if (carTeethHasLenA.includes(first)) {
                obj.lenA = lineNightObj[obj.num]
              }
              if (carObj[key]) {
                carObj[key].count += obj.count
              } else {
                carObj[key] = obj
              }
            } else if (first === 'ID') {
              // 馬椅另外處理 因為它屬於彎料但又不在 line 28-32
              const [lenB, lenA] = second.split('X')
              obj.lenB = lenB
              obj.lenA = lenA
              obj.lenC = lineNightObj[obj.num]
              // 計算總長度 馬椅算法：A*2 + C*2 +B
              obj.tLen = handleStringSum(lenA * 2, lineNightObj[obj.num] * 2, lenB)
              const key = `${obj.num}_${obj.tNo}_${lenB}_${lenA}`
              if (stirrupsObj[key]) {
                stirrupsObj[key].count += obj.count
              } else {
                stirrupsObj[key] = obj
              }
            }
          }
        })
      }
    }

    const handleStirrups = ([size, count]) => {
      const arr = []

      //#region  size
      const sizeRow = ws.getRow(size)
      sizeRow.eachCell({ includeEmpty: false }, function (cell, colNumber) {
        const first = cell.result
        const condF = first && first.length === 2 && ~first.indexOf('#')

        const second = sizeRow.getCell(colNumber + 1).result
        const secArr =
          second &&
          second
            .trim()
            .replace(/\(|\)|x/gi, '')
            .split(' ')
        const condS = second && numMap.has(secArr[0])
        if (condF && condS) {
          arr.push([first, ...secArr])
        }
      })
      //#endregion

      //#region  count

      const countRow = ws.getRow(count)
      const countArr = []
      countRow.eachCell({ includeEmpty: false }, function (cell, colNumber) {
        if (cell.result !== '箍筋數量' && cell.result) {
          countArr.push(cell.result)
        }
      })
      let sum = 0
      // 用6位去取，如果下一位是雙/三 自己要+2/3 次
      for (let j = 0; j < countArr.length; j++) {
        if (obj[countArr[j + 1]]) {
          for (let z = 1; z <= obj[countArr[j + 1]]; z++) {
            sum += countArr[j]
          }
        } else if (obj[countArr[j]]) {
          sum += 0
        } else {
          sum += countArr[j]
        }
        if (!((j + 1) % 6)) {
          arr[(j + 1) / 6 - 1].push(sum)
          sum = 0
        }
      }

      // 加蓋子
      const hatArr = Object.assign([], arr).map(([num, , lenB, , count]) => {
        const lenA = lineTwenSevenObj[num]
        const lenC = lineTwenSixObj[num]
        return [num, 'E', lenB, lenA, count, lenC]
      })

      //#endregion
      const stirrupObj = arr.reduce((obj, [num, tNo, lenB, lenA, count]) => {
        const key = `${num}_${tNo}_${lenB}_${lenA}`
        // 計算總長度 因為箍筋算法：A*2 +B + 2* 號數去讀 lineTwenSevenObj 對應 #X
        const tLen = handleStringSum(lenB, lenA * 2, lineTwenSevenObj[num] * 2)
        if (obj[key]) {
          obj[key].count += count
        } else {
          obj[key] = { num, tNo, lenB, lenA, count, tLen }
        }
        return obj
      }, {})

      const stirrupHatObj = hatArr.reduce((obj, [num, tNo, lenB, lenA, count, lenC]) => {
        const key = `${num}_${lenB}`
        // 計算總長度 因為箍筋蓋算法：A+B+C
        const tLen = handleStringSum(lenB, lenA, lenC)
        if (obj[key]) {
          obj[key].count += count
        } else {
          obj[key] = { num, tNo, lenB, lenA, count, lenC, tLen }
        }

        return obj
      }, {})

      // 要先將自己 (stirrupsObj) 放進去 因為馬椅可能已經先存進去了
      stirrupsObj = { ...stirrupsObj, ...stirrupObj, ...stirrupHatObj }
    }

    //#region 將資料塞到 aObj carObj stirrupsObj
    const handleSheet = () => {
      // 除了箍筋之外的鋼筋 讀20~27 + 32~42
      const otherBarrangeArr = [
        { s: 20, l: 8 },
        { s: 32, l: 11 },
      ]
      otherBarrangeArr.forEach(handleBars)

      // 箍筋 讀 28~31 行
      const stirrupsRangeArr = [28, 31]
      handleStirrups(stirrupsRangeArr)
    }
    //#endregion

    const tidyA = () => {
      const { a } = sheetObj
      Object.values(aObj).forEach((value, i) => {
        const { num, tNo, lenB, count, remark = '', lenC = '', lenA = '' } = value
        const tLen = handleStringSum(lenB, lenA, lenC) || 0
        const weight = Math.round(COF[num] * count * tLen)
        a.push(
          { ...rowInit, lenB: lenB },
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
    }
    const tidyCar = () => {
      const { car } = sheetObj
      Object.values(carObj).forEach((value, i) => {
        const { num, tNo, lenB, count, lenA } = value
        const tLen = handleStringSum(lenB, lenA, carTeethMap.get(tNo)) || 0
        const weight = Math.round(COF[num] * count * tLen)
        car.push(
          { ...rowInit, lenB },
          {
            no: i + 1,
            tNo: numMap.get(tNo),
            num,
            lenA,
            lenB: '',
            lenC: '',
            tLen,
            count,
            weight,
            remark: '',
            imageName: `${tNo}.png`,
          }
        )
      })
    }

    const tidyStirrups = () => {
      const { stirrups } = sheetObj
      Object.values(stirrupsObj).forEach((value, i) => {
        const { num, tNo, lenB, lenA, count, lenC, tLen } = value
        const weight = Math.round(COF[num] * count * tLen)
        stirrups.push(
          { ...rowInit, lenB: lenB },
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
            remark: '',
            imageName: `${tNo}.png`,
          }
        )
      })
    }

    const setSheetToWB = (workbook) => {
      Object.entries(sheetObj).forEach(async ([key, dataArr]) => {
        const sheet = workbook.addWorksheet(`${floorName}樑-${sheetNameObj[key]}`) //在檔案中新增工作表
        //#region step1 定義欄位
        sheet.columns = [
          { header: '編號', key: 'no', width: 9, style: cellCenterStyle },
          { header: '組編號', key: 'tNo', width: 9, style: cellCenterStyle },
          { header: '號數', key: 'num', width: 9, style: cellCenterStyle },
          { header: '長A', key: 'lenA', width: 9, style: cellCenterStyle },
          { header: '型狀/長度B', key: 'lenB', width: 15, style: cellCenterStyle },
          { header: '長C', key: 'lenC', width: 9, style: cellCenterStyle },
          { header: '總長度', key: 'tLen', width: 9, style: cellCenterStyle },
          { header: '支數', key: 'count', width: 9, style: cellCenterStyle },
          { header: '重量', key: 'weight', width: 10, style: cellCenterStyle },
          { header: '備註', key: 'remark', width: 9, style: cellCenterStyle },
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
          }
        })

        //#endregion

        //#region step3 insert header
        const headerRows = [
          { ...rowInit, lenA: buildName },
          rowInit,
          { ...rowInit, no: `料單內容 ： ${floorName}樑-${sheetNameObj[key]}` },
        ]

        sheet.insertRows(1, headerRows)

        sheet.mergeCells('D1', 'H1')
        sheet.mergeCells('A3', 'C3')
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

            sheet.addImage(image, {
              tl: { col: 4.2, row: i + 4 }, // 左上點 右上點
              br: { col: 4.9, row: i + 4.9 }, // 左下點 右下點
              editAs: 'oneCell',
            })
          }
        })
        //#endregion

        //#region  step5 add footer
        const sumRow = getSumRow(dataArr)
        const footerRows = [
          rowInit,
          rowInit,
          {
            no: '#3',
            tNo: '#4',
            num: '#5',
            lenA: '#6',
            lenB: '#7',
            lenC: '#8',
            tLen: '#10',
            count: '#11',
            weight: '合計(KG)',
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

    const handleArrtoSheetObj = () => {
      tidyA()
      tidyCar()
      tidyStirrups()
    }

    const handleWrite = async () => {
      const workbook = new Excel.Workbook() // 創建試算表檔案
      // 將 sheetObj 三個類別裡面的資料彙總 (有可能同一種資料有兩筆，筆數要加總)
      setSheetToWB(workbook)

      //#region 產生檔案

      dialog
        .showSaveDialog({
          defaultPath: path.join(__dirname, '料單.xlsx'),
          buttonLabel: '存檔',
          filters: [{ name: 'Excel 活頁簿', extensions: ['xlsx'] }],
        })
        .then((resolve) => {
          const { canceled, filePath } = resolve
          if (!canceled) {
            workbook.xlsx.writeFile(filePath)
          }
        })
      //#endregion
    }

    workbook.eachSheet((sheet, id) => {
      const nameRex = new RegExp(/^[a-zA-Z0-9]+$/gim)
      if (nameRex.test(sheet.name)) {
        ws = sheet
        handleSheet()
      }
    })

    handleArrtoSheetObj()
    handleWrite()
  } catch (error) {
    dialog.showErrorBox('錯誤', error.stack)
  }
}

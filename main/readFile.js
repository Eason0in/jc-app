const { dialog } = require('electron')
const path = require('path')
const Excel = require('exceljs')

const { numMap, obj, rowInit, sheetNameObj, carTeethHasLenA, carTeethMap, COF } = require('../src/data')
const {
  createOuterBorder,
  cellCenterStyle,
  getStatsObj,
  handleStringSum,
  getSumRow,
  handleSort,
  handleOthersSort,
  commaStyle,
} = require('../src/util')

module.exports = async (e, data) => {
  try {
    const { filePath, range } = data
    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(filePath)
    const buildName = workbook.getWorksheet('統計').getCell('D13').value
    const floorName = workbook.getWorksheet('統計').getCell('D14').value
    const lineNightObj = getStatsObj(workbook.getWorksheet('統計'), 4, 9)
    const lineTwenSixObj = getStatsObj(workbook.getWorksheet('統計'), 25, 26)
    const lineTwenSevenObj = getStatsObj(workbook.getWorksheet('統計'), 25, 27)

    //#region dataS
    let ws = ''

    const sheetObj = { a: [], car: [], stirrups: [], others: [] }

    const aObj = {} // 直料 放進去資料格式 { num, tNo, lenB, count, lenA}
    const carObj = {} // 車牙料 放進去資料格式 { num, tNo, lenB, count, lenA}
    let stirrupsObj = {} // 彎料 放進去資料格式 { num, tNo, lenB, lenA, count, lenC, tLen }

    const othersObj = {} // 箍筋, 馬椅, 斜撐, 腰筋  放進去資料格式 { num, tNo, lenB, lenA, count, lenC, tLen }

    const needTidyObj = {
      a: {},
      stirrups: {},
      car: {},
    }

    const tidiedObj = {
      a: [],
      stirrups: [],
      car: [],
    }

    //#endregion

    const handleBars = ({ s, l }) => {
      // 抓主筋 (沒有 #x 要補的數字)
      const mainBar = `#${ws.getCell('D4').value}`

      for (let i = s; i < s + l; i++) {
        const row = ws.getRow(i)
        row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
          const resultOrValue = cell.result || cell.value // 代號可能是直接英文或組成的
          const tNo = resultOrValue ? typeof resultOrValue === 'string' && resultOrValue.toUpperCase() : undefined
          const condF = tNo && numMap.has(tNo)

          const lenB = row.getCell(colNumber + 1).result
          const condS = lenB && typeof lenB === 'number'
          // }

          const count = row.getCell(colNumber + 2).result || row.getCell(colNumber + 2).value
          const regex = new RegExp('x', 'gi')
          const condT = count && regex.test(count)

          if (condF && condS && condT) {
            // 如果前一個有 #4 就抓 沒有就補
            const zero = row.getCell(colNumber - 1).result || row.getCell(colNumber - 1).value
            let num = mainBar
            if (zero && ~zero.indexOf('#')) {
              num = zero
            }

            const obj = { num, tNo, lenB, count: +count.replace(/X|x/gi, ''), lenA: '' }

            // #10_A_750
            const key = `${num}_${tNo}_${lenB}`
            const { a, stirrups, car } = needTidyObj
            if (tNo === 'A') {
              // const isNeedRemark = i === 33 // 讀到 33 腰筋搭接 備註加 腰筋
              // if (isNeedRemark) obj.remark = '腰筋'

              if (a[key]) {
                a[key].count += obj.count
              } else {
                a[key] = obj
              }
            } else if (numMap.get(tNo) === '車牙料') {
              // CD CE FC 有長度A 讀統計 sheet line 9 對應 #X
              if (carTeethHasLenA.includes(tNo)) {
                obj.lenA = lineNightObj[obj.num]
              }
              if (car[key]) {
                car[key].count += obj.count
              } else {
                car[key] = obj
              }
            } else if (numMap.get(tNo) === '彎料') {
              // AC 有長度C 讀統計 sheet line 9 對應 #X
              // CC 有長度A、C 讀統計 sheet line 9 對應 #X
              if (tNo === 'AC') {
                obj.lenC = lineNightObj[obj.num]
              } else if (tNo === 'CC') {
                obj.lenA = lineNightObj[obj.num]
                obj.lenC = lineNightObj[obj.num]
              }

              if (stirrups[key]) {
                stirrups[key].count += obj.count
              } else {
                stirrups[key] = obj
              }
            }
            // else if (tNo === 'ID') {
            //   // 馬椅另外處理 因為它屬於彎料但又不在 line 29-33
            //   const [lenB, lenA] = lenB.split('X')
            //   obj.lenB = lenB
            //   obj.lenA = lenA
            //   obj.lenC = lineNightObj[obj.num]
            //   // 計算總長度 馬椅算法：A*2 + C*2 +B
            //   obj.tLen = handleStringSum(lenA * 2, lineNightObj[obj.num] * 2, lenB)
            //   const key = `${obj.num}_${obj.tNo}_${lenB}_${lenA}`
            //   if (stirrupsObj[key]) {
            //     stirrupsObj[key].count += obj.count
            //   } else {
            //     stirrupsObj[key] = obj
            //   }
            // }
          }
        })
      }
    }

    const handleOthers = (linesArr) => {
      // 抓主筋 (沒有 #x 要補的數字)
      const mainBar = `#${ws.getCell('D4').value}`
      const [twentySeven, twentyNight, thirtyThree] = linesArr

      //#region 馬椅 斜撐 27
      const twentySevenRow = ws.getRow(twentySeven)
      twentySevenRow.eachCell({ includeEmpty: false }, function (cell, colNumber) {
        const resultOrValue = cell.result || cell.value // 代號可能是直接英文或組成的
        const tNo = resultOrValue ? typeof resultOrValue === 'string' && resultOrValue.toUpperCase() : undefined
        const condF = tNo && numMap.has(tNo)

        let second = ''
        let condS = false
        let lenB = ''
        let lenA = ''
        // 馬椅 第二個條件要判斷分解後的 lenB lenA
        if (tNo === 'ID') {
          second = twentySevenRow.getCell(colNumber + 1).result
          ;[lenB, lenA] = second.split('X')
          condS = second && lenB && lenA
        } else {
          lenB = twentySevenRow.getCell(colNumber + 1).result
          condS = lenB && typeof lenB === 'number'
        }

        const count = twentySevenRow.getCell(colNumber + 2).result || twentySevenRow.getCell(colNumber + 2).value
        const regex = new RegExp('x', 'gi')
        const condT = count && regex.test(count)

        if (condF && condS && condT) {
          // 如果前一個有 #4 就抓 沒有就補
          const zero = twentySevenRow.getCell(colNumber - 1).result || twentySevenRow.getCell(colNumber - 1).value
          let num = mainBar
          if (zero && ~zero.indexOf('#')) {
            num = zero
          }

          let key = ''
          let obj = {}
          if (tNo === 'A') {
            // 斜撐
            // #10_A_750
            key = `${num}_${tNo}_${lenB}`

            obj = { num, tNo, lenB, count: +count.replace(/X|x/gi, ''), tLen: lenB }
          } else if (tNo === 'ID') {
            // 馬椅
            const lenC = lineNightObj[num]
            // 計算總長度 馬椅算法：A*2 + C*2 +B
            const tLen = handleStringSum(lenA * 2, lineNightObj[num] * 2, lenB)

            obj = { num, tNo, lenB, count: +count.replace(/X|x/gi, ''), lenA, lenC, tLen }

            key = `${obj.num}_${obj.tNo}_${lenB}_${lenA}`
          }

          if (othersObj[key]) {
            othersObj[key].count += obj.count
          } else {
            othersObj[key] = obj
          }
        }
      })

      //#endregion

      //#region 箍筋 29
      const arr = []
      let subArr = []
      const twentyNightRow = ws.getRow(twentyNight)
      twentyNightRow.eachCell({ includeEmpty: false }, function (cell, colNumber) {
        const first = cell.result
        const condF = first && first.length === 2 && ~first.indexOf('#')

        const second = twentyNightRow.getCell(colNumber + 1).result
        const secArr =
          second &&
          second
            .toString()
            .trim()
            .replace(/\(|\)|x/gi, '')
            .split(' ')
        const condS = second && numMap.has(secArr[0])

        if (condF && condS) {
          subArr.push(first, ...secArr)
        }

        if (cell.value === '箍筋總數' && typeof second === 'number') {
          const count = second
          subArr.push(count)
          arr.push(subArr)
          subArr = []
        }
      })

      // 加蓋子
      const hatArr = Object.assign([], arr).map(([num, , lenB, , count]) => {
        const lenA = lineTwenSevenObj[num]
        const lenC = lineTwenSixObj[num]
        return [num, 'E', lenB, lenA, count, lenC]
      })

      arr.forEach(([num, tNo, lenB, lenA, count]) => {
        const key = `${num}_${tNo}_${lenB}_${lenA}`
        // 計算總長度 因為箍筋算法：A*2 +B + 2* 號數去讀 lineTwenSevenObj 對應 #X
        const tLen = handleStringSum(lenB, lenA * 2, lineTwenSevenObj[num] * 2)
        if (othersObj[key]) {
          othersObj[key].count += count
        } else {
          othersObj[key] = { num, tNo, lenB, lenA, count, tLen }
        }
        return obj
      })

      hatArr.forEach(([num, tNo, lenB, lenA, count, lenC]) => {
        const key = `${num}_${lenB}`
        // 計算總長度 因為箍筋蓋算法：A+B+C
        const tLen = handleStringSum(lenB, lenA, lenC)
        if (othersObj[key]) {
          othersObj[key].count += count
        } else {
          othersObj[key] = { num, tNo, lenB, lenA, count, lenC, tLen }
        }

        return obj
      })

      //#endregion

      //#region 腰筋 33
      const thirtyThreeRow = ws.getRow(thirtyThree)
      thirtyThreeRow.eachCell({ includeEmpty: false }, function (cell, colNumber) {
        const resultOrValue = cell.result || cell.value // 代號可能是直接英文或組成的
        const tNo = resultOrValue ? typeof resultOrValue === 'string' && resultOrValue.toUpperCase() : undefined
        const condF = tNo && numMap.has(tNo)

        const lenB = thirtyThreeRow.getCell(colNumber + 1).result
        const condS = lenB && typeof lenB === 'number'

        const count = thirtyThreeRow.getCell(colNumber + 2).result || thirtyThreeRow.getCell(colNumber + 2).value
        const regex = new RegExp('x', 'gi')
        const condT = count && regex.test(count)

        if (condF && condS && condT) {
          // 如果前一個有 #4 就抓 沒有就補
          const zero = thirtyThreeRow.getCell(colNumber - 1).result || thirtyThreeRow.getCell(colNumber - 1).value
          let num = mainBar
          if (zero && ~zero.indexOf('#')) {
            num = zero
          }

          const obj = { num, tNo, lenB, count: +count.replace(/X|x/gi, ''), remark: '腰筋', tLen: lenB }

          // #10_A_750
          const key = `${num}_${tNo}_${lenB}`
          if (othersObj[key]) {
            othersObj[key].count += obj.count
          } else {
            othersObj[key] = obj
          }
        }
      })
      //#endregion
    }

    //#region 將資料塞到 aObj carObj stirrupsObj
    const handleSheet = () => {
      // // 除了箍筋之外的鋼筋 讀20~28 + 33~43
      // const otherBarrangeArr = [
      //   { s: 20, l: 9 },
      //   { s: 33, l: 11 },
      // ]
      // otherBarrangeArr.forEach(handleBars)

      // 20~26 , 34~43 直 彎 車
      const needTidyArr = [
        { s: 20, l: 7 },
        { s: 34, l: 10 },
      ]
      needTidyArr.forEach(handleBars)

      //  27-馬椅 斜撐 29-箍筋 33-腰筋 讀 27 29 33 行
      const othersRangeArr = [27, 29, 33]
      handleOthers(othersRangeArr)
    }
    //#endregion

    const handleTidy = () => {
      Object.entries(needTidyObj).forEach(([key, arr]) => {
        // 用號數分類，並做排序
        const arrangeObj = Object.values(arr)
          .sort((pre, next) => next.lenB - pre.lenB)
          .sort((pre, next) => (next.num > pre.num ? -1 : 1))
          .reduce((obj, item) => {
            const { num } = item
            return {
              ...obj,
              [num]: obj[num] ? [...obj[num], item] : [item],
            }
          }, {})

        const tidiedArr = Object.values(arrangeObj).map((subArr) => {
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
      })
    }

    const setA = () => {
      const { a } = tidiedObj
      const togetherObj = a.flat().reduce((obj, value) => {
        const { num, tNo, newLenB, count } = value
        const key = `${num}_${tNo}_${newLenB}`
        if (obj[key]) {
          obj[key].count += count
        } else {
          obj[key] = value
        }
        return obj
      }, {})

      const aFillTLenArr = Object.values(togetherObj).map((value) => {
        const { num, newLenB, count, lenC = '', lenA = '' } = value
        const tLen = handleStringSum(newLenB, lenA, lenC) || 0
        const weight = Math.round(COF[num] * count * tLen)
        return { ...value, tLen, weight }
      })

      handleSort(aFillTLenArr).forEach((value, i) => {
        const { num, tNo, newLenB, count, remark = '', lenC = '', lenA = '', tLen, weight } = value
        sheetObj.a.push(
          { ...rowInit, lenB: newLenB },
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

    const setCar = () => {
      const { car } = tidiedObj
      const togetherObj = car.flat().reduce((obj, value) => {
        const { num, tNo, newLenB, count } = value
        const key = `${num}_${tNo}_${newLenB}`
        if (obj[key]) {
          obj[key].count += count
        } else {
          obj[key] = value
        }
        return obj
      }, {})
      const carFillTLenArr = Object.values(togetherObj).map((value) => {
        const { num, tNo, newLenB, count, lenA } = value
        const tLen = handleStringSum(newLenB, lenA, carTeethMap.get(tNo)) || 0
        const weight = Math.round(COF[num] * count * tLen)
        return { ...value, tLen, weight }
      })

      handleSort(carFillTLenArr).forEach((value, i) => {
        const { num, tNo, newLenB, count, lenA, tLen, weight } = value

        sheetObj.car.push(
          { ...rowInit, lenB: newLenB },
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

    const setStirrups = () => {
      const { stirrups } = tidiedObj
      const togetherObj = stirrups.flat().reduce((obj, value) => {
        const { num, tNo, newLenB, count } = value
        const key = `${num}_${tNo}_${newLenB}`
        if (obj[key]) {
          obj[key].count += count
        } else {
          obj[key] = value
        }
        return obj
      }, {})
      const stirrupsFillTLenArr = Object.values(togetherObj).map((value) => {
        const { num, count, newLenB, lenC = '', lenA = '' } = value
        const tLen = handleStringSum(newLenB, lenA, lenC) || 0
        const weight = Math.round(COF[num] * count * tLen)
        return { ...value, tLen, weight }
      })
      handleSort(stirrupsFillTLenArr).forEach((value, i) => {
        const { num, tNo, newLenB, lenA, count, lenC, tLen, weight } = value
        sheetObj.stirrups.push(
          { ...rowInit, lenB: newLenB },
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

    const setOthers = () => {
      const othersFillTLenArr = Object.values(othersObj).map((value) => {
        const { num, tLen, count } = value
        const weight = Math.round(COF[num] * count * tLen)
        return { ...value, weight }
      })
      handleOthersSort(othersFillTLenArr).forEach((value, i) => {
        const { num, tNo, lenB, lenA, count, lenC, tLen, weight, remark = '' } = value
        sheetObj.others.push(
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

    const setSheetToWB = (workbook) => {
      Object.entries(sheetObj).forEach(async ([key, dataArr]) => {
        const sheet = workbook.addWorksheet(`${floorName}樑-${sheetNameObj[key]}`) //在檔案中新增工作表

        //#region step1 定義欄位
        sheet.columns = [
          { header: '編號', key: 'no', width: 5.625, style: cellCenterStyle },
          { header: '組編號', key: 'tNo', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '號數', key: 'num', width: 5.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '長A', key: 'lenA', width: 8.625, style: { ...cellCenterStyle, ...commaStyle } },
          { header: '型狀/長度B', key: 'lenB', width: 15.625, style: { ...cellCenterStyle, ...commaStyle } },
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

    const handleToSheetObj = () => {
      setA()
      setCar()
      setStirrups()
      setOthers()
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

    const handleTidyWrite = async () => {
      const workbook = new Excel.Workbook() // 創建試算表檔案
      // 將 tidiedObj 三個類別裡面的資料彙總個別產生一個 sheet
      Object.entries(tidiedObj).forEach(async ([key, dataArr]) => {
        const sheet = workbook.addWorksheet(sheetNameObj[key])
        //在檔案中新增工作表

        //#region step1 定義欄位
        sheet.columns = [
          { header: '號數', key: 'num', width: 9, style: cellCenterStyle },
          { header: '原長度', key: 'lenB', width: 9, style: cellCenterStyle },
          { header: '歸整後長度', key: 'newLenB', width: 15, style: cellCenterStyle },
          { header: '支數', key: 'count', width: 15, style: cellCenterStyle },
        ]

        //#endregion

        //#region step2 把內容資料先放進去並設定 border
        // 將rows 加入 sheet
        sheet.addRows(dataArr.flat())

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
      })

      //#region 產生檔案

      dialog
        .showSaveDialog({
          defaultPath: path.join(__dirname, '歸整.xlsx'),
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
        handleSheet() // 將資料讀進 { key: {}}
      }
    })

    handleTidy() // 歸整
    handleTidyWrite() // 寫歸整檔案

    // handleToSheetObj() // 將 tidiedObj othersObj 排序並放入 sheetObj
    // handleWrite()
  } catch (error) {
    dialog.showErrorBox('錯誤', error.stack)
  }
}

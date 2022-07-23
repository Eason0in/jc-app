const { numMap, order, lineTwenSixObj, lineTwenSevenObj, lineTwenEightObj } = require('./data')
const createOuterBorder = (worksheet, start = { row: 1, col: 1 }, end = { row: 1, col: 1 }, borderWidth = 'medium') => {
  const borderStyle = {
    style: borderWidth,
  }

  for (let i = start.row; i <= end.row; i++) {
    const leftBorderCell = worksheet.getCell(i, start.col)
    const rightBorderCell = worksheet.getCell(i, end.col)
    leftBorderCell.border = {
      ...leftBorderCell.border,
      left: borderStyle,
    }
    rightBorderCell.border = {
      ...rightBorderCell.border,
      right: borderStyle,
    }
  }

  for (let i = start.col; i <= end.col; i++) {
    const topBorderCell = worksheet.getCell(start.row, i)
    const bottomBorderCell = worksheet.getCell(end.row, i)
    topBorderCell.border = {
      ...topBorderCell.border,
      top: borderStyle,
    }
    bottomBorderCell.border = {
      ...bottomBorderCell.border,
      bottom: borderStyle,
    }
  }
}

// 欄位置中
const cellCenterStyle = {
  alignment: { vertical: 'middle', horizontal: 'center' },
}

//千位符號
const commaStyle = {
  numFmt: '#,##0',
}

const getStatsObj = (worksheet, headerColNum, contentColNum) => {
  const range = [3, 4, 5, 6, 7, 8, 9, 10, 11]
  const headerRow = worksheet.getRow(headerColNum)
  const contentRow = worksheet.getRow(contentColNum)
  const resultObj = {}

  range.forEach((col) => {
    const key = headerRow.getCell(col)
    const { value } = contentRow.getCell(col)
    resultObj[key] = value
  })

  return resultObj
}

const handleStringSum = (...args) => {
  return args.reduce((sum, num) => (sum += +num), 0)
}

const getSumRow = (dataArr) => {
  const sumObj = {
    '#3': 'tNo',
    '#4': 'num',
    '#5': 'lenA',
    '#6': 'lenB',
    '#7': 'lenC',
    '#8': 'tLen',
    '#10': 'count',
    '#11': 'weight',
  }

  const reusltObj = {
    no: '合計',
    tNo: 0,
    num: 0,
    lenA: 0,
    lenB: 0,
    lenC: 0,
    tLen: 0,
    count: 0,
    weight: 0,
    remark: 0,
  }

  let total = 0
  dataArr.forEach(({ num, weight }) => {
    if (num) {
      const key = sumObj[num]
      reusltObj[key] += weight
      total += weight
    }
  })

  reusltObj.remark = total

  return reusltObj
}

/**
 * 排序方式  程式先從第三順序開始排
  1: 組編號 直料->彎料->車牙 tNo
  2: 號數 由大到小 num
  3: 總長度 長到短 tLen
 * @param {array} objFillTLenArr 
 */
const handleSort = (objFillTLenArr) => {
  const orderList = [
    (pre, next) => next.tLen - pre.tLen,
    (pre, next) => +next.num.replace('#', '') - +pre.num.replace('#', ''),
    (pre, next) => order[numMap.get(pre.tNo)] - order[numMap.get(next.tNo)],
  ]
  for (const fun of orderList) {
    objFillTLenArr.sort(fun)
  }
  return objFillTLenArr
}

/**
 * 排序方式  程式先從第四順序開始排
  1: 組編號-中文 直料->彎料 tNo
  2: 組編號-英文 tNo A~ZZ
  3: 號數 由大到小 num
  4: 總長度 長到短 tLen
 * @param {array} objFillTLenArr 
 */
const handleOthersSort = (objFillTLenArr) => {
  const orderList = [
    (pre, next) => next.tLen - pre.tLen,
    (pre, next) => +next.num.replace('#', '') - +pre.num.replace('#', ''),
    (pre, next) => next.tNo > pre.tNo,
    (pre, next) => order[numMap.get(pre.tNo)] - order[numMap.get(next.tNo)],
  ]
  for (const fun of orderList) {
    objFillTLenArr.sort(fun)
  }
  return objFillTLenArr
}

//柱- 箍筋 繫筋 計算公式
const othersFormula = (tNo, lenA, lenB, lenC, num) => {
  switch (tNo) {
    case 'F': // (A + B) * 2 +  135*2
      return handleStringSum(lenA * 2, lenB * 2, lineTwenSevenObj[num] * 2)

    case 'GI': // (A + B) * 2 +  90*2
      return handleStringSum(lenA * 2, lenB * 2, lineTwenSixObj[num] * 2)

    case 'D': // A*2 + B +  135*2
      return handleStringSum(lenA * 2, lenB, lineTwenSevenObj[num] * 2)

    case 'GA': // A*2 + B +  180*2
      return handleStringSum(lenA * 2, lenB, lineTwenEightObj[num] * 2)

    case 'E': // B +  135 90
      return handleStringSum(lenB, lineTwenSevenObj[num], lineTwenSixObj[num])

    case 'CI': // B +  180 90
      return handleStringSum(lenB, lineTwenEightObj[num], lineTwenSixObj[num])

    case 'FE': // B +  135 *2
      return handleStringSum(lenB, lineTwenSevenObj[num] * 2)

    case 'FH': // B +  180*2
      return handleStringSum(lenB, lineTwenEightObj[num] * 2)

    case 'G': // A + B + C
      return handleStringSum(lenA, lenB, lenC)

    case 'B': // A + B + 135*2
      return handleStringSum(lenA, lenB, lineTwenSevenObj[num] * 2)

    case 'C': //  A + B + 135 90
    case 'HA':
      return handleStringSum(lenA, lenB, lineTwenSevenObj[num], lineTwenSixObj[num])

    case 'GC': //  A + B + 180 90
      return handleStringSum(lenA, lenB, lineTwenEightObj[num], lineTwenSixObj[num])

    case 'GB': //  A*2 + B + 90
      return handleStringSum(lenA * 2, lenB, lineTwenSixObj[num])

    default:
      return 0
  }
}

module.exports = {
  createOuterBorder,
  cellCenterStyle,
  getStatsObj,
  handleStringSum,
  getSumRow,
  handleSort,
  handleOthersSort,
  commaStyle,
  othersFormula,
}

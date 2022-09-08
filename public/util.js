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

/** 樑 
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

/** 樑
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

/** 柱
 * 排序方式  程式先從第四順序開始排
  1: 組編號-英文 tNo 排一起
  2: 號數 由大到小 num
  3: 總長度 長到短 tLen
 * @param {array} objFillTLenArr 
 */
const handleColumnOthersSort = (objFillTLenArr) => {
  const orderList = [
    (pre, next) => +next.tLen - +pre.tLen,
    (pre, next) => +next.num.replace('#', '') - +pre.num.replace('#', ''),
    (pre, next) => (next.tNo > pre.tNo ? -1 : 1),
  ]
  for (const fun of orderList) {
    objFillTLenArr.sort(fun)
  }
  return objFillTLenArr
}

//柱- 箍筋 繫筋 計算 tLen 和部分代號塞 長度A 長度C
const othersFormula = (tNo, num, obj) => {
  const { lenB, lenA, lenC } = obj

  switch (tNo) {
    case 'F': // (A + B) * 2 +  135*2
      return { ...obj, tLen: handleStringSum(lenA * 2, lenB * 2, lineTwenSevenObj[num] * 2) }

    case 'GI': // (A + B) * 2 +  90*2
      return { ...obj, tLen: handleStringSum(lenA * 2, lenB * 2, lineTwenSixObj[num] * 2) }

    case 'D': // A*2 + B +  135*2
      return { ...obj, tLen: handleStringSum(lenA * 2, lenB, lineTwenSevenObj[num] * 2) }

    case 'GA': // A*2 + B +  180*2
      return { ...obj, tLen: handleStringSum(lenA * 2, lenB, lineTwenEightObj[num] * 2) }

    case 'E': // B +  A135 +  C90
      return {
        ...obj,
        tLen: handleStringSum(lenB, lineTwenSevenObj[num], lineTwenSixObj[num]),
        lenA: lineTwenSevenObj[num],
        lenC: lineTwenSixObj[num],
      }

    case 'CI': // B +  C180 + A90
      return {
        ...obj,
        tLen: handleStringSum(lenB, lineTwenEightObj[num], lineTwenSixObj[num]),
        lenA: lineTwenSixObj[num],
        lenC: lineTwenEightObj[num],
      }

    case 'FE': // B +  135 *2
      return {
        ...obj,
        tLen: handleStringSum(lenB, lineTwenSevenObj[num] * 2),
        lenA: lineTwenSevenObj[num],
        lenC: lineTwenSevenObj[num],
      }

    case 'FH': // B +  180*2
      return {
        ...obj,
        tLen: handleStringSum(lenB, lineTwenEightObj[num] * 2),
        lenA: lineTwenEightObj[num],
        lenC: lineTwenEightObj[num],
      }

    case 'G': // A*2 + B
      return { ...obj, tLen: handleStringSum(lenA * 2, lenB) }

    case 'B': // A + B + 135*2
      return { ...obj, tLen: handleStringSum(lenA, lenB, lineTwenSevenObj[num] * 2) }

    case 'C': //  A + B + 135 90
    case 'HA':
      return { ...obj, tLen: handleStringSum(lenA, lenB, lineTwenSevenObj[num], lineTwenSixObj[num]) }

    case 'GC': //  A + B + 180 90
      return { ...obj, tLen: handleStringSum(lenA, lenB, lineTwenEightObj[num], lineTwenSixObj[num]) }

    case 'GB': //  A*2 + B + 90
      return { ...obj, tLen: handleStringSum(lenA * 2, lenB, lineTwenSixObj[num]) }

    default:
      return { ...obj, tLen: lenB }
  }
}

//梁- 箍筋 GA HJ計算 tLen
const beamOthersCal = ({ tNo, lenB, lenA, num }) => {
  switch (tNo) {
    case 'GA': // A*2 +B + 2* 號數去讀 lineTwenSevenObj 對應 #X
      return handleStringSum(lenB, lenA * 2, lineTwenSevenObj[num] * 2)
    case 'HJ':
      return handleStringSum(lenB * 2, lenA * 2, lineTwenSixObj[num], lineTwenEightObj[num])
    default:
      return 0
  }
}

module.exports = {
  createOuterBorder,
  cellCenterStyle,
  handleStringSum,
  getSumRow,
  handleSort,
  handleOthersSort,
  commaStyle,
  othersFormula,
  handleColumnOthersSort,
  beamOthersCal,
}

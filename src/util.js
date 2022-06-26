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
    '#3': 'no',
    '#4': 'tNo',
    '#5': 'num',
    '#6': 'lenA',
    '#7': 'lenB',
    '#8': 'lenC',
    '#10': 'tLen',
    '#11': 'count',
  }

  const reusltObj = {
    no: 0,
    tNo: 0,
    num: 0,
    lenA: 0,
    lenB: 0,
    lenC: 0,
    tLen: 0,
    count: 0,
    weight: 0,
  }

  let total = 0
  dataArr.forEach(({ num, weight }) => {
    if (num) {
      const key = sumObj[num]
      reusltObj[key] += weight
      total += weight
    }
  })

  reusltObj.weight = total

  return reusltObj
}

module.exports = { createOuterBorder, cellCenterStyle, getStatsObj, handleStringSum, getSumRow }

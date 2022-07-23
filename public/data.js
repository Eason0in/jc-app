const numArr = [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'AA',
  'AB',
  'AC',
  'AD',
  'AE',
  'AF',
  'AG',
  'AH',
  'AI',
  'AJ',
  'BA',
  'BB',
  'BC',
  'BD',
  'BE',
  'BF',
  'BG',
  'BH',
  'BI',
  'BJ',
  'CA',
  'CB',
  'CC',
  'CD',
  'CE',
  'CF',
  'CG',
  'CH',
  'CI',
  'CJ',
  'DA',
  'DB',
  'DC',
  'DD',
  'DE',
  'DF',
  'DG',
  'DH',
  'DI',
  'DJ',
  'EA',
  'EB',
  'EC',
  'ED',
  'EE',
  'EF',
  'EG',
  'EH',
  'EI',
  'EJ',
  'FA',
  'FB',
  'FC',
  'FD',
  'FE',
  'FF',
  'FG',
  'FH',
  'FI',
  'FJ',
  'GA',
  'GB',
  'GC',
  'GD',
  'GE',
  'GF',
  'GG',
  'GH',
  'GI',
  'GJ',
  'HA',
  'HB',
  'HC',
  'HD',
  'HE',
  'HF',
  'HG',
  'HH',
  'HI',
  'HJ',
  'IA',
  'IB',
  'IC',
  'ID',
  'IE',
  'IF',
  'IG',
  'IH',
  'II',
  'IJ',
  'JA',
  'JB',
  'JC',
  'JD',
  'JE',
  'JF',
  'JG',
  'JH',
  'JI',
  'JJ',
  'Kn',
  'Ln',
  'Mn',
  'Nn',
  'Pn',
  'Qn',
  'Rn',
  'Sn',
  'ZZ',
  'Tn',
  'Un',
  'Vn',
  'Wn',
  'CK',
  'CL',
]

const carTeethMap = new Map([
  ['AD', 5],
  ['AE', 5],
  ['CD', 5],
  ['CE', 5],
  ['DA', 5],
  ['EB', 5],
  ['FC', 5],
  ['DD', 10],
  ['DE', 10],
  ['ED', 10],
  ['EE', 10],
  ['FA', 10],
  ['FB', 10],
  ['FD', 10],
  ['EA', 0],
  ['BE', 0],
  ['BD', 0],
])

const carTeethHasLenA = ['CD', 'CE', 'FC']
const AMap = new Map([['A', '直料']])

const numMap = (() => {
  const resultMap = new Map()
  numArr.forEach((item) => {
    let value = ''
    if (AMap.has(item)) value = '直料'
    else if (carTeethMap.has(item)) {
      value = '車牙料'
    } else value = '彎料'
    resultMap.set(item, value)
  })
  return resultMap
})()

const rowInit = {
  no: '', //編號
  tNo: '', //組編號
  num: '', //號數
  lenA: '', //長A
  lenB: '', //'型狀/長度B'
  lenC: '', //長C
  tLen: '', //總長度
  count: '', //支數
  weight: '', //重量
  remark: '', //備註
}

const COF = {
  '#3': 0.0056,
  '#4': 0.00994,
  '#5': 0.0156,
  '#6': 0.0225,
  '#7': 0.0305,
  '#8': 0.0398,
  '#10': 0.0641,
  '#11': 0.079,
}

// 主筋彎鉤長 90 度
const lineNightObj = {
  '#3': 15,
  '#4': 20,
  '#5': 25,
  '#6': 30,
  '#7': 35,
  '#8': 40,
  '#9': 45,
  '#10': 50,
  '#11': 56,
}

// 箍筋彎鉤長 90 度
const lineTwenSixObj = {
  '#3': 10,
  '#4': 13,
  '#5': 16,
  '#6': 30,
  '#7': 35,
  '#8': 40,
  '#9': 45,
  '#10': 50,
  '#11': 56,
}

// 箍筋彎鉤長 135 度
const lineTwenSevenObj = {
  '#3': 10,
  '#4': 14,
  '#5': 17,
  '#6': 20,
  '#7': 23,
  '#8': 27,
  '#9': 30,
  '#10': 34,
  '#11': 37,
}

// 箍筋彎鉤長 180 度
const lineTwenEightObj = {
  '#3': 10,
  '#4': 14,
  '#5': 17,
  '#6': 20,
  '#7': 23,
  '#8': 27,
  '#9': 30,
  '#10': 34,
  '#11': 37,
}
const order = {
  直料: 1,
  彎料: 2,
  車牙料: 3,
}
module.exports = {
  numMap,
  rowInit,
  carTeethHasLenA,
  carTeethMap,
  COF,
  lineNightObj,
  lineTwenSixObj,
  lineTwenSevenObj,
  lineTwenEightObj,
  order,
}

//#region 樑
const beamReadInput = document.getElementById('beamReadInput')
const beamMaterialFile = document.getElementById('beamMaterialFile')
const beamTidyFile = document.getElementById('beamTidyFile')
const beamConstructionFile = document.getElementById('beamConstructionFile')

const beamBtnClear = document.getElementById('beamBtnClear')
const selectBeamRange = document.getElementById('beamRange')

beamReadInput.addEventListener('change', (e) => {
  const [file] = e.target.files
  const data = { filePath: file.path, range: +selectBeamRange.value }
  window.electronAPI.beamReadFile(data)
})

window.electronAPI.sendBeamTidyFile((event, content) => {
  const blobData = new Blob([content], {
    type: 'application/vnd.ms-excel;charset=utf-8;',
  })

  beamTidyFile.href = URL.createObjectURL(blobData)
  beamTidyFile.innerText = '歸整.xlsx'
})

window.electronAPI.sendBeamMaterialFile((event, content) => {
  const blobData = new Blob([content], {
    type: 'application/vnd.ms-excel;charset=utf-8;',
  })

  beamMaterialFile.href = URL.createObjectURL(blobData)
  beamMaterialFile.innerText = '料單.xlsx'
})

window.electronAPI.sendBeamConstructionFile((event, content) => {
  const blobData = new Blob([content], {
    type: 'application/vnd.ms-excel;charset=utf-8;',
  })

  beamConstructionFile.href = URL.createObjectURL(blobData)
  beamConstructionFile.innerText = '歸整後施工圖.xlsx'
})

beamBtnClear.addEventListener('click', (e) => {
  beamReadInput.value = ''
  beamMaterialFile.href = ''
  beamMaterialFile.innerText = ''
  beamTidyFile.href = ''
  beamTidyFile.innerText = ''
  beamConstructionFile.href = ''
  beamConstructionFile.innerText = ''
})
//#endregion

//#region 柱
const columnReadInput = document.getElementById('columnReadInput')
const columnMaterialFile = document.getElementById('columnMaterialFile')
const columnTidyFile = document.getElementById('columnTidyFile')
// const beamConstructionFile = document.getElementById('beamConstructionFile')

const columnBtnClear = document.getElementById('columnBtnClear')
// const selectBeamRange = document.getElementById('beamRange')

columnReadInput.addEventListener('change', (e) => {
  const [file] = e.target.files
  const data = { filePath: file.path }
  window.electronAPI.columnReadFile(data)
})

window.electronAPI.sendColumnTidyFile((event, content) => {
  const blobData = new Blob([content], {
    type: 'application/vnd.ms-excel;charset=utf-8;',
  })

  columnTidyFile.href = URL.createObjectURL(blobData)
  columnTidyFile.innerText = '萃取整理檔案.xlsx'
})

window.electronAPI.sendColumnMaterialFile((event, content) => {
  const blobData = new Blob([content], {
    type: 'application/vnd.ms-excel;charset=utf-8;',
  })

  columnMaterialFile.href = URL.createObjectURL(blobData)
  columnMaterialFile.innerText = '料單.xlsx'
})

// window.electronAPI.sendBeamConstructionFile((event, content) => {
//   const blobData = new Blob([content], {
//     type: 'application/vnd.ms-excel;charset=utf-8;',
//   })

//   beamConstructionFile.href = URL.createObjectURL(blobData)
//   beamConstructionFile.innerText = '歸整後施工圖.xlsx'
// })

columnBtnClear.addEventListener('click', (e) => {
  columnReadInput.value = ''
  columnMaterialFile.href = ''
  columnMaterialFile.innerText = ''
  columnTidyFile.href = ''
  columnTidyFile.innerText = ''
})
//#endregion

// //#region insertFile
// const insertFile = document.getElementById('insertInput')

// insertFile.addEventListener('change', (e) => {
//   const [file] = e.target.files
//   window.electronAPI.insertFile(file.path)
// })
// //#endregion

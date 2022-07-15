const inputFile = document.getElementById('readInput')
const materialFile = document.getElementById('materialFile')
const tidyFile = document.getElementById('tidyFile')
const constructionFile = document.getElementById('constructionFile')

// const insertFile = document.getElementById('insertInput')

const btnClear = document.getElementById('btnClear')
const selectRange = document.getElementById('range')

inputFile.addEventListener('change', (e) => {
  const [file] = e.target.files
  const data = { filePath: file.path, range: +selectRange.value }
  window.electronAPI.readFile(data)
})

// insertFile.addEventListener('change', (e) => {
//   const [file] = e.target.files
//   window.electronAPI.insertFile(file.path)
// })

window.electronAPI.sendTidyFile((event, content) => {
  const blobData = new Blob([content], {
    type: 'application/vnd.ms-excel;charset=utf-8;',
  })

  tidyFile.href = URL.createObjectURL(blobData)
  tidyFile.innerText = '歸整.xlsx'
})

window.electronAPI.sendMaterialFile((event, content) => {
  const blobData = new Blob([content], {
    type: 'application/vnd.ms-excel;charset=utf-8;',
  })

  materialFile.href = URL.createObjectURL(blobData)
  materialFile.innerText = '料單.xlsx'
})

window.electronAPI.sendConstructionFile((event, content) => {
  const blobData = new Blob([content], {
    type: 'application/vnd.ms-excel;charset=utf-8;',
  })

  constructionFile.href = URL.createObjectURL(blobData)
  constructionFile.innerText = '歸整後施工圖.xlsx'
})

btnClear.addEventListener('click', (e) => {
  inputFile.value = ''
  materialFile.href = ''
  materialFile.innerText = ''
  tidyFile.href = ''
  tidyFile.innerText = ''
  constructionFile.href = ''
  constructionFile.innerText = ''
})

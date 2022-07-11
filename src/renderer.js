const inputFile = document.getElementById('readInput')
// const insertFile = document.getElementById('insertInput')
// const btn = document.getElementById('btn')
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

// btn.addEventListener('click', (e) => {
//   window.electronAPI.createFile()
// })

btnClear.addEventListener('click', (e) => {
  inputFile.value = ''
})

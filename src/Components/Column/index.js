import { useState } from 'react'
import './index.css'

function Column() {
  const [columnReadInput, setColumnReadInput] = useState('')
  const [columnMaterialFileA, setColumnMaterialFileA] = useState('')
  const [columnMaterialFile, setColumnMaterialFile] = useState('')
  const [columnTidyFile, setColumnTidyFile] = useState('')
  const [columnTidyFileA, setColumnTidyFileA] = useState('')

  const handleClear = () => {
    setColumnReadInput('')
    setColumnMaterialFile('')
    setColumnMaterialFileA('')
    setColumnTidyFile('')
    setColumnTidyFileA('')
  }

  const handleInputChange = (e) => {
    const [file] = e.target.files
    const data = { filePath: file.path }
    window.electronAPI.columnReadFile(data)
  }

  window.electronAPI.sendColumnMaterialFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setColumnMaterialFile(URL.createObjectURL(blobData))
    setColumnMaterialFileA('料單.xlsx')
  })

  window.electronAPI.sendColumnTidyFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setColumnTidyFile(URL.createObjectURL(blobData))
    setColumnTidyFileA('萃取整理檔案.xlsx')
  })

  return (
    <section id="column">
      <h2>柱</h2>
      <div>
        <input type="file" value={columnReadInput} onChange={handleInputChange} accept=".xlsx" />
        <button onClick={handleClear}>清除</button>
        <hr />
        <label>檔案:</label>
        <ul>
          <li>
            <a href={columnTidyFile} download="萃取整理檔案.xlsx">
              {columnTidyFileA}
            </a>
          </li>
          <li>
            <a href={columnMaterialFile} download="料單.xlsx">
              {columnMaterialFileA}
            </a>
          </li>
        </ul>
      </div>
    </section>
  )
}

export default Column

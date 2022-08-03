import { useState, useRef } from 'react'
import './index.scss'

function Column() {
  const [columnMaterialFileA, setColumnMaterialFileA] = useState('')
  const [columnMaterialFile, setColumnMaterialFile] = useState('')
  const [columnTidyFile, setColumnTidyFile] = useState('')
  const [columnTidyFileA, setColumnTidyFileA] = useState('')
  const fileInputRef = useRef('')

  const handleClear = () => {
    setColumnMaterialFile('')
    setColumnMaterialFileA('')
    setColumnTidyFile('')
    setColumnTidyFileA('')
    fileInputRef.current.value = ''
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
      <input ref={fileInputRef} type="file" onChange={handleInputChange} accept=".xlsx" />
      <button onClick={handleClear}>清除檔案</button>
      <label className="fileName">檔案:</label>
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
    </section>
  )
}

export default Column

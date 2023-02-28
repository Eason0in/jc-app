import { useState } from 'react'
import './index.scss'
import Dropzone from '../Dropzone'

function Column() {
  const [columnMaterialFileA, setColumnMaterialFileA] = useState('')
  const [columnMaterialFile, setColumnMaterialFile] = useState('')
  const [columnTidyFile, setColumnTidyFile] = useState('')
  const [columnTidyFileA, setColumnTidyFileA] = useState('')
  const [fileName, setFileName] = useState('')

  const handleClear = () => {
    setColumnMaterialFile('')
    setColumnMaterialFileA('')
    setColumnTidyFile('')
    setColumnTidyFileA('')
    setFileName('')
  }

  const handleInputChange = (e) => {
    const [file] = e
    setFileName(() => file.name)
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
      <Dropzone classArr="dropZone" fileName={fileName} accept=".xlsx" handleInputChange={handleInputChange} />
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

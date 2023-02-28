import { useState } from 'react'
import './index.scss'
import Dropzone from '../Dropzone'

function Wall() {
  const [wallMaterialFileA, setWallMaterialFileA] = useState('')
  const [wallMaterialFile, setWallMaterialFile] = useState('')
  const [wallTidyFile, setWallTidyFile] = useState('')
  const [wallTidyFileA, setWallTidyFileA] = useState('')
  const [selectWallRange, setSelectWallRange] = useState(10)
  const [fileName, setFileName] = useState('')
  const [isNeedTidy, setIsNeedTidy] = useState(false)

  const handleClear = () => {
    setWallMaterialFile('')
    setWallMaterialFileA('')
    setWallTidyFile('')
    setWallTidyFileA('')
    setFileName('')
  }

  const handleInputChange = (e) => {
    const [file] = e
    setFileName(() => file.name)
    const data = { filePath: file.path, range: selectWallRange, isNeedTidy }
    window.electronAPI.wallReadFile(data)
  }
  const handleSelectChange = (e) => {
    setSelectWallRange(e.target.value)
  }

  const handleNeedTidy = () => {
    setIsNeedTidy(!isNeedTidy)
  }

  window.electronAPI.sendWallMaterialFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setWallMaterialFile(URL.createObjectURL(blobData))
    setWallMaterialFileA('料單.xlsx')
  })

  window.electronAPI.sendWallTidyFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setWallTidyFile(URL.createObjectURL(blobData))
    setWallTidyFileA('歸整.xlsx')
  })

  return (
    <section id="wall">
      <div className="needTidy">
        <label htmlFor="isNeedTidy">是否需要歸整</label>
        <input id="isNeedTidy" type="checkbox" value={isNeedTidy} onClick={handleNeedTidy} />
      </div>

      <div className="range">
        <label htmlFor="wallRange">歸整間距：</label>
        <select name="wallRange" id="wallRange" value={selectWallRange} onChange={handleSelectChange}>
          <option value="10">10</option>
          <option value="20">20</option>
          <option value="30">30</option>
          <option value="40">40</option>
          <option value="50">50</option>
        </select>
      </div>

      <button onClick={handleClear}>清除檔案</button>

      <Dropzone classArr="dropZone" fileName={fileName} accept=".xlsx" handleInputChange={handleInputChange} />

      <label className="fileName">檔案:</label>
      <ul>
        <li>
          <a href={wallTidyFile} download="歸整.xlsx">
            {wallTidyFileA}
          </a>
        </li>
        <li>
          <a href={wallMaterialFile} download="料單.xlsx">
            {wallMaterialFileA}
          </a>
        </li>
      </ul>
    </section>
  )
}

export default Wall

import { useState, useRef } from 'react'
import './index.scss'

function Wall() {
  const [wallMaterialFileA, setWallMaterialFileA] = useState('')
  const [wallMaterialFile, setWallMaterialFile] = useState('')
  const [wallTidyFile, setWallTidyFile] = useState('')
  const [wallTidyFileA, setWallTidyFileA] = useState('')
  const [selectWallRange, setSelectWallRange] = useState(10)
  const [isNeedTidy, setIsNeedTidy] = useState(false)
  const fileInputRef = useRef('')

  const handleClear = () => {
    setWallMaterialFile('')
    setWallMaterialFileA('')
    setWallTidyFile('')
    setWallTidyFileA('')
    fileInputRef.current.value = ''
  }

  const handleInputChange = (e) => {
    const [file] = e.target.files
    const data = { filePath: file.path, range: selectWallRange, isNeedTidy }
    window.electronAPI.wallReadFile(data)
  }
  const handleSelectChange = (e) => {
    setSelectWallRange(e.target.value)
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
        <input id="isNeedTidy" type="checkbox" value={isNeedTidy} onClick={() => setIsNeedTidy(!isNeedTidy)} />
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

      <input ref={fileInputRef} type="file" onChange={handleInputChange} accept=".xlsx" />

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

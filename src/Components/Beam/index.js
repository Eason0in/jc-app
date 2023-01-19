import { useState, useRef } from 'react'
import './index.scss'

function Beam() {
  const [beamMaterialFileA, setBeamMaterialFileA] = useState('')
  const [beamMaterialFile, setBeamMaterialFile] = useState('')
  const [beamTidyFile, setBeamTidyFile] = useState('')
  const [beamTidyFileA, setBeamTidyFileA] = useState('')
  const [beamConstructionFile, setBeamConstructionFile] = useState('')
  const [beamConstructionFileA, setBeamConstructionFileA] = useState('')
  const [selectBeamRange, setSelectBeamRange] = useState(10)
  const fileInputRef = useRef('')
  const [isNeedTidy, setIsNeedTidy] = useState(false)

  const handleClear = () => {
    setBeamMaterialFile('')
    setBeamMaterialFileA('')
    setBeamTidyFile('')
    setBeamTidyFileA('')
    setBeamConstructionFile('')
    setBeamConstructionFileA('')
    fileInputRef.current.value = ''
  }

  const handleInputChange = (e) => {
    const [file] = e.target.files
    const data = { filePath: file.path, range: selectBeamRange, isNeedTidy }
    window.electronAPI.beamReadFile(data)
  }
  const handleSelectChange = (e) => {
    setSelectBeamRange(e.target.value)
  }

  window.electronAPI.sendBeamMaterialFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setBeamMaterialFile(URL.createObjectURL(blobData))
    setBeamMaterialFileA('料單.xlsx')
  })

  window.electronAPI.sendBeamTidyFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setBeamTidyFile(URL.createObjectURL(blobData))
    setBeamTidyFileA('歸整.xlsx')
  })

  window.electronAPI.sendBeamConstructionFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setBeamConstructionFile(URL.createObjectURL(blobData))
    setBeamConstructionFileA('歸整後施工圖.xlsx')
  })

  return (
    <section id="beam">
      <div className="needTidy">
        <label htmlFor="isNeedTidy">是否需要歸整(CC例外)</label>
        <input id="isNeedTidy" type="checkbox" value={isNeedTidy} onClick={() => setIsNeedTidy(!isNeedTidy)} />
      </div>

      <div className="range">
        <label htmlFor="beamRange">歸整間距(CC例外)：</label>
        <select name="beamRange" id="beamRange" value={selectBeamRange} onChange={handleSelectChange}>
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
          <a href={beamTidyFile} download="歸整.xlsx">
            {beamTidyFileA}
          </a>
        </li>
        <li>
          <a href={beamMaterialFile} download="料單.xlsx">
            {beamMaterialFileA}
          </a>
        </li>
        <li>
          <a href={beamConstructionFile} download="歸整後施工圖.xlsx">
            {beamConstructionFileA}
          </a>
        </li>
      </ul>
    </section>
  )
}

export default Beam

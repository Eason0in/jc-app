import { useState } from 'react'
import './index.css'

function Beam() {
  const [beamReadInput, setBeamReadInput] = useState('')
  const [beamMaterialFileA, setBeamMaterialFileA] = useState('')
  const [beamMaterialFile, setBeamMaterialFile] = useState('')
  const [beamTidyFile, setBeamTidyFile] = useState('')
  const [beamTidyFileA, setBeamTidyFileA] = useState('')
  const [beamConstructionFile, setBeamConstructionFile] = useState('')
  const [beamConstructionFileA, setBeamConstructionFileA] = useState('')
  const [selectBeamRange, setSelectBeamRange] = useState(10)

  const handleClear = () => {
    setBeamReadInput('')
    setBeamMaterialFile('')
    setBeamMaterialFileA('')
    setBeamTidyFile('')
    setBeamTidyFileA('')
    setBeamConstructionFile('')
    setBeamConstructionFileA('')
  }

  const handleInputChange = (e) => {
    const [file] = e.target.files
    const data = { filePath: file.path, range: selectBeamRange }
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
      <h2>樑</h2>
      <div>
        <select name="beamRange" id="beamRange" value={selectBeamRange} onChange={handleSelectChange}>
          <option value="10">10</option>
          <option value="20">20</option>
          <option value="30">30</option>
          <option value="40">40</option>
          <option value="50">50</option>
        </select>

        <input type="file" value={beamReadInput} onChange={handleInputChange} accept=".xlsx" />
        <button onClick={handleClear}>清除</button>
        <hr />
        <label>檔案:</label>
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
      </div>
    </section>
  )
}

export default Beam

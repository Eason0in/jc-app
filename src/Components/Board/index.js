import { useState } from 'react'
import './index.scss'
import Dropzone from '../Dropzone'
import { SelectRange } from '../../constants'

function Board() {
  const [boardMaterialFileA, setBoardMaterialFileA] = useState('')
  const [boardMaterialFile, setBoardMaterialFile] = useState('')
  const [boardTidyFile, setBoardTidyFile] = useState('')
  const [boardTidyFileA, setBoardTidyFileA] = useState('')
  const [selectBoardRange, setSelectBoardRange] = useState(SelectRange[0])
  const [fileName, setFileName] = useState('')
  const [isNeedTidy, setIsNeedTidy] = useState(false)

  const handleClear = () => {
    setBoardMaterialFile('')
    setBoardMaterialFileA('')
    setBoardTidyFile('')
    setBoardTidyFileA('')
  }

  const handleInputChange = (e) => {
    const [file] = e
    setFileName(() => file.name)
    const data = { filePath: file.path, range: selectBoardRange, isNeedTidy }
    window.electronAPI.boardReadFile(data)
  }
  const handleSelectChange = (e) => {
    setSelectBoardRange(e.target.value)
  }

  const handleNeedTidy = () => {
    setIsNeedTidy(!isNeedTidy)
  }


  window.electronAPI.sendBoardMaterialFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setBoardMaterialFile(URL.createObjectURL(blobData))
    setBoardMaterialFileA('料單.xlsx')
  })

  window.electronAPI.sendBoardTidyFile((event, content) => {
    const blobData = new Blob([content], {
      type: 'application/vnd.ms-excel;charset=utf-8;',
    })

    setBoardTidyFile(URL.createObjectURL(blobData))
    setBoardTidyFileA('歸整.xlsx')
  })

  return (
    <section id="board">
      <div className="needTidy">
        <label htmlFor="isNeedTidy">是否需要歸整(CC例外)</label>
        <input id="isNeedTidy" type="checkbox" value={isNeedTidy} onClick={handleNeedTidy} />
      </div>

      <div className="range">
        <label htmlFor="boardRange">歸整間距：</label>
        <select name="boardRange" id="boardRange" value={selectBoardRange} onChange={handleSelectChange}>
          {SelectRange.map((range) => (
            <option key={range} value={range}>
              {range}
            </option>
          ))}
        </select>
      </div>

   
      <button onClick={handleClear}>清除檔案</button>

      <Dropzone classArr="dropZone" fileName={fileName} accept=".xlsx" handleInputChange={handleInputChange} />

      <label className="fileName">檔案:</label>
      <ul>
        <li>
          <a href={boardTidyFile} download="歸整.xlsx">
            {boardTidyFileA}
          </a>
        </li>
        <li>
          <a href={boardMaterialFile} download="料單.xlsx">
            {boardMaterialFileA}
          </a>
        </li>
      </ul>
    </section>
  )
}

export default Board

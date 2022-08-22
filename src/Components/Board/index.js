import { useState, useRef } from 'react'
import './index.scss'

function Board() {
  const [boardMaterialFileA, setBoardMaterialFileA] = useState('')
  const [boardMaterialFile, setBoardMaterialFile] = useState('')
  const [boardTidyFile, setBoardTidyFile] = useState('')
  const [boardTidyFileA, setBoardTidyFileA] = useState('')
  const [selectBoardRange, setSelectBoardRange] = useState(10)
  const [isNeedTidy, setIsNeedTidy] = useState(false)
  const fileInputRef = useRef('')

  const handleClear = () => {
    setBoardMaterialFile('')
    setBoardMaterialFileA('')
    setBoardTidyFile('')
    setBoardTidyFileA('')
    fileInputRef.current.value = ''
  }

  const handleInputChange = (e) => {
    const [file] = e.target.files
    const data = { filePath: file.path, range: selectBoardRange, isNeedTidy }
    window.electronAPI.boardReadFile(data)
  }
  const handleSelectChange = (e) => {
    setSelectBoardRange(e.target.value)
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
        <label htmlFor="isNeedTidy">是否需要歸整</label>
        <input id="isNeedTidy" type="checkbox" value={isNeedTidy} onClick={() => setIsNeedTidy(!isNeedTidy)} />
      </div>

      <div className="range">
        <label htmlFor="boardRange">歸整間距：</label>
        <select name="boardRange" id="boardRange" value={selectBoardRange} onChange={handleSelectChange}>
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
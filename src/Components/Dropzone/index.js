import React from 'react'
import { useDropzone } from 'react-dropzone'
import './index.scss'

function Dropzone({ accept, handleInputChange, fileName, classArr }) {
  const onDrop = (acceptedFiles) => handleInputChange(acceptedFiles)

  const { getRootProps, getInputProps, open } = useDropzone({
    accept,
    onDrop,
    noClick: true,
    multiple: false,
  })

  return (
    <div {...getRootProps({ className: `drop ${classArr}` })}>
      <div className="dropzone">
        <input {...getInputProps()} />
        <p>拖曳檔案到這裡或點選 開啟檔案</p>
        <button onClick={open}>開啟檔案</button>
      </div>
      <div>{fileName}</div>
    </div>
  )
}
export default Dropzone

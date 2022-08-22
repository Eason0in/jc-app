const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  beamReadFile: (filePath) => {
    ipcRenderer.invoke('beam-read-file', filePath)
  },
  sendBeamTidyFile: (callback) => ipcRenderer.on('beam-tidy-file', callback),
  sendBeamMaterialFile: (callback) => ipcRenderer.on('beam-material-file', callback),
  sendBeamConstructionFile: (callback) => ipcRenderer.on('beam-construction-file', callback),

  columnReadFile: (filePath) => {
    ipcRenderer.invoke('column-read-file', filePath)
  },
  sendColumnMaterialFile: (callback) => ipcRenderer.on('column-material-file', callback),

  sendColumnTidyFile: (callback) => ipcRenderer.on('column-tidy-file', callback),
  boardReadFile: (filePath) => {
    ipcRenderer.invoke('board-read-file', filePath)
  },
  sendBoardTidyFile: (callback) => ipcRenderer.on('board-tidy-file', callback),
  sendBoardMaterialFile: (callback) => ipcRenderer.on('board-material-file', callback),
})

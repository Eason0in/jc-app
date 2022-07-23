const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  beamReadFile: (filePath) => {
    ipcRenderer.invoke('beam-read-file', filePath)
  },
  sendBeamTidyFile: (callback) => ipcRenderer.on('beam-tidy-file', callback),
  sendBeamMaterialFile: (callback) => ipcRenderer.on('beam-material-file', callback),
  sendBeamConstructionFile: (callback) => ipcRenderer.on('beam-construction-file', callback),

  insertFile: (filePath) => {
    ipcRenderer.invoke('insert-file', filePath)
  },

  columnReadFile: (filePath) => {
    ipcRenderer.invoke('column-read-file', filePath)
  },
  sendColumnMaterialFile: (callback) => ipcRenderer.on('column-material-file', callback),
  sendColumnTidyFile: (callback) => ipcRenderer.on('column-tidy-file', callback),
})

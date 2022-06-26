const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  readFile: (filePath) => {
    ipcRenderer.invoke('read-file', filePath)
  },
  insertFile: (filePath) => {
    ipcRenderer.invoke('insert-file', filePath)
  },
  createFile: () => {
    ipcRenderer.invoke('create-file')
  },
})

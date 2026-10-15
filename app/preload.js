// 工具箱壳 · preload：渲染进程唯一桥梁（contextIsolation 开启）
const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('toolbox', {
  // 文件与路径
  selectFiles: () => ipcRenderer.invoke('dlg:files'),
  selectFolder: () => ipcRenderer.invoke('dlg:folder'),
  saveDialog: (defaultName) => ipcRenderer.invoke('dlg:save', defaultName),
  showInFolder: (p) => ipcRenderer.invoke('open:show', p),
  pathForFile: (file) => webUtils.getPathForFile(file),   // Electron 33: File.path 已移除
  // 引擎
  scan: (payload) => ipcRenderer.invoke('engine:scan', payload),
  merge: (payload) => ipcRenderer.invoke('engine:merge', payload),
  cancel: () => ipcRenderer.send('engine:cancel'),
  onProgress: (cb) => {
    const h = (_e, msg) => cb(msg);
    ipcRenderer.on('engine:progress', h);
    return () => ipcRenderer.removeListener('engine:progress', h);
  },
  // 应用信息
  info: () => ipcRenderer.invoke('app:info'),
});

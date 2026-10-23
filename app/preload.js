// 工具箱壳 · preload：渲染进程唯一桥梁（contextIsolation 开启）
const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('toolbox', {
  // 文件与路径
  selectFiles: () => ipcRenderer.invoke('dlg:files'),
  selectFolder: () => ipcRenderer.invoke('dlg:folder'),
  expandPaths: (paths) => ipcRenderer.invoke('paths:expand', paths),
  saveDialog: (defaultName) => ipcRenderer.invoke('dlg:save', defaultName),
  selectOutputFolder: () => ipcRenderer.invoke('dlg:output-folder'),
  showInFolder: (p) => ipcRenderer.invoke('open:show', p),
  pathForFile: (file) => webUtils.getPathForFile(file),   // Electron 33: File.path 已移除
  // 引擎
  scan: (payload) => ipcRenderer.invoke('engine:scan', payload),
  merge: (payload) => ipcRenderer.invoke('engine:merge', payload),
  preview: (payload) => ipcRenderer.invoke('engine:preview', payload),
  mergePreview: (payload) => ipcRenderer.invoke('engine:merge_preview', payload),
  splitPreview: (payload) => ipcRenderer.invoke('engine:split_preview', payload),
  split: (payload) => ipcRenderer.invoke('engine:split', payload),
  reconcilePreview: (payload) => ipcRenderer.invoke('engine:reconcile_preview', payload),
  reconcile: (payload) => ipcRenderer.invoke('engine:reconcile', payload),
  cancel: () => ipcRenderer.send('engine:cancel'),
  onProgress: (cb) => {
    const h = (_e, msg) => cb(msg);
    ipcRenderer.on('engine:progress', h);
    return () => ipcRenderer.removeListener('engine:progress', h);
  },
  // 窗口控制
  winControl: (action) => ipcRenderer.send('win:ctl', action),
  setBackground: (color) => ipcRenderer.send('win:set-bg', color),
  // 应用信息
  info: () => ipcRenderer.invoke('app:info'),
  setAngle: (value) => ipcRenderer.invoke('gpu:set-angle', value),
});

// 预加载脚本：只把需要的几个能力暴露给页面（window.xioffice），页面本身不能访问 Node
const { contextBridge, ipcRenderer, webUtils } = require('electron');

const listen = (channel) => (cb) => {
  const h = (_e, arg) => cb(arg);
  ipcRenderer.on(channel, h);
  return () => ipcRenderer.removeListener(channel, h);
};

contextBridge.exposeInMainWorld('xioffice', {
  platform: process.platform,
  readFile: (p) => ipcRenderer.invoke('fs:read', p),
  writeFile: (p, data) => ipcRenderer.invoke('fs:write', p, data),
  listDir: (dir) => ipcRenderer.invoke('fs:list', dir),
  exportFile: (dir, name, data) => ipcRenderer.invoke('fs:export', dir, name, data),
  showOpen: (opts) => ipcRenderer.invoke('dialog:open', opts),
  showSave: (opts) => ipcRenderer.invoke('dialog:save', opts),
  pickFolder: (title) => ipcRenderer.invoke('dialog:folder', title),
  openPath: (p) => ipcRenderer.invoke('shell:open-path', p),
  openUrl: (url) => ipcRenderer.invoke('shell:open-url', url),
  fontFaces: () => ipcRenderer.invoke('fonts:faces'),
  info: () => ipcRenderer.invoke('app:info'),
  takePendingFiles: () => ipcRenderer.invoke('app:take-pending'),
  setDirty: (v) => ipcRenderer.send('app:set-dirty', v),
  /** 拖进窗口的文件在磁盘上的路径 */
  pathOf: (file) => webUtils.getPathForFile(file),
  onOpenFiles: listen('open-files'),
  onMenu: listen('menu'),
});

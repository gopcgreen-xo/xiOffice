// Electron 主进程：窗口、菜单、文件读写和系统对话框（渲染进程通过 preload.cjs 暴露的 window.xioffice 调用）
const { app, BrowserWindow, Menu, dialog, ipcMain, net, protocol, shell } = require('electron');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { allFaces } = require('./fonts.cjs');

const DEV_URL = process.env.XIOFFICE_DEV_URL;
const DIST = path.join(__dirname, '..', 'dist');
const ISSUES = 'https://github.com/gopcgreen-xo/xiOffice/issues';
const HOME = 'https://github.com/gopcgreen-xo/xiOffice';

/** 编辑器能打开的扩展名（命令行 / 双击文件打开时过滤） */
const OPEN_EXTS = ['docx', 'doc', 'dot', 'xlsx', 'xlsm', 'xls', 'csv'];
/** 允许写入的扩展名：编辑器只会保存 / 导出这些类型 */
const WRITE_EXTS = new Set([...OPEN_EXTS, 'pdf', 'png', 'jpg', 'jpeg', 'txt', 'html', 'md', 'json', 'wbp']);

const extOf = (p) => path.extname(p).slice(1).toLowerCase();

// 打包后的页面从 app:// 加载（file:// 下不能起模块 Worker）
protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true } }]);

if (!app.requestSingleInstanceLock()) app.quit();

/** 渲染进程准备好之前要打开的文件 */
let pending = [];
let win = null;
let dirty = false;

function filesFromArgv(argv) {
  return argv
    .slice(app.isPackaged ? 1 : 2)
    .filter((a) => !a.startsWith('-') && OPEN_EXTS.includes(extOf(a)))
    .map((a) => path.resolve(a))
    .filter((p) => fs.existsSync(p));
}

function openFiles(paths) {
  if (!paths.length) return;
  if (win && !win.webContents.isLoading()) win.webContents.send('open-files', paths);
  else pending.push(...paths);
}

function createWindow() {
  win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 800,
    minHeight: 560,
    title: 'xiOffice',
    backgroundColor: '#ffffff',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: false,
    },
  });
  win.once('ready-to-show', () => win.show());
  // 外部链接交给系统浏览器，窗口里只显示本应用的页面
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto):/i.test(url)) void shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (new URL(url).origin !== new URL(win.webContents.getURL()).origin) {
      e.preventDefault();
      if (/^https?:/i.test(url)) void shell.openExternal(url);
    }
  });
  win.on('close', (e) => {
    if (!dirty) return;
    const r = dialog.showMessageBoxSync(win, {
      type: 'warning',
      title: 'xiOffice',
      message: '有未保存的修改，确定要退出吗？',
      detail: '退出后未保存的修改会丢失（下次打开同一文件时可能可以从自动保存的草稿恢复）。',
      buttons: ['退出', '取消'],
      defaultId: 1,
      cancelId: 1,
    });
    if (r === 1) e.preventDefault();
  });
  win.on('closed', () => (win = null));
  if (DEV_URL) void win.loadURL(DEV_URL);
  else void win.loadURL('app://xioffice/index.html');
}

function send(cmd) {
  win?.webContents.send('menu', cmd);
}

function buildMenu() {
  const template = [
    {
      label: '文件',
      submenu: [
        { label: '新建 Word 文档', click: () => send('new-word') },
        { label: '新建 Excel 工作簿', click: () => send('new-sheet') },
        { label: '打开…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
        { type: 'separator' },
        { label: '退出', role: 'quit' },
      ],
    },
    {
      label: '视图',
      submenu: [
        { label: '重新加载', role: 'reload' },
        { label: '开发者工具', role: 'toggleDevTools' },
        { type: 'separator' },
        { label: '全屏', role: 'togglefullscreen' },
      ],
    },
    {
      label: '帮助',
      submenu: [
        { label: '提交问题…', click: () => void shell.openExternal(`${ISSUES}/new/choose`) },
        { label: '项目主页', click: () => void shell.openExternal(HOME) },
        { type: 'separator' },
        {
          label: '关于 xiOffice',
          click: () =>
            void dialog.showMessageBox(win, {
              title: '关于',
              message: `xiOffice ${app.getVersion()}`,
              detail: `Word / Excel 编辑器\nElectron ${process.versions.electron} · Chromium ${process.versions.chrome}\n${HOME}`,
            }),
        },
      ],
    },
  ];
  if (process.platform === 'darwin') template.unshift({ label: app.name, submenu: [{ role: 'about' }, { type: 'separator' }, { role: 'quit' }] });
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

/** 写文件：先写临时文件再改名，写到一半出错不会损坏原文件 */
async function writeAtomic(p, data) {
  if (!WRITE_EXTS.has(extOf(p))) throw new Error(`不允许写入这种类型的文件：${path.basename(p)}`);
  const tmp = `${p}.${process.pid}.tmp`;
  await fsp.writeFile(tmp, data);
  try {
    await fsp.rename(tmp, p);
  } catch (e) {
    await fsp.rm(tmp, { force: true });
    throw e;
  }
}

function pad(n) {
  return String(n).padStart(2, '0');
}

function registerIpc() {
  ipcMain.handle('fs:read', (_e, p) => fsp.readFile(p));
  ipcMain.handle('fs:write', (_e, p, data) => writeAtomic(p, data));
  ipcMain.handle('fs:list', async (_e, dir) => {
    const out = [];
    for (const e of await fsp.readdir(dir, { withFileTypes: true })) {
      const p = path.join(dir, e.name);
      try {
        const s = await fsp.stat(p);
        const m = s.mtime;
        const modified = `${m.getFullYear()}-${pad(m.getMonth() + 1)}-${pad(m.getDate())} ${pad(m.getHours())}:${pad(m.getMinutes())}`;
        out.push({ name: e.name, path: p, size: s.size, modified, isDir: s.isDirectory() });
      } catch {
        // 没有权限等：跳过
      }
    }
    return out;
  });
  ipcMain.handle('fs:export', async (_e, dir, name, data) => {
    // 从不覆盖：同名时加序号
    const ext = path.extname(name);
    const stem = name.slice(0, name.length - ext.length);
    let target = path.join(dir, name);
    for (let i = 2; fs.existsSync(target); i++) target = path.join(dir, `${stem} (${i})${ext}`);
    await writeAtomic(target, data);
    return path.basename(target);
  });
  ipcMain.handle('dialog:open', async (_e, { title, exts, multiple }) => {
    const filters = exts?.length ? [{ name: '支持的文件', extensions: exts }, { name: '所有文件', extensions: ['*'] }] : undefined;
    const r = await dialog.showOpenDialog(win, { title, filters, properties: ['openFile', ...(multiple ? ['multiSelections'] : [])] });
    return r.canceled ? [] : r.filePaths;
  });
  ipcMain.handle('dialog:folder', async (_e, title) => {
    const r = await dialog.showOpenDialog(win, { title, properties: ['openDirectory', 'createDirectory'] });
    return r.canceled ? null : r.filePaths[0];
  });
  ipcMain.handle('dialog:save', async (_e, { title, defaultPath, filters }) => {
    const r = await dialog.showSaveDialog(win, { title, defaultPath, filters: filters.map(([name, extensions]) => ({ name, extensions })) });
    return r.canceled || !r.filePath ? null : r.filePath;
  });
  ipcMain.handle('shell:open-path', async (_e, p) => {
    const err = await shell.openPath(p);
    if (err) throw new Error(err);
  });
  ipcMain.handle('shell:open-url', (_e, url) => {
    if (!/^(https?|mailto):/i.test(url)) throw new Error('只能打开网页链接');
    return shell.openExternal(url);
  });
  ipcMain.handle('fonts:faces', () => allFaces());
  ipcMain.handle('app:info', () => ({ documents: app.getPath('documents'), version: app.getVersion() }));
  ipcMain.handle('app:take-pending', () => pending.splice(0));
  ipcMain.on('app:set-dirty', (_e, v) => (dirty = !!v));
}

app.on('second-instance', (_e, argv) => {
  if (win) {
    if (win.isMinimized()) win.restore();
    win.focus();
  }
  openFiles(filesFromArgv(argv));
});

// macOS：双击文件 / 拖到 Dock 图标
app.on('open-file', (e, p) => {
  e.preventDefault();
  openFiles([p]);
});

app.whenReady().then(() => {
  protocol.handle('app', (req) => {
    const file = path.normalize(path.join(DIST, decodeURIComponent(new URL(req.url).pathname)));
    if (!file.startsWith(DIST + path.sep)) return new Response('Not found', { status: 404 });
    return net.fetch(pathToFileURL(file).toString());
  });
  pending.push(...filesFromArgv(process.argv));
  registerIpc();
  buildMenu();
  createWindow();
  app.on('activate', () => BrowserWindow.getAllWindows().length === 0 && createWindow());
});

app.on('window-all-closed', () => process.platform !== 'darwin' && app.quit());

/**
 * 编辑器宿主（EditorHost）：编辑器只通过它读写文件、弹提示。
 * - 桌面版（Electron）：读写磁盘上的真实文件，打开 / 保存用系统对话框，排版用系统字体的度量。
 * - 浏览器里（npm run dev:web）：文件放在内存里（虚拟文件夹 /files），保存时下载到本机。
 * 内置示例在两种环境下都先放进内存；桌面版保存示例时会弹“另存为”。
 */
import { useSyncExternalStore } from 'react';
import { basename, dirname, extOf, printImages, setEditorHost, type DirEntry, type EditorHost } from '@wz/editor-shared';
import wordPkg from '../editors/word/package.json';
import sheetPkg from '../editors/sheet/package.json';
import { desktop } from './desktop';
import { getState, openTab, patchTab, popDialog, pushDialog, removeTab, setUserName, showToast, subscribe, useStore, type Kind } from './store';

/** 内存里的虚拟文件夹 */
export const DIR = '/files';

export const EDITORS: { kind: Kind; label: string; exts: string[]; create: { name: string; title: string } }[] = [
  { kind: 'word', label: wordPkg.wz.label, exts: wordPkg.wz.exts, create: wordPkg.wz.create },
  { kind: 'sheet', label: sheetPkg.wz.label, exts: sheetPkg.wz.exts, create: sheetPkg.wz.create },
];

export const ALL_EXTS = EDITORS.flatMap((e) => e.exts);

export const kindOf = (name: string): Kind | undefined => EDITORS.find((e) => e.exts.includes(extOf(name)))?.kind;

/** 桌面版的“文档”文件夹（新建文件默认保存到这里） */
const documents = desktop ? (await desktop.info()).documents : DIR;

interface MemFile {
  data: Uint8Array;
  modified: Date;
}

const files = new Map<string, MemFile>();
const inMemory = (p: string) => files.has(p);

// ---- 文件变动通知（useFsEvent） ----
let fsTick = { tick: 0, dirs: [] as string[] };
const fsSubs = new Set<() => void>();
function touched(path: string) {
  fsTick = { tick: fsTick.tick + 1, dirs: [dirname(path)] };
  fsSubs.forEach((f) => f());
}

function store(path: string, data: Uint8Array) {
  files.set(path, { data, modified: new Date() });
  touched(path);
}

function readMem(path: string): Uint8Array {
  const f = files.get(path);
  if (!f) throw new Error(`找不到文件：${path}`);
  return f.data;
}

/** 下载到本机（浏览器里的“保存”） */
export function download(name: string, data: Uint8Array | Blob) {
  const blob = data instanceof Blob ? data : new Blob([data as BlobPart]);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** 把文件内容放进内存文件系统，返回路径（重名的后来者覆盖） */
async function importFile(file: File): Promise<string> {
  const path = `${DIR}/${file.name}`;
  store(path, new Uint8Array(await file.arrayBuffer()));
  return path;
}

/** 按路径打开（编辑器能处理的才打开） */
export function openPaths(paths: string[]) {
  for (const p of paths) {
    const kind = kindOf(p);
    if (kind) openTab({ kind, title: basename(p), path: p });
    else showToast(`不支持的文件类型：${basename(p)}`);
  }
}

/** 打开 File 对象（拖进窗口、浏览器里选的文件、内置示例）：桌面版能拿到磁盘路径时直接打开原文件 */
export async function openFiles(list: Iterable<File>) {
  for (const f of list) {
    if (!kindOf(f.name)) {
      showToast(`不支持的文件类型：${f.name}`);
      continue;
    }
    const onDisk = desktop?.pathOf(f);
    openPaths([onDisk || (await importFile(f))]);
  }
}

/** 浏览器里的文件选择框 */
function chooseFiles(exts: string[], multiple = true): Promise<File[]> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = multiple;
    input.accept = exts.map((e) => `.${e}`).join(',');
    input.onchange = () => resolve([...(input.files ?? [])]);
    input.oncancel = () => resolve([]);
    input.click();
  });
}

/** “打开文件”：桌面版弹系统对话框，浏览器里弹文件选择框 */
export async function chooseAndOpen(exts = ALL_EXTS) {
  if (desktop) openPaths(await desktop.showOpen({ title: '打开', exts, multiple: true }));
  else await openFiles(await chooseFiles(exts));
}

export function newDocument(kind: Kind) {
  const e = EDITORS.find((x) => x.kind === kind)!;
  openTab({ kind, title: e.create.title });
}

const saves = new Map<string, () => Promise<boolean>>();

/** 关标签：有未保存修改时先问 */
export async function closeTab(id: string) {
  const tab = getState().tabs.find((t) => t.id === id);
  if (!tab) return;
  if (tab.dirty) {
    const hint = desktop ? '' : '\n\n选“保存”会把文件下载到本机；';
    const save = await host.confirm(`「${tab.title}」有未保存的修改，保存吗？${hint}选“取消”放弃修改。`, '保存');
    if (save && !(await saves.get(id)?.())) return;
  }
  removeTab(id);
}

function ask(title: string, value = ''): Promise<string | null> {
  return new Promise((resolve) => {
    const d = { type: 'ask' as const, title, value, resolve: (v: string | null) => (popDialog(d), resolve(v)) };
    pushDialog(d);
  });
}

function confirmBox(message: string, ok = '确定'): Promise<boolean> {
  return new Promise((resolve) => {
    const d = { type: 'confirm' as const, message, ok, resolve: (v: boolean) => (popDialog(d), resolve(v)) };
    pushDialog(d);
  });
}

function decodeText(data: Uint8Array): { content: string; encoding: string } {
  if (data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf) return { content: new TextDecoder().decode(data.subarray(3)), encoding: 'UTF-8' };
  try {
    return { content: new TextDecoder('utf-8', { fatal: true }).decode(data), encoding: 'UTF-8' };
  } catch {
    return { content: new TextDecoder('gbk').decode(data), encoding: 'GBK' };
  }
}

const join = (dir: string, name: string) => `${dir.replace(/[\\/]+$/, '')}${dir.includes('\\') ? '\\' : '/'}${name}`;

const pickSavePath = desktop
  ? (title: string, name: string, dir: string | null, filters: [string, string[]][]) =>
      desktop!.showSave({ title, defaultPath: join(dir && !inMemory(join(dir, name)) && dir !== DIR ? dir : documents, name), filters })
  : undefined;

/** 写文件：桌面版写磁盘；内存里的文件（浏览器里的文件、示例）写回内存，浏览器里同时下载，桌面版改为“另存为” */
async function writeFile(p: string, data: Uint8Array) {
  if (desktop && !inMemory(p)) return desktop.writeFile(p, data);
  if (!desktop) {
    store(p, data);
    return download(basename(p), data);
  }
  const ext = extOf(p);
  const target = await pickSavePath!('保存', basename(p), null, [[ext.toUpperCase(), [ext]]]);
  if (!target) throw new Error('已取消保存');
  await desktop.writeFile(target, data);
  // 标签页改到磁盘上的新文件
  for (const t of getState().tabs) if (t.path === p) patchTab(t.id, { path: target, title: basename(target) });
}

export const host: EditorHost = {
  platform: 'desktop',
  readBytes: async (p) => (desktop && !inMemory(p) ? desktop.readFile(p) : readMem(p)),
  readText: async (p) => decodeText(desktop && !inMemory(p) ? await desktop.readFile(p) : readMem(p)),
  writeBytes: writeFile,
  canWrite: async () => true,
  writeText: async (p, text) => {
    // 只能编码 UTF-8：一律带 BOM 写出，Excel 打开 csv 才不乱码
    const body = new TextEncoder().encode(text);
    const data = new Uint8Array(body.length + 3);
    data.set([0xef, 0xbb, 0xbf]);
    data.set(body, 3);
    await writeFile(p, data);
  },
  listDir: async (dir) => {
    if (desktop && dir !== DIR) return desktop.listDir(dir);
    const prefix = dir.replace(/[\\/]+$/, '') + '/';
    const out: DirEntry[] = [];
    for (const [path, f] of files) {
      if (!path.startsWith(prefix) || path.slice(prefix.length).includes('/')) continue;
      const m = f.modified;
      const pad = (n: number) => String(n).padStart(2, '0');
      const modified = `${m.getFullYear()}-${pad(m.getMonth() + 1)}-${pad(m.getDate())} ${pad(m.getHours())}:${pad(m.getMinutes())}`;
      out.push({ name: basename(path), path, size: f.data.length, modified, isDir: false });
    }
    return out;
  },
  readXls: async () => {
    throw new Error('暂不支持旧版 .xls，请先用 Excel / WPS 另存为 .xlsx');
  },
  fontFaces: desktop ? () => desktop!.fontFaces() : undefined,
  openWithSystem: async (p) => (desktop && !inMemory(p) ? desktop.openPath(p) : download(basename(p), readMem(p))),
  openUrl: async (url) => (desktop ? desktop.openUrl(url) : void window.open(url, '_blank', 'noopener')),
  workFolders: () => [documents],
  toast: (msg, ms) => showToast(msg, ms),
  askText: ask,
  pickSavePath,
  confirm: confirmBox,
  guard: async (p) => {
    try {
      return await p;
    } catch (e) {
      console.error(e);
      showToast(`出错了：${e instanceof Error ? e.message : String(e)}`, 5000);
      return undefined;
    }
  },
  registerSave: (id, save) => {
    saves.set(id, save);
    return () => void saves.delete(id);
  },
  pushPopup: () => undefined,
  popPopup: () => undefined,
  updateTab: (id, patch) => patchTab(id, patch),
  printImages: (urls, size) => printImages(urls, size),
  log: (level, source, msg) => console[level](`[${source}]`, msg),
  useWorkFolders: () => [documents],
  useFsEvent: () =>
    useSyncExternalStore(
      (cb) => (fsSubs.add(cb), () => void fsSubs.delete(cb)),
      () => fsTick,
    ),
  useUserName: () => useStore((s) => s.userName),
  setUserName,
  newDocument: (kind) => newDocument(kind as Kind),
  openDocument: (kind) => chooseAndOpen(EDITORS.find((e) => e.kind === kind)?.exts),
  openPath: (path, kind, title) => {
    const k = (kind as Kind | undefined) ?? kindOf(path);
    if (k) openTab({ kind: k, title: title ?? basename(path), path });
  },
  closeTab: (id) => removeTab(id),
  pickFiles: async (title, exts) =>
    desktop ? desktop.showOpen({ title, exts, multiple: true }) : Promise.all((await chooseFiles(exts ?? [])).map(importFile)),
  pickFolder: desktop ? (title) => desktop!.pickFolder(title) : undefined,
  // 浏览器里导出直接下载，不用选文件夹
  pickExportDir: async (title) => (desktop ? desktop.pickFolder(title) : DIR),
  exportFile: async (dir, name, data) => (desktop ? desktop.exportFile(dir, name, data) : (download(name, data), name)),
};

setEditorHost(host);

if (desktop) {
  // 关窗口前由主进程询问（Electron 里 beforeunload 拦截会让窗口关不掉且没有提示）
  let last = false;
  subscribe(() => {
    const dirty = getState().tabs.some((t) => t.dirty);
    if (dirty !== last) desktop!.setDirty((last = dirty));
  });
  // 双击文件、命令行参数、第二次启动时传来的文件
  openPaths(await desktop.takePendingFiles());
  desktop.onOpenFiles(openPaths);
  desktop.onMenu((cmd) => {
    if (cmd === 'new-word') newDocument('word');
    else if (cmd === 'new-sheet') newDocument('sheet');
    else if (cmd === 'open') void chooseAndOpen();
  });
} else {
  // 有未保存的修改时，关闭 / 刷新页面前提醒
  window.addEventListener('beforeunload', (e) => {
    if (getState().tabs.some((t) => t.dirty)) e.preventDefault();
  });
}

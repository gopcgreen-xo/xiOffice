/**
 * 浏览器里的编辑器宿主（EditorHost）：编辑器只通过它读写文件、弹提示。
 * 文件放在内存里（虚拟文件夹 /files），从本机打开的文件先读进来；保存时写回内存并下载到本机。
 */
import { useSyncExternalStore } from 'react';
import { basename, extOf, printImages, setEditorHost, type DirEntry, type EditorHost } from '@wz/editor-shared';
import wordPkg from '../editors/word/package.json';
import sheetPkg from '../editors/sheet/package.json';
import { getState, openTab, patchTab, popDialog, pushDialog, removeTab, setUserName, showToast, useStore, type Kind } from './store';

/** 虚拟工作文件夹 */
export const DIR = '/files';

export const EDITORS: { kind: Kind; label: string; exts: string[]; create: { name: string; title: string } }[] = [
  { kind: 'word', label: wordPkg.wz.label, exts: wordPkg.wz.exts, create: wordPkg.wz.create },
  { kind: 'sheet', label: sheetPkg.wz.label, exts: sheetPkg.wz.exts, create: sheetPkg.wz.create },
];

export const kindOf = (name: string): Kind | undefined => EDITORS.find((e) => e.exts.includes(extOf(name)))?.kind;

interface MemFile {
  data: Uint8Array;
  modified: Date;
}

const files = new Map<string, MemFile>();

// ---- 文件变动通知（useFsEvent） ----
let fsTick = { tick: 0, dirs: [] as string[] };
const fsSubs = new Set<() => void>();
function touched(path: string) {
  fsTick = { tick: fsTick.tick + 1, dirs: [path.slice(0, path.lastIndexOf('/')) || '/'] };
  fsSubs.forEach((f) => f());
}

function store(path: string, data: Uint8Array) {
  files.set(path, { data, modified: new Date() });
  touched(path);
}

function read(path: string): Uint8Array {
  const f = files.get(path);
  if (!f) throw new Error(`找不到文件：${path}`);
  return f.data;
}

/** 下载到本机（浏览器的“保存”） */
export function download(name: string, data: Uint8Array | Blob) {
  const blob = data instanceof Blob ? data : new Blob([data as BlobPart]);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** 把本机文件读进内存文件系统，返回路径（重名的后来者覆盖） */
export async function importFile(file: File): Promise<string> {
  const path = `${DIR}/${file.name}`;
  store(path, new Uint8Array(await file.arrayBuffer()));
  return path;
}

/** 打开本机文件（编辑器能处理的才打开） */
export async function openFiles(list: Iterable<File>) {
  for (const f of list) {
    const kind = kindOf(f.name);
    if (!kind) {
      showToast(`不支持的文件类型：${f.name}`);
      continue;
    }
    openTab({ kind, title: f.name, path: await importFile(f) });
  }
}

/** 弹出文件选择框 */
export function chooseFiles(exts: string[], multiple = true): Promise<File[]> {
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
    const save = await host.confirm(`「${tab.title}」有未保存的修改，保存吗？\n\n选“保存”会把文件下载到本机；选“取消”放弃修改。`, '保存');
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

export const host: EditorHost = {
  platform: 'desktop',
  readBytes: async (p) => read(p),
  readText: async (p) => decodeText(read(p)),
  writeBytes: async (p, data) => {
    store(p, data);
    download(basename(p), data);
  },
  canWrite: async () => true,
  writeText: async (p, text) => {
    // 浏览器只能编码 UTF-8：一律带 BOM 写出，Excel 打开 csv 才不乱码
    const body = new TextEncoder().encode(text);
    const data = new Uint8Array(body.length + 3);
    data.set([0xef, 0xbb, 0xbf]);
    data.set(body, 3);
    store(p, data);
    download(basename(p), data);
  },
  listDir: async (dir) => {
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
    throw new Error('浏览器版暂不支持旧版 .xls（需要后端解析），请先用 Excel / WPS 另存为 .xlsx');
  },
  openWithSystem: async (p) => download(basename(p), read(p)),
  openUrl: async (url) => void window.open(url, '_blank', 'noopener'),
  workFolders: () => [DIR],
  toast: (msg, ms) => showToast(msg, ms),
  askText: ask,
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
  useWorkFolders: () => [DIR],
  useFsEvent: () =>
    useSyncExternalStore(
      (cb) => (fsSubs.add(cb), () => void fsSubs.delete(cb)),
      () => fsTick,
    ),
  useUserName: () => useStore((s) => s.userName),
  setUserName,
  newDocument: (kind) => newDocument(kind as Kind),
  openDocument: async (kind) => openFiles(await chooseFiles(EDITORS.find((e) => e.kind === kind)?.exts ?? [])),
  openPath: (path, kind, title) => {
    const k = (kind as Kind | undefined) ?? kindOf(path);
    if (k) openTab({ kind: k, title: title ?? basename(path), path });
  },
  closeTab: (id) => removeTab(id),
  pickFiles: async (_title, exts) => Promise.all((await chooseFiles(exts ?? [])).map(importFile)),
  // 导出直接下载，不用选文件夹
  pickExportDir: async () => DIR,
  exportFile: async (_dir, name, data) => (download(name, data), name),
};

setEditorHost(host);

// 有未保存的修改时，关闭 / 刷新页面前提醒
window.addEventListener('beforeunload', (e) => {
  if (getState().tabs.some((t) => t.dirty)) e.preventDefault();
});

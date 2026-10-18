/**
 * 工作包的“虚拟路径”：包内条目在工作台里是 `<包文件路径>!/items/<id>/<文件名>`。
 * withBundles 把宿主的读写方法包一层：路径指向包内时转给容器读写（再整包原子写回），其余原样转发，
 * 所以 word / sheet / markdown / html 等编辑器不用改，仍然“以为自己在读写一个文件”。
 * 同一个包在内存里只有一份 Bundle（多个标签页共享），写盘按包排队。
 */
import { unzipSync } from 'fflate';
import { requireLevel, requireManage, packAccess } from './access';
import { PackLockedError, encryptedKeys } from './packKeys';
import { Bundle, joinVPath, kindOfName, splitVPath, itemInner, type BundleItem } from './container';
import { basename, type BundleContext, type EditorHost } from './host';
import type { SheetColumns } from './schema';

const cache = new Map<string, Bundle>();
const loading = new Map<string, Promise<Bundle>>();
const writing = new Map<string, Promise<void>>();
let base: EditorHost | null = null;

function baseHost(): EditorHost {
  if (!base) throw new Error('工作包还没有初始化（withBundles）');
  return base;
}

/** 用户明确“展开”过的包：只含一个白板时也显示工作包页，不再自动进入白板状态（本次运行有效） */
const expanded = new Set<string>();
export const expandPack = (pack: string): void => void expanded.add(pack);
export const isPackExpanded = (pack: string): boolean => expanded.has(pack);

/**
 * “打开表单录入”的请求：工作包页点“表单录入”后，表格条目在编辑器里打开（或已经打开）时弹出表单面板。
 * 键是 `包路径|条目 id`；表格还没加载好时请求先留着，加载后取走。
 */
const formRequests = new Map<string, string>();
const formListeners = new Set<(key: string) => void>();
export const formRequestKey = (pack: string, itemId: string): string => `${pack}|${itemId}`;
export function requestForm(pack: string, itemId: string, formId = ''): void {
  const key = formRequestKey(pack, itemId);
  formRequests.set(key, formId);
  for (const cb of [...formListeners]) cb(key);
}
/** 取走一个请求（没有返回 undefined；表单 id 为空串表示默认表单） */
export function takeFormRequest(pack: string, itemId: string): string | undefined {
  const key = formRequestKey(pack, itemId);
  const v = formRequests.get(key);
  formRequests.delete(key);
  return v;
}
export function onFormRequest(cb: (key: string) => void): () => void {
  formListeners.add(cb);
  return () => void formListeners.delete(cb);
}

/** 加密包的解压密钥（只在当前窗口内存里；登录时解开） */
const packKeys = new Map<string, string>();
export const unlockPack = (pack: string, key: string): void => void packKeys.set(pack, key);
export const packKey = (pack: string): string | undefined => packKeys.get(pack);
/** 之后写盘不再加密（关闭加密时用；包里的 meta/keys.json 由调用方删掉） */
export const forgetPackKey = (pack: string): void => void packKeys.delete(pack);

/** 读出包的不加密 zip 字节：加密包要先有密钥，否则抛 PackLockedError */
async function readPlain(pack: string): Promise<Uint8Array> {
  const bytes = await baseHost().readBytes(pack);
  const keys = encryptedKeys(bytes);
  if (!keys) return bytes;
  const key = packKeys.get(pack);
  const read = baseHost().readPack;
  if (!key || !read) throw new PackLockedError(pack, keys);
  return read(pack, key);
}

/** 取得（并缓存）一个包；整个应用里同一个包只有一个 Bundle */
export async function loadBundle(pack: string): Promise<Bundle> {
  const have = cache.get(pack);
  if (have) return have;
  let p = loading.get(pack);
  if (!p) {
    p = (async () => {
      const b = Bundle.open(await readPlain(pack));
      cache.set(pack, b);
      return b;
    })().finally(() => loading.delete(pack));
    loading.set(pack, p);
  }
  return p;
}

/** 包文件被外部改了（文件夹变动事件）：没有未保存修改时重新读入 */
export async function refreshBundle(pack: string): Promise<boolean> {
  const b = cache.get(pack);
  if (!b || b.dirty) return false;
  const bytes = await readPlain(pack);
  const same = (() => {
    const cur = b.snapshot().bytes;
    if (cur.length === bytes.length && cur.every((v, i) => v === bytes[i])) return true;
    // 加密包解密后是重新打包的 zip，字节不同但内容可能一样：逐个部件比较
    if (!packKeys.has(pack)) return false;
    const a = unzipSync(cur);
    const c = unzipSync(bytes);
    const names = Object.keys(a);
    return names.length === Object.keys(c).length && names.every((n) => c[n] && a[n]!.length === c[n]!.length && a[n]!.every((v, i) => v === c[n]![i]));
  })();
  if (same) return false;
  b.reloadFrom(bytes);
  return true;
}

/** 把包的当前状态写盘（按包排队；保存期间又有改动的话下一次再写） */
export function flushBundle(pack: string): Promise<void> {
  const prev = writing.get(pack) ?? Promise.resolve();
  const next = prev
    .catch(() => undefined)
    .then(async () => {
      const b = cache.get(pack);
      if (!b || !b.dirty) return;
      const { bytes, rev } = b.snapshot();
      const key = packKeys.get(pack);
      const write = baseHost().writePack;
      // 加密包：整包加密后写盘（不能逐字节就地修补，内存里仍按不加密的 zip 修补）
      if (key && write) await write(pack, bytes, key);
      else await baseHost().writeBytes(pack, bytes);
      b.markSaved(bytes, rev);
    });
  writing.set(pack, next);
  return next;
}

/** 新建一个空包文件（已存在的文件由调用方先检查，这里不覆盖） */
export async function newBundleFile(pack: string, title = ''): Promise<Bundle> {
  const b = Bundle.create(title);
  cache.set(pack, b);
  await flushBundle(pack);
  return b;
}

/** 把磁盘上的文件加进包：文本类按编码读出再存成 UTF-8，其余原字节；返回新增的条目 */
export async function addFilesToBundle(pack: string, paths: string[]): Promise<BundleItem[]> {
  const b = await loadBundle(pack);
  requireManage(b, pack);
  const out: BundleItem[] = [];
  for (const path of paths) {
    const name = path.split(/[\\/]/).pop() ?? path;
    const kind = kindOfName(name);
    let bytes: Uint8Array;
    if (kind === 'md' || kind === 'html' || kind === 'script') {
      bytes = new TextEncoder().encode((await baseHost().readText(path)).content);
    } else {
      bytes = await baseHost().readBytes(path);
    }
    out.push(b.addItem(name, bytes));
  }
  await flushBundle(pack);
  return out;
}

const atLeastRead = (b: Bundle, pack: string, id: string): boolean => {
  const l = packAccess(b, pack).level(id);
  return l === 'read' || l === 'write';
};

/** 包内路径 `items/<id>/…` → 条目（只看 id，不看文件名） */
function itemOf(b: Bundle, inner: string): BundleItem | undefined {
  const m = /^items\/([^/]+)\//.exec(inner);
  return m ? b.item(m[1]) : undefined;
}

/** 包内路径 → 虚拟路径 */
export const itemPath = (pack: string, item: Pick<BundleItem, 'id' | 'name'>): string => joinVPath(pack, itemInner(item));

function contextFor(pack: string, item: BundleItem): BundleContext {
  const cur = () => {
    const b = cache.get(pack);
    if (!b) throw new Error('工作包已关闭');
    return b;
  };
  return {
    pack,
    itemId: item.id,
    itemName: item.name,
    get packId() {
      return cur().manifest.id;
    },
    get packTitle() {
      return cur().manifest.title || basename(pack).replace(/\.wbp$/i, '');
    },
    columns: () => cur().columns(item.id),
    setColumns(sheet: string, cols: SheetColumns | null) {
      const b = cur();
      requireLevel(b, pack, item.id, 'write');
      const file = b.columns(item.id);
      if (cols && cols.columns.length) file.sheets[sheet] = cols;
      else delete file.sheets[sheet];
      b.setColumns(item.id, file);
    },
    forms: () => cur().forms().filter((f) => f.target.item === item.id),
    async saveForm(form) {
      requireLevel(cur(), pack, item.id, 'write');
      cur().saveForm(form);
      await flushBundle(pack);
    },
    async deleteForm(id) {
      requireLevel(cur(), pack, item.id, 'write');
      cur().deleteForm(id);
      await flushBundle(pack);
    },
    newFormId: () => cur().newFormId(),
    logSubmission: (formId, record) => {
      requireLevel(cur(), pack, item.id, 'write');
      return cur().logSubmission(formId, record);
    },
    items: () => [...cur().manifest.items],
    addFile: async (name, data) => {
      requireManage(cur(), pack);
      const item = cur().addItem(name, data);
      await flushBundle(pack);
      return item;
    },
    // 选项引用别的表格：只按本包的数据读，不受条目级别限制（只读出下拉选项）
    readRange: (ref) => cur().readRange(ref),
    /** 当前登录身份对这个条目能不能写（只读时编辑器可以提示） */
    get writable() {
      const b = cache.get(pack);
      return !b || packAccess(b, pack).level(item.id) === 'write';
    },
    readText: (itemId) => {
      const it = cur().item(itemId);
      if (it && !atLeastRead(cur(), pack, itemId)) return undefined;
      return it && (it.kind === 'md' || it.kind === 'html' || it.kind === 'script') ? decode(cur().itemBytes(itemId) ?? new Uint8Array()).content : undefined;
    },
    flush: () => flushBundle(pack),
    subscribe: (cb) => cur().subscribe(cb),
  };
}

const decode = (bytes: Uint8Array): { content: string; encoding: string } => {
  try {
    const text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    return { content: text.replace(/^﻿/, ''), encoding: 'UTF-8' };
  } catch {
    return { content: new TextDecoder('gbk').decode(bytes), encoding: 'GBK' };
  }
};

/** 给宿主加上工作包能力 */
export function withBundles(host: EditorHost): EditorHost {
  base = host;
  const itemAt = async (path: string) => {
    const v = splitVPath(path);
    if (!v) return null;
    const b = await loadBundle(v.pack);
    const item = itemOf(b, v.inner);
    // 条目按 id 找：改名后旧标签页里的路径仍然有效，读写落到条目现在的位置
    return { v: item ? { pack: v.pack, inner: itemInner(item) } : v, b, item };
  };
  const writeBytes = async (path: string, data: Uint8Array): Promise<void> => {
    const at = await itemAt(path);
    if (!at) return host.writeBytes(path, data);
    if (!at.item) throw new Error(`工作包里没有 ${at.v.inner}`);
    requireLevel(at.b, at.v.pack, at.item.id, 'write');
    at.b.write(at.v.inner, data);
    await flushBundle(at.v.pack);
  };
  return {
    ...host,
    async readBytes(path) {
      const at = await itemAt(path);
      if (!at) return host.readBytes(path);
      if (at.item) requireLevel(at.b, at.v.pack, at.item.id, 'read');
      const bytes = at.b.read(at.v.inner);
      if (!bytes) throw new Error(`工作包里没有 ${at.v.inner}`);
      return bytes;
    },
    async readText(path) {
      const at = await itemAt(path);
      if (!at) return host.readText(path);
      if (at.item) requireLevel(at.b, at.v.pack, at.item.id, 'read');
      const bytes = at.b.read(at.v.inner);
      if (!bytes) throw new Error(`工作包里没有 ${at.v.inner}`);
      return decode(bytes);
    },
    writeBytes,
    async writeText(path, text, encoding) {
      if (!splitVPath(path)) return host.writeText(path, text, encoding);
      // 包内文本一律存成 UTF-8（加入时已转码）
      await writeBytes(path, new TextEncoder().encode(text));
    },
    async canWrite(path) {
      const v = splitVPath(path);
      if (!v) return host.canWrite ? host.canWrite(path) : host.workFolders().some((f) => path.startsWith(f));
      const at = await itemAt(path);
      if (at?.item && packAccess(at.b, v.pack).level(at.item.id) !== 'write') return false;
      return host.canWrite ? host.canWrite(v.pack) : true;
    },
    async openWithSystem(path) {
      if (splitVPath(path)) throw new Error('包内的文件不能直接用系统程序打开，请先导出');
      return host.openWithSystem(path);
    },
    bundleOf(path) {
      const v = splitVPath(path);
      if (!v) return null;
      const b = cache.get(v.pack);
      const item = b && itemOf(b, v.inner);
      return item ? contextFor(v.pack, item) : null;
    },
  };
}


/**
 * 工作包（.wbp）容器：一个 zip，把 xlsx / docx / md / html / 白板 / 脚本等原文件和“列属性、表单”放在一起。
 *
 *   manifest.json               格式版本、包名、条目清单
 *   items/<条目id>/<原文件名>     原样存放的文件（原字节，不转换）
 *   meta/<条目id>.columns.json   该表格条目的列属性（schema.ts）
 *   forms/<表单id>.json          表单定义（schema.ts）
 *   flows/<流程id>.json          联动流程（首页的流程卡，schema.ts）
 *   thumb.png 等                 不认识的部件按原字节保留
 *
 * 保存 = 在原 zip 上修补：没改过的部件按原字节写回，一处都没改时保存结果就是原文件字节（同 .wbd）。
 * 读入时校验路径、条目数和大小，不认识的部件保留；比本程序新的版本只读打开，避免丢内容。
 * 本文件不依赖 DOM。
 */
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { colIndex, readXlsx, type Sheet } from './xlsx';
import { cleanColumnsFile, cleanFlow, cleanForm, columnsFileJson, flowJson, formJson, type ColumnsFile, type FlowDef, type FormDef } from './schema';

export const BUNDLE_EXT = 'wbp';
/** 格式主版本：只在不兼容改动时加一（见 docs/WBP_FORMAT.md） */
export const FORMAT_VERSION = 1;
/** 文件头 `mimetype` 部件的内容，也是工作包的媒体类型（zip 的第一个条目、不压缩，方便按内容识别文件） */
export const MIMETYPE = 'application/vnd.wz.workpackage+zip';
/** 写出文件的程序标识，写进清单 generator */
export const GENERATOR = 'xioffice';
/** 虚拟路径分隔：`<包文件路径>!/<包内路径>`，如 `D:\work\订单.wbp!/items/i1/订单.xlsx` */
export const VSEP = '!/';

const MAX_ENTRY = 512 * 1024 * 1024;
const MAX_ENTRIES = 5000;
const MTIME = Date.UTC(2026, 0, 1);
/** 本身已压缩的文件，包里不再压缩（省时间，也不放大） */
const STORED_EXTS = new Set(['xlsx', 'xlsm', 'docx', 'pptx', 'zip', 'png', 'jpg', 'jpeg', 'jfif', 'webp', 'gif', 'pdf', 'wbp']);

export type ItemKind = 'sheet' | 'doc' | 'md' | 'html' | 'board' | 'script' | 'file';

export interface BundleItem {
  id: string;
  kind: ItemKind;
  /** 原文件名（含扩展名） */
  name: string;
  /** 条目树里显示的名字 */
  title: string;
  /** 媒体类型（由扩展名决定，不信任文件里写的；未知为 application/octet-stream） */
  mediaType: string;
  /** 扩展字段：清单里以 `x-` 开头的键，原样保留（第三方程序可以放自己的信息） */
  ext?: Record<string, unknown>;
}

export interface Manifest {
  format: 'wbp';
  /** 格式主版本；比本程序新时只读打开 */
  version: number;
  /** 工作包的唯一标识（UUID），复制 / 分享后仍指同一个包 */
  id: string;
  title: string;
  generator?: string;
  /** 创建 / 最后修改时间（ISO 8601，UTC） */
  created?: string;
  modified?: string;
  items: BundleItem[];
  /** 扩展字段：清单顶层以 `x-` 开头的键，原样保留 */
  ext?: Record<string, unknown>;
}

const MEDIA_TYPES: Record<string, string> = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  md: 'text/markdown',
  markdown: 'text/markdown',
  html: 'text/html',
  htm: 'text/html',
  wbd: 'application/vnd.wz.whiteboard+zip',
  js: 'text/javascript',
  ps1: 'text/x-powershell',
  py: 'text/x-python',
  json: 'application/json',
  txt: 'text/plain',
  csv: 'text/csv',
  pdf: 'application/pdf',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

export const mediaTypeOfName = (name: string): string => MEDIA_TYPES[extOfName(name)] ?? 'application/octet-stream';

/** 条目扩展字段：标记“这是流程生成的输出文件”（文件区放进“输出”虚拟文件夹；见 docs/WBP_FORMAT.md 4.1） */
export const OUTPUT_EXT = 'x-wz-output';
export const isOutputItem = (it: Pick<BundleItem, 'ext'>): boolean => it.ext?.[OUTPUT_EXT] === true;

/** 取出以 x- 开头的扩展字段 */
function extFields(o: Record<string, unknown>): Record<string, unknown> | undefined {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(o)) if (/^x-[A-Za-z0-9_.-]{1,60}$/.test(k)) out[k] = v;
  return Object.keys(out).length ? out : undefined;
}

const isoOrUndef = (v: unknown): string | undefined => (typeof v === 'string' && !Number.isNaN(Date.parse(v)) ? v : undefined);

function newUuid(): string {
  const c = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto;
  if (c?.randomUUID) return c.randomUUID();
  const h = () => Math.floor(Math.random() * 16).toString(16);
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (ch) => (ch === 'x' ? h() : ((Math.random() * 4) | 8).toString(16)));
}

const isUuid = (v: unknown): v is string => typeof v === 'string' && /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(v);

// ---------- 路径 ----------

export function splitVPath(path: string): { pack: string; inner: string } | null {
  const i = path.indexOf(VSEP);
  if (i <= 0) return null;
  const pack = path.slice(0, i);
  const inner = path.slice(i + VSEP.length);
  return inner ? { pack, inner } : null;
}

export const joinVPath = (pack: string, inner: string): string => `${pack}${VSEP}${inner}`;
export const isVPath = (path: string): boolean => splitVPath(path) !== null;

const extOfName = (name: string): string => {
  const i = name.lastIndexOf('.');
  return i > 0 ? name.slice(i + 1).toLowerCase() : '';
};

export function kindOfName(name: string): ItemKind {
  switch (extOfName(name)) {
    case 'xlsx':
    case 'xlsm':
      return 'sheet';
    case 'docx':
      return 'doc';
    case 'md':
    case 'markdown':
      return 'md';
    case 'html':
    case 'htm':
      return 'html';
    case 'wbd':
      return 'board';
    case 'js':
    case 'ps1':
    case 'py':
      return 'script';
    default:
      return 'file';
  }
}

/** 条目在工作台里用哪个编辑器打开（编辑器的 wz.kind）；打不开的返回 null */
export function editorKindOf(kind: ItemKind): string | null {
  switch (kind) {
    case 'sheet':
      return 'sheet';
    case 'doc':
      return 'word';
    case 'md':
      return 'markdown';
    case 'html':
      return 'html';
    case 'board':
      return 'whiteboard';
    case 'script':
      return 'script';
    default:
      return null;
  }
}

/** 包内路径是否安全：不能有 ..、绝对路径、反斜杠、盘符、空段和控制字符 */
export function safeInner(p: string): boolean {
  if (!p || p.length > 400 || p.startsWith('/') || p.includes('\\') || /^[a-zA-Z]:/.test(p)) return false;
  // eslint-disable-next-line no-control-regex
  if (/[\x00-\x1f]/.test(p)) return false;
  return p.split('/').every((s) => s !== '' && s !== '.' && s !== '..');
}

/** 文件名是否可以作为条目名 */
export function safeName(name: string): boolean {
  // eslint-disable-next-line no-control-regex
  return !!name && name.length <= 200 && !/[\\/:*?"<>|\x00-\x1f]/.test(name) && name !== '.' && name !== '..' && name.trim() === name;
}

const safeId = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,40}$/.test(v);

export const itemInner = (it: Pick<BundleItem, 'id' | 'name'>): string => `items/${it.id}/${it.name}`;
export const columnsInner = (itemId: string): string => `meta/${itemId}.columns.json`;
export const formInner = (formId: string): string => `forms/${formId}.json`;
export const flowInner = (flowId: string): string => `flows/${flowId}.json`;

// ---------- 清单 ----------

function cleanManifest(raw: unknown, has: (path: string) => boolean): { manifest: Manifest; newer: boolean } {
  const o = raw && typeof raw === 'object' && !Array.isArray(raw) ? (raw as Record<string, unknown>) : {};
  const version = typeof o.version === 'number' && Number.isFinite(o.version) ? o.version : FORMAT_VERSION;
  const items: BundleItem[] = [];
  const ids = new Set<string>();
  if (Array.isArray(o.items)) {
    for (const r of o.items) {
      if (!r || typeof r !== 'object') continue;
      const it = r as Record<string, unknown>;
      if (!safeId(it.id) || ids.has(it.id) || typeof it.name !== 'string' || !safeName(it.name)) continue;
      const item: BundleItem = {
        id: it.id,
        name: it.name,
        // 种类由文件名决定，不信任文件里写的
        kind: kindOfName(it.name),
        title: typeof it.title === 'string' && it.title ? it.title.slice(0, 200) : it.name,
        mediaType: mediaTypeOfName(it.name),
      };
      const itemExt = extFields(it);
      if (itemExt) item.ext = itemExt;
      if (!has(itemInner(item))) continue;
      ids.add(item.id);
      items.push(item);
    }
  }
  const manifest: Manifest = {
    format: 'wbp',
    version,
    id: isUuid(o.id) ? o.id : newUuid(),
    title: typeof o.title === 'string' ? o.title.slice(0, 200) : '',
    items,
  };
  if (typeof o.generator === 'string') manifest.generator = o.generator.slice(0, 100);
  const created = isoOrUndef(o.created);
  const modified = isoOrUndef(o.modified);
  if (created) manifest.created = created;
  if (modified) manifest.modified = modified;
  const ext = extFields(o);
  if (ext) manifest.ext = ext;
  return { manifest, newer: version > FORMAT_VERSION };
}

/** 清单写成 JSON 时的字段顺序：固定字段在前，扩展字段（x-）在后 */
function manifestJson(m: Manifest): Record<string, unknown> {
  return {
    format: m.format,
    version: m.version,
    id: m.id,
    title: m.title,
    ...(m.generator ? { generator: m.generator } : {}),
    ...(m.created ? { created: m.created } : {}),
    ...(m.modified ? { modified: m.modified } : {}),
    items: m.items.map((i) => ({ id: i.id, kind: i.kind, name: i.name, title: i.title, mediaType: i.mediaType, ...i.ext })),
    ...m.ext,
  };
}

// ---------- 容器 ----------

export class Bundle {
  manifest: Manifest;
  /** 比本程序新的版本写的：只读，不能保存 */
  newer: boolean;
  /** 当前所有部件（路径 → 字节），保持原来的顺序 */
  private parts = new Map<string, Uint8Array>();
  /** 读入 / 上次保存时的状态，保存时据此修补 */
  private origin: { bytes: Uint8Array; parts: Map<string, Uint8Array>; manifest: string } | null = null;
  private listeners = new Set<() => void>();
  /** 自上次保存后改过的部件（含已删除的） */
  private changed = new Set<string>();
  private manifestChanged = false;
  /** 每次改动加一：保存期间又有改动时，不把它当成已保存 */
  rev = 0;

  /** 当前时间（ISO，测试里可替换成固定值） */
  static now = (): string => new Date().toISOString();

  private constructor(manifest: Manifest, newer: boolean) {
    this.manifest = manifest;
    this.newer = newer;
  }

  /** 新建空包（整篇写出只用于新建文件） */
  static create(title = ''): Bundle {
    const now = Bundle.now();
    const b = new Bundle({ format: 'wbp', version: FORMAT_VERSION, id: newUuid(), title, generator: GENERATOR, created: now, modified: now, items: [] }, false);
    b.manifestChanged = true;
    return b;
  }

  static open(bytes: Uint8Array): Bundle {
    let files: Record<string, Uint8Array>;
    try {
      files = unzipSync(bytes, { filter: (f) => f.originalSize <= MAX_ENTRY });
    } catch {
      throw new Error('不是有效的工作包文件');
    }
    const names = Object.keys(files);
    if (names.length > MAX_ENTRIES) throw new Error('工作包里的文件太多');
    const parts = new Map<string, Uint8Array>();
    for (const n of names) {
      if (n.endsWith('/')) continue;
      if (!safeInner(n)) throw new Error(`工作包里有不安全的路径：${n}`);
      parts.set(n, files[n]);
    }
    const mt = parts.get('mimetype');
    if (mt && strFromU8(mt).trim() !== MIMETYPE) throw new Error('不是工作包文件');
    const mf = parts.get('manifest.json');
    if (!mf) throw new Error('工作包缺少 manifest.json');
    let raw: unknown;
    try {
      raw = JSON.parse(strFromU8(mf));
    } catch {
      throw new Error('工作包的 manifest.json 无法解析');
    }
    const { manifest, newer } = cleanManifest(raw, (p) => parts.has(p));
    const b = new Bundle(manifest, newer);
    b.parts = parts;
    b.origin = { bytes, parts: new Map(parts), manifest: strFromU8(mf) };
    return b;
  }

  // ---- 读 ----

  has(path: string): boolean {
    return this.parts.has(path);
  }

  read(path: string): Uint8Array | undefined {
    return this.parts.get(path);
  }

  readText(path: string): string | undefined {
    const b = this.parts.get(path);
    return b ? strFromU8(b) : undefined;
  }

  paths(): string[] {
    return [...this.parts.keys()];
  }

  item(id: string): BundleItem | undefined {
    return this.manifest.items.find((i) => i.id === id);
  }

  /** 包内路径 → 条目（items/<id>/<name>） */
  itemByInner(inner: string): BundleItem | undefined {
    return this.manifest.items.find((i) => itemInner(i) === inner);
  }

  itemBytes(id: string): Uint8Array | undefined {
    const it = this.item(id);
    return it ? this.parts.get(itemInner(it)) : undefined;
  }

  // ---- 写 ----

  private ensureWritable(): void {
    if (this.newer) throw new Error('这个工作包是更新版本保存的，只能查看；请用新版程序编辑');
  }

  private touch(path: string): void {
    this.changed.add(path);
    this.emit();
  }

  /** 写一个部件。条目文件、列属性、表单都走这里 */
  write(path: string, data: Uint8Array): void {
    this.ensureWritable();
    if (!safeInner(path)) throw new Error(`不安全的包内路径：${path}`);
    const old = this.parts.get(path);
    if (old && old.length === data.length && old.every((v, i) => v === data[i])) return;
    this.parts.set(path, data);
    this.touch(path);
  }

  writeText(path: string, text: string): void {
    this.write(path, strToU8(text));
  }

  remove(path: string): void {
    this.ensureWritable();
    if (this.parts.delete(path)) this.touch(path);
  }

  private nextId(): string {
    const used = new Set(this.manifest.items.map((i) => i.id));
    for (let n = this.manifest.items.length + 1; ; n++) if (!used.has(`i${n}`)) return `i${n}`;
  }

  /** 同名条目自动改名 name (1).ext，不覆盖 */
  uniqueName(name: string, except?: string): string {
    const taken = new Set(this.manifest.items.filter((i) => i.id !== except).map((i) => i.name.toLowerCase()));
    if (!taken.has(name.toLowerCase())) return name;
    const dot = name.lastIndexOf('.');
    const stem = dot > 0 ? name.slice(0, dot) : name;
    const ext = dot > 0 ? name.slice(dot) : '';
    for (let n = 1; ; n++) {
      const cand = `${stem} (${n})${ext}`;
      if (!taken.has(cand.toLowerCase())) return cand;
    }
  }

  addItem(name: string, bytes: Uint8Array, title?: string, ext?: Record<string, unknown>): BundleItem {
    this.ensureWritable();
    if (!safeName(name)) throw new Error(`文件名不合法：${name}`);
    const finalName = this.uniqueName(name);
    const item: BundleItem = { id: this.nextId(), name: finalName, kind: kindOfName(name), title: title || finalName, mediaType: mediaTypeOfName(name) };
    if (ext) item.ext = ext;
    this.manifest = { ...this.manifest, items: [...this.manifest.items, item] };
    this.manifestChanged = true;
    this.parts.set(itemInner(item), bytes);
    this.touch(itemInner(item));
    return item;
  }

  /** 删除条目，同时删掉它的列属性和指向它的表单 */
  removeItem(id: string): void {
    this.ensureWritable();
    const it = this.item(id);
    if (!it) return;
    for (const p of [itemInner(it), columnsInner(id)]) {
      if (this.parts.delete(p)) this.changed.add(p);
    }
    for (const f of this.forms()) if (f.target.item === id) this.deleteForm(f.id);
    this.manifest = { ...this.manifest, items: this.manifest.items.filter((i) => i.id !== id) };
    this.manifestChanged = true;
    this.emit();
  }

  /** 改条目的文件名（扩展名不变；包内路径跟着变，id 不变所以列属性和表单不受影响） */
  renameItem(id: string, name: string): void {
    this.ensureWritable();
    const it = this.item(id);
    if (!it || it.name === name) return;
    if (!safeName(name)) throw new Error(`文件名不合法：${name}`);
    if (extOfName(name) !== extOfName(it.name)) throw new Error('改名不能改扩展名');
    const finalName = this.uniqueName(name, id);
    const from = itemInner(it);
    const bytes = this.parts.get(from);
    if (!bytes) return;
    const next: BundleItem = { ...it, name: finalName, title: it.title === it.name ? finalName : it.title };
    this.parts.delete(from);
    this.changed.add(from);
    this.parts.set(itemInner(next), bytes);
    this.changed.add(itemInner(next));
    this.manifest = { ...this.manifest, items: this.manifest.items.map((i) => (i.id === id ? next : i)) };
    this.manifestChanged = true;
    this.emit();
  }

  /** 调整条目顺序：把 id 条目移到 targetId 条目的前面（after 为真时移到后面）；顺序写在 manifest 里 */
  moveItem(id: string, targetId: string, after = false): void {
    this.ensureWritable();
    if (id === targetId) return;
    const items = this.manifest.items.filter((i) => i.id !== id);
    const moving = this.item(id);
    const at = items.findIndex((i) => i.id === targetId);
    if (!moving || at < 0) return;
    items.splice(after ? at + 1 : at, 0, moving);
    if (items.every((it, i) => it === this.manifest.items[i])) return;
    this.manifest = { ...this.manifest, items };
    this.manifestChanged = true;
    this.emit();
  }

  setTitle(title: string): void {
    this.ensureWritable();
    if (this.manifest.title === title) return;
    this.manifest = { ...this.manifest, title };
    this.manifestChanged = true;
    this.emit();
  }

  // ---- 跨条目引用（单选 / 多选的选项取自包里别的表格）----

  private sheetCache = new WeakMap<Uint8Array, Sheet[]>();

  /**
   * 读 `[条目名或id]工作表!A2:A20` 里的文字（去空去重，按出现顺序，最多 2000 格）。条目 / 工作表 / 区域不对返回空数组。
   * 取的是 xlsx 里存的值（公式取缓存结果），不依赖打开了表格编辑器。
   */
  readRange(ref: string): string[] {
    const m = /^\[([^\]]+)\]\s*(?:(?:'((?:[^']|'')+)'|([^!']+))!)?\s*\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(ref.trim());
    if (!m) return [];
    const key = m[1]!;
    const item = this.manifest.items.find((i) => i.id === key || i.name === key || i.title === key);
    const bytes = item && item.kind === 'sheet' ? this.itemBytes(item.id) : undefined;
    if (!item || !bytes) return [];
    let sheets = this.sheetCache.get(bytes);
    if (!sheets) {
      try {
        sheets = readXlsx(bytes);
      } catch {
        return [];
      }
      this.sheetCache.set(bytes, sheets);
    }
    const name = (m[2] ?? m[3] ?? '').replace(/''/g, "'").trim();
    const sheet = name ? sheets.find((x) => x.name === name) : sheets[0];
    if (!sheet) return [];
    const c0 = colIndex(m[4]!);
    const r0 = Number(m[5]) - 1;
    const c1 = m[6] ? colIndex(m[6]) : c0;
    const r1 = m[7] ? Number(m[7]) - 1 : r0;
    const seen = new Set<string>();
    const out: string[] = [];
    for (let r = Math.min(r0, r1); r <= Math.max(r0, r1) && r <= Math.min(r0, r1) + 2000; r++) {
      for (let c = Math.min(c0, c1); c <= Math.max(c0, c1); c++) {
        const v = sheet.rows[r]?.[c];
        const t = v === null || v === undefined ? '' : String(v).trim();
        if (t && !seen.has(t)) {
          seen.add(t);
          out.push(t);
        }
      }
    }
    return out;
  }

  // ---- 列属性 / 表单 ----

  columns(itemId: string): ColumnsFile {
    const text = this.readText(columnsInner(itemId));
    if (text === undefined) return { v: 1, sheets: {} };
    try {
      return cleanColumnsFile(JSON.parse(text));
    } catch {
      return { v: 1, sheets: {} };
    }
  }

  setColumns(itemId: string, file: ColumnsFile): void {
    const path = columnsInner(itemId);
    if (!Object.keys(file.sheets).length) return this.remove(path);
    this.writeText(path, JSON.stringify(columnsFileJson(file), null, 2));
  }

  forms(): FormDef[] {
    const out: FormDef[] = [];
    for (const p of this.parts.keys()) {
      if (!/^forms\/[^/]+\.json$/.test(p)) continue;
      try {
        const f = cleanForm(JSON.parse(this.readText(p)!));
        if (f && this.item(f.target.item)) out.push(f);
      } catch {
        /* 坏的表单文件跳过，但部件仍保留 */
      }
    }
    return out;
  }

  form(id: string): FormDef | undefined {
    return this.forms().find((f) => f.id === id);
  }

  saveForm(form: FormDef): void {
    // 先展开成文件里的 JSON 形式再清洗，保证 ext 里的 x- 字段不丢
    const clean = cleanForm(formJson(form));
    if (!clean) throw new Error('表单定义不完整');
    this.writeText(formInner(clean.id), JSON.stringify(formJson(clean), null, 2));
  }

  /** 追加一行到 JSON Lines 部件（提交记录）；上一行没有换行结尾时先补上 */
  appendLine(path: string, line: string): void {
    const old = this.readText(path) ?? '';
    const nl = String.fromCharCode(10);
    this.writeText(path, `${old}${old && !old.endsWith(nl) ? nl : ''}${line}${nl}`);
  }

  logSubmission(formId: string, record: Record<string, unknown>): void {
    this.appendLine(`forms/${formId}.log.jsonl`, JSON.stringify(record));
  }

  deleteForm(id: string): void {
    this.remove(formInner(id));
  }

  newFormId(): string {
    const used = new Set(this.forms().map((f) => f.id));
    for (let n = used.size + 1; ; n++) if (!used.has(`f${n}`)) return `f${n}`;
  }

  // ---- 联动流程 ----

  flows(): FlowDef[] {
    const out: FlowDef[] = [];
    for (const p of [...this.parts.keys()].sort()) {
      if (!/^flows\/[^/]+\.json$/.test(p)) continue;
      try {
        const f = cleanFlow(JSON.parse(this.readText(p)!));
        if (f) out.push(f);
      } catch {
        /* 坏的流程文件跳过，但部件仍保留 */
      }
    }
    return out;
  }

  flow(id: string): FlowDef | undefined {
    return this.flows().find((f) => f.id === id);
  }

  saveFlow(flow: FlowDef): void {
    const clean = cleanFlow(flowJson(flow));
    if (!clean) throw new Error('流程定义不完整');
    this.writeText(flowInner(clean.id), JSON.stringify(flowJson(clean), null, 2));
  }

  deleteFlow(id: string): void {
    this.remove(flowInner(id));
  }

  newFlowId(): string {
    const used = new Set(this.flows().map((f) => f.id));
    for (let n = used.size + 1; ; n++) if (!used.has(`w${n}`)) return `w${n}`;
  }

  // ---- 保存 ----

  get dirty(): boolean {
    return this.manifestChanged || this.changed.size > 0;
  }

  /** 保存成的字节：没改过 = 原文件字节；否则在原 zip 上修补（没改的部件按原字节写回） */
  serialize(): Uint8Array {
    this.ensureWritable();
    if (this.origin && !this.dirty) return this.origin.bytes;
    // 有改动就更新修改时间（只在真的要写出时更新，没改动时字节不变）
    if (this.dirty) this.manifest = { ...this.manifest, modified: Bundle.now(), generator: GENERATOR };
    const manifestText = JSON.stringify(manifestJson(this.manifest), null, 2);
    const out: Zippable = {};
    const entry = (name: string, data: Uint8Array) => {
      out[name] = [data, { level: STORED_EXTS.has(extOfName(name)) ? 0 : 6, mtime: MTIME }];
    };
    // mimetype 必须是第一个条目、不压缩（按内容识别文件类型）
    out.mimetype = [strToU8(MIMETYPE), { level: 0, mtime: MTIME }];
    entry('manifest.json', strToU8(manifestText));
    for (const [name, data] of this.parts) if (name !== 'manifest.json' && name !== 'mimetype') entry(name, data);
    return zipSync(out);
  }

  /** 写盘成功后调用：以刚保存的字节作为新的原文件 */
  markSaved(bytes: Uint8Array, rev = this.rev): void {
    // 保存期间又有改动：保持“已修改”，下一次保存再写
    if (rev !== this.rev) return;
    this.parts.set('manifest.json', strToU8(JSON.stringify(manifestJson(this.manifest), null, 2)));
    this.parts.set('mimetype', strToU8(MIMETYPE));
    this.origin = { bytes, parts: new Map(this.parts), manifest: strFromU8(this.parts.get('manifest.json')!) };
    this.changed.clear();
    this.manifestChanged = false;
  }

  /** 磁盘上的文件被外部改了：就地换成新内容（订阅者收到通知）。有未保存的修改时不要调用 */
  reloadFrom(bytes: Uint8Array): void {
    const nb = Bundle.open(bytes);
    this.manifest = nb.manifest;
    this.newer = nb.newer;
    this.parts = nb.parts;
    this.origin = nb.origin;
    this.changed.clear();
    this.manifestChanged = false;
    this.emit();
  }

  /** 保存用的字节和当时的改动序号（传给 markSaved） */
  snapshot(): { bytes: Uint8Array; rev: number } {
    return { bytes: this.serialize(), rev: this.rev };
  }

  // ---- 订阅 ----

  subscribe(cb: () => void): () => void {
    this.listeners.add(cb);
    return () => void this.listeners.delete(cb);
  }

  private emit(): void {
    this.rev++;
    for (const cb of [...this.listeners]) cb();
  }
}

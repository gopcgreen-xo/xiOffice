/**
 * 表格的编辑操作。每个操作都在一个事务（`Tx`）里修改模型，事务记下改动前后的内容，用于撤销 / 重做：
 * - 单元格级的改动只记被改的格子（大表格上编辑一格不用复制整张表）；
 * - 插入 / 删除行列这类结构改动记整张表的前后快照；
 * - 行高列宽、合并、冻结等表属性记 Meta。
 * 样式只会追加（deriveStyle），撤销时不用回退 styleBook。
 */
import type { SheetColumns } from '@wz/editor-shared';
import { shiftColumnMeta } from './colmeta';
import { formatCode, deriveStyle, type StyleChange } from './styles';
import { borderFor, type BorderKind } from './format';
import { parseInput, type Parsed } from './input';
import { normalizeFormula } from './formula/parse';
import { formulaFormat } from './formula/eval';
import { MAX_COL, MAX_ROW, mapRefs, moveFormula, shiftFormula, shiftSpan, shiftSqref, withinBand, type RefPart, type ShiftSpec } from './formula/refs';
import { deriveAxes, refreshText, syncStyles, styleAt, updateExtent, type AutoFilter, type ColDef, type PageSetup, type Range, type RowDef, type SCell, type SheetData, type Workbook } from './workbook';

interface Meta {
  merges: Range[];
  colDefs: ColDef[];
  rowDefs: Record<number, RowDef>;
  freeze: { rows: number; cols: number };
  autoFilter?: AutoFilter;
  page?: PageSetup;
  cf?: SheetData['cf'];
  cfRemoved?: number[];
  protected?: boolean;
  protection?: SheetData['protection'];
}

/** 整本快照：工作表列表与定义名称（增删、移动、改名工作表时用） */
interface BookSnap {
  sheets: SheetData[];
  names?: Workbook['names'];
  wbProtection?: Workbook['wbProtection'];
}

interface CellRec {
  si: number;
  r: number;
  c: number;
  before?: SCell;
  blankBefore?: number;
  after?: SCell;
  blankAfter?: number;
}

export interface Entry {
  label: string;
  cells: CellRec[];
  sheets: { si: number; before: SheetData; after?: SheetData }[];
  metas: { si: number; before: Meta; after?: Meta }[];
  /** 整个工作表列表（增删工作表） */
  book?: { before: BookSnap; after?: BookSnap };
  /** 列属性的改动（工作包里的字段类型；只改它不算改了 xlsx，不会让工作表被重写） */
  cols?: { si: number; before?: SheetColumns; after?: SheetColumns }[];
  /** 界面状态（选区），撤销时恢复 */
  ui?: unknown;
  /** 保护 / 取消保护本身的改动：不受工作表保护限制 */
  bypass?: boolean;
}

const clone = <T>(x: T): T => structuredClone(x);
const metaOf = (s: SheetData): Meta => clone({ merges: s.merges, colDefs: s.colDefs ?? [], rowDefs: s.rowDefs ?? {}, freeze: s.freeze, autoFilter: s.autoFilter, page: s.page, cf: s.cf, cfRemoved: s.cfRemoved, protected: s.protected, protection: s.protection });
const bookOf = (b: Workbook): BookSnap => clone({ sheets: b.sheets, names: b.names, wbProtection: b.wbProtection });
const k3 = (si: number, r: number, c: number) => `${si}:${r}:${c}`;

export class Tx {
  readonly entry: Entry;
  private cellIdx = new Map<string, CellRec>();
  private snapped = new Set<number>();
  private metaed = new Set<number>();
  private colsed = new Set<number>();
  /** 改过值 / 公式的格子（重算用） */
  readonly changed: { si: number; r: number; c: number }[] = [];
  /** 有结构改动（插删行列、排序等）：依赖关系要整体重建 */
  structural = false;
  /** 保护 / 取消保护：这次改动不受工作表保护限制 */
  set bypass(v: boolean) {
    this.entry.bypass = v || undefined;
  }

  constructor(
    readonly book: Workbook,
    label: string,
    ui?: unknown,
  ) {
    this.entry = { label, cells: [], sheets: [], metas: [], ui };
  }

  /** 修改格子前调用 */
  touch(si: number, r: number, c: number): void {
    const sheet = this.book.sheets[si];
    sheet.dirty = true;
    this.changed.push({ si, r, c });
    if (this.snapped.has(si) || this.entry.book) return;
    const key = k3(si, r, c);
    if (this.cellIdx.has(key)) return;
    const rec: CellRec = { si, r, c, before: clone(sheet.cells[r]?.[c]), blankBefore: sheet.blankStyles?.[r]?.[c] };
    this.cellIdx.set(key, rec);
    this.entry.cells.push(rec);
  }

  /** 结构改动前调用：记整张表 */
  snap(si: number): void {
    const sheet = this.book.sheets[si];
    sheet.dirty = true;
    this.structural = true;
    if (this.snapped.has(si) || this.entry.book) return;
    this.snapped.add(si);
    this.entry.sheets.push({ si, before: clone(sheet) });
  }

  /** 改表属性（合并、行列定义、冻结）前调用 */
  meta(si: number): void {
    const sheet = this.book.sheets[si];
    sheet.dirty = true;
    if (this.snapped.has(si) || this.metaed.has(si) || this.entry.book) return;
    this.metaed.add(si);
    this.entry.metas.push({ si, before: metaOf(sheet) });
  }

  /** 改列属性（sheet.columns）前调用。不标记工作表已修改：只改列属性不会重写 xlsx */
  columns(si: number): void {
    if (this.colsed.has(si) || this.snapped.has(si) || this.entry.book) return;
    this.colsed.add(si);
    (this.entry.cols ??= []).push({ si, before: clone(this.book.sheets[si].columns) });
  }

  /** 增删 / 移动工作表前调用：记整个工作表列表 */
  wholeBook(): void {
    this.structural = true;
    if (!this.entry.book) this.entry.book = { before: bookOf(this.book) };
  }

  /** 结束事务：记下改动后的内容；什么都没改返回 null */
  commit(): Entry | null {
    const e = this.entry;
    const b = this.book;
    if (e.book) {
      e.book.after = bookOf(b);
      e.cells = [];
      e.sheets = [];
      e.metas = [];
      e.cols = undefined;
    } else {
      for (const rec of e.cells) {
        const s = b.sheets[rec.si];
        rec.after = clone(s.cells[rec.r]?.[rec.c]);
        rec.blankAfter = s.blankStyles?.[rec.r]?.[rec.c];
      }
      e.cells = e.cells.filter((x) => JSON.stringify([x.before, x.blankBefore]) !== JSON.stringify([x.after, x.blankAfter]));
      for (const x of e.sheets) x.after = clone(b.sheets[x.si]);
      for (const x of e.metas) x.after = metaOf(b.sheets[x.si]);
      for (const x of e.cols ?? []) x.after = clone(b.sheets[x.si].columns);
      e.cols = e.cols?.filter((x) => JSON.stringify(x.before) !== JSON.stringify(x.after));
      if (!e.cols?.length) e.cols = undefined;
    }
    for (const s of b.sheets) if (s.dirty) finish(s);
    syncStyles(b);
    return e.book || e.cells.length || e.sheets.length || e.metas.length || e.cols?.length ? e : null;
  }
}

/**
 * 一次改动是否违反工作表保护（同 Excel 默认的保护选项）：受保护的表只能改未锁定单元格的内容，
 * 行列、合并、筛选、排序等结构和属性都不能改。单元格是否锁定看样式的 protection locked（缺省锁定）。
 */
export function violatesProtection(book: Workbook, e: Entry): boolean {
  if (e.bypass) return false;
  const prot = (si: number) => !!book.sheets[si]?.protected;
  if (e.sheets.some((x) => prot(x.si)) || e.metas.some((x) => prot(x.si))) return true;
  return e.cells.some((x) => {
    if (!prot(x.si)) return false;
    const s = x.before ? x.before.s : (x.blankBefore ?? styleAt(book.sheets[x.si], x.r, x.c));
    const xf = book.styleBook?.xfs[s ?? 0];
    return !/<protection\b[^>]*\blocked="(0|false)"/.test(xf?.inner ?? '');
  });
}

function finish(s: SheetData) {
  deriveAxes(s);
  updateExtent(s);
}

function putCell(s: SheetData, r: number, c: number, cell: SCell | undefined, blank: number | undefined) {
  if (cell) (s.cells[r] ??= [])[c] = clone(cell);
  else if (s.cells[r]) s.cells[r][c] = undefined;
  if (blank !== undefined) ((s.blankStyles ??= [])[r] ??= [])[c] = blank;
  else if (s.blankStyles?.[r]) s.blankStyles[r][c] = undefined;
}

/** 撤销（dir = -1）或重做（dir = 1）一条记录 */
export function applyEntry(book: Workbook, e: Entry, dir: 1 | -1): void {
  const pick = <T>(x: { before: T; after?: T }) => clone(dir < 0 ? x.before : (x.after as T));
  if (e.book) {
    const snap = pick(e.book);
    book.sheets = snap.sheets;
    book.names = snap.names;
    book.wbProtection = snap.wbProtection;
    book.namesDirty = true;
    for (const s of book.sheets) s.dirty = true;
    return;
  }
  for (const x of e.sheets) {
    book.sheets[x.si] = pick(x);
    book.sheets[x.si].dirty = true;
  }
  for (const x of e.metas) {
    const s = book.sheets[x.si];
    Object.assign(s, pick(x));
    s.dirty = true;
  }
  for (const x of e.cols ?? []) {
    const s = book.sheets[x.si];
    const cols = clone(dir < 0 ? x.before : x.after);
    if (cols) s.columns = cols;
    else delete s.columns;
  }
  const list = dir < 0 ? [...e.cells].reverse() : e.cells;
  for (const x of list) {
    const s = book.sheets[x.si];
    putCell(s, x.r, x.c, dir < 0 ? x.before : x.after, dir < 0 ? x.blankBefore : x.blankAfter);
    s.dirty = true;
  }
  for (const s of book.sheets) if (s.dirty) finish(s);
  syncStyles(book);
}

// ---------- 单元格 ----------

const fmtAt = (book: Workbook, s: number | undefined) => (book.styleBook ? formatCode(book.styleBook, s) : 'General');

/** 写入一个格子的内容（保留样式；新格子沿用行 / 列格式） */
export function setParsed(tx: Tx, si: number, r: number, c: number, p: Parsed): void {
  const book = tx.book;
  const sheet = book.sheets[si];
  tx.touch(si, r, c);
  const old = sheet.cells[r]?.[c];
  let s = old ? old.s : styleAt(sheet, r, c);
  if (p.kind === 'empty') {
    putCell(sheet, r, c, undefined, s !== undefined && s > 0 ? s : undefined);
    return;
  }
  if (p.kind === 'value' && p.fmt && book.styleBook) s = deriveStyle(book.styleBook, s, { numFmt: p.fmt });
  const f = p.kind === 'formula' ? normalizeFormula(p.f) : '';
  if (p.kind === 'formula' && book.styleBook && fmtAt(book, s) === 'General') {
    const fmt = formulaFormat(f);
    if (fmt) s = deriveStyle(book.styleBook, s, { numFmt: fmt });
  }
  const cell: SCell = p.kind === 'formula' ? { v: '', t: '', f, dyn: true } : { v: p.v, t: '' };
  if (s !== undefined && s > 0) cell.s = s;
  if (old?.xa) cell.xa = old.xa;
  refreshText(book, cell);
  putCell(sheet, r, c, cell, undefined);
}

/** 输入框文字写入格子 */
export function setInput(tx: Tx, si: number, r: number, c: number, raw: string): void {
  const sheet = tx.book.sheets[si];
  const s = sheet.cells[r]?.[c]?.s ?? styleAt(sheet, r, c);
  setParsed(tx, si, r, c, parseInput(raw, fmtAt(tx.book, s), tx.book.date1904));
}

/** 区域里的格子（跳过合并区域里非左上角的格子） */
export function eachCell(sheet: SheetData, g: Range, fn: (r: number, c: number) => void): void {
  const covered = new Set<number>();
  for (const m of sheet.merges) {
    if (m.r1 < g.r0 || m.r0 > g.r1 || m.c1 < g.c0 || m.c0 > g.c1) continue;
    for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(r * 16384 + c);
  }
  for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) if (!covered.has(r * 16384 + c)) fn(r, c);
}

/** 清除内容（what = contents 保留格式；formats 只清格式；all 全部） */
export function clearRange(tx: Tx, si: number, g: Range, what: 'contents' | 'formats' | 'all' = 'contents'): void {
  const sheet = tx.book.sheets[si];
  const r1 = Math.min(g.r1, Math.max(sheet.rows, sheet.blankStyles?.length ?? 0));
  for (let r = g.r0; r <= r1; r++) {
    const row = sheet.cells[r];
    const brow = sheet.blankStyles?.[r];
    const c1 = Math.min(g.c1, Math.max(row?.length ?? 0, brow?.length ?? 0) - 1);
    for (let c = g.c0; c <= c1; c++) {
      const cell = row?.[c];
      const blank = brow?.[c];
      if (!cell && blank === undefined) continue;
      tx.touch(si, r, c);
      if (what === 'contents') putCell(sheet, r, c, undefined, cell ? (cell.s ? cell.s : undefined) : blank);
      else if (what === 'all') putCell(sheet, r, c, undefined, undefined);
      else if (cell) {
        const next = { ...cell };
        delete next.s;
        refreshText(tx.book, next);
        putCell(sheet, r, c, next, undefined);
      } else putCell(sheet, r, c, undefined, undefined);
    }
  }
  if (what !== 'contents') {
    // 合并区域也属于格式
    if (what === 'all' || what === 'formats') {
      const keep = sheet.merges.filter((m) => !(m.r0 >= g.r0 && m.r1 <= g.r1 && m.c0 >= g.c0 && m.c1 <= g.c1));
      if (keep.length !== sheet.merges.length) {
        tx.meta(si);
        sheet.merges = keep;
      }
    }
  }
}

/** 改格式：区域内每个格子（含空格子）在原格式上改 */
export function applyStyle(tx: Tx, si: number, g: Range, ch: StyleChange): void {
  const book = tx.book;
  if (!book.styleBook) return;
  const sheet = book.sheets[si];
  const cache = new Map<number | undefined, number>();
  const derive = (s: number | undefined) => {
    let n = cache.get(s);
    if (n === undefined) cache.set(s, (n = deriveStyle(book.styleBook!, s, ch)));
    return n;
  };
  for (let r = g.r0; r <= g.r1; r++) {
    for (let c = g.c0; c <= g.c1; c++) {
      tx.touch(si, r, c);
      const cell = sheet.cells[r]?.[c];
      if (cell) {
        const next: SCell = { ...cell, s: derive(cell.s) };
        if (!next.s) delete next.s;
        refreshText(book, next);
        putCell(sheet, r, c, next, undefined);
      } else {
        const s = derive(styleAt(sheet, r, c));
        putCell(sheet, r, c, undefined, s || undefined);
      }
    }
  }
}

// ---------- 复制粘贴 / 填充 ----------

export interface ClipCell {
  v: string | number | boolean;
  t: string;
  s?: number;
  f?: string;
  /** 动态数组公式 */
  dyn?: boolean;
  k?: 'e' | 'd';
  /** 来源位置（公式相对引用按它平移） */
  r: number;
  c: number;
}

export interface Block {
  rows: number;
  cols: number;
  /** 按 [行][列]，空格子为 undefined */
  cells: (ClipCell | undefined)[][];
  /** 来源区域里的合并（相对左上角） */
  merges?: Range[];
  /** 来自本表格（带格式、公式）；外部文字只有值 */
  internal: boolean;
  /** 来源的行高 / 整行格式（相对第一行），插入复制的整行时带上 */
  rowDefs?: Record<number, RowDef>;
  /** 来源的列定义（相对第一列），插入复制的整列时带上 */
  colDefs?: ColDef[];
}

/** 复制区域到剪贴板模型 */
export function copyBlock(book: Workbook, si: number, g: Range): Block {
  const sheet = book.sheets[si];
  const cells: (ClipCell | undefined)[][] = [];
  for (let r = g.r0; r <= g.r1; r++) {
    const row: (ClipCell | undefined)[] = [];
    for (let c = g.c0; c <= g.c1; c++) {
      const x = sheet.cells[r]?.[c];
      const s = x ? x.s : styleAt(sheet, r, c);
      if (x) row.push({ v: x.v, t: x.t, s: x.s, f: x.f, ...(x.dyn && { dyn: true }), k: x.k, r, c });
      else row.push(s !== undefined && s > 0 ? { v: '', t: '', s, r, c } : undefined);
    }
    cells.push(row);
  }
  const merges = sheet.merges
    .filter((m) => m.r0 >= g.r0 && m.r1 <= g.r1 && m.c0 >= g.c0 && m.c1 <= g.c1)
    .map((m) => ({ r0: m.r0 - g.r0, c0: m.c0 - g.c0, r1: m.r1 - g.r0, c1: m.c1 - g.c0 }));
  const rowDefs: Record<number, RowDef> = {};
  for (const [k, d] of Object.entries(sheet.rowDefs ?? {})) if (+k >= g.r0 && +k <= g.r1) rowDefs[+k - g.r0] = { ...d };
  const colDefs = (sheet.colDefs ?? [])
    .filter((d) => d.max >= g.c0 && d.min <= g.c1)
    .map((d) => ({ ...d, min: Math.max(d.min, g.c0) - g.c0, max: Math.min(d.max, g.c1) - g.c0 }));
  return { rows: g.r1 - g.r0 + 1, cols: g.c1 - g.c0 + 1, cells, merges, internal: true, rowDefs, colDefs };
}

/**
 * 剪切移动后改写公式：指向 sheet 表、完全落在 src 里的引用（单格、区域，整行 / 整列引用要求 src 也是整行 / 整列）
 * 改指平移 (dr, dc) 后的位置，绝对引用也跟着走（引用的是被移动的格子本身，同 Excel）。owner 为公式所在工作表名。
 */
export function retargetFormula(formula: string, owner: string, sheet: string, src: Range, dr: number, dc: number): string {
  const inside = (p: RefPart) =>
    (p.r === undefined ? src.r0 === 0 && src.r1 >= MAX_ROW : p.r >= src.r0 && p.r <= src.r1) &&
    (p.c === undefined ? src.c0 === 0 && src.c1 >= MAX_COL : p.c >= src.c0 && p.c <= src.c1);
  const mv = (p: RefPart): RefPart => ({ ...p, ...(p.r !== undefined && { r: p.r + dr }), ...(p.c !== undefined && { c: p.c + dc }) });
  return mapRefs(formula, (ref) => {
    if ((ref.sheet ?? owner).toLowerCase() !== sheet.toLowerCase()) return undefined;
    if (!inside(ref.a) || (ref.b && !inside(ref.b))) return undefined;
    return { sheet: ref.sheet, a: mv(ref.a), ...(ref.b && { b: mv(ref.b) }) };
  });
}

/**
 * 插入剪切的单元格（同 Excel）：把 src 移到 (r0, c0) 前面，原位置的格子删掉并让后面的格子补上来。
 * shift 同 insertCopied（'rows' / 'cols' 时 src 是整行 / 整列）。公式里指向被移动格子的引用跟着走，其它引用随插入 / 删除平移。
 * 做法：在目标处插入 → 把格子、合并、批注、行高 / 列宽搬过去并改写引用 → 删掉原位置。
 * 返回移动后的区域；目标就在原位置（不用移动）时返回 src；不能移动时返回原因。
 */
export function insertCut(tx: Tx, si: number, src: Range, r0: number, c0: number, shift: 'rows' | 'cols' | 'down' | 'right'): Range | string {
  const sheet = tx.book.sheets[si];
  const byRow = shift === 'rows' || shift === 'down';
  const axis = byRow ? 'row' : 'col';
  const h = src.r1 - src.r0 + 1;
  const w = src.c1 - src.c0 + 1;
  const n = byRow ? h : w;
  const lo = byRow ? src.r0 : src.c0;
  const at = byRow ? r0 : c0;
  // 部分单元格：另一方向上的位置（下移时是列，右移时是行）
  const so = byRow ? src.c0 : src.r0;
  const to = shift === 'rows' || shift === 'cols' ? so : byRow ? c0 : r0;
  const span = byRow ? w : h;
  const band: [number, number] | undefined = shift === 'down' ? [c0, c0 + w - 1] : shift === 'right' ? [r0, r0 + h - 1] : undefined;
  const srcBand: [number, number] | undefined = band && (byRow ? [src.c0, src.c1] : [src.r0, src.r1]);
  // 同一列（行）带里、目标就在原位置或紧挨着它下面：不用动
  if (to === so && at >= lo && at <= lo + n) return src;
  // 目标带和原位置的带部分重叠：插入时原格子只移动一部分，无法完成（Excel 也不允许）
  if (band && to !== so && to < so + span && so < to + span) return '无法把剪切的单元格插入到与原位置部分重叠的地方';
  if (band) {
    const blocked = cellShiftBlocked(sheet, axis, at, n, band) ?? cellShiftBlocked(sheet, axis, lo, -n, srcBand!);
    if (blocked) return blocked;
  }
  // 1. 目标处插入；原位置在插入带里且在后面时被挤开
  insertDelete(tx, si, axis, at, n, band);
  const pushed = lo >= at && (!band || to === so);
  const from = pushed ? lo + n : lo;
  const moved: Range = byRow ? { r0: from, r1: from + h - 1, c0: src.c0, c1: src.c1 } : { r0: src.r0, r1: src.r1, c0: from, c1: from + w - 1 };
  const dr = byRow ? at - from : shift === 'cols' ? 0 : r0 - src.r0;
  const dc = byRow ? (shift === 'rows' ? 0 : c0 - src.c0) : at - from;
  // 整行 / 整列插入时新行列沿用了相邻行列的格式，搬过来的格子全用来源的
  if (shift === 'rows') for (let i = 0; i < h; i++) delete sheet.blankStyles?.[at + i];
  if (shift === 'cols') sheet.blankStyles?.forEach((row) => {
    for (let j = 0; j < w; j++) delete row?.[at + j];
  });
  // 2. 搬格子（原样，包括公式文字）
  const take = <T,>(grid: (T | undefined)[][] | undefined, r: number, c: number) => {
    const v = grid?.[r]?.[c];
    if (grid?.[r]) delete grid[r][c];
    return v;
  };
  const cells: { r: number; c: number; x?: SCell; b?: number }[] = [];
  const cTo = shift === 'rows' ? Math.max(src.c1, ...sheet.cells.slice(moved.r0, moved.r1 + 1).map((row) => (row?.length ?? 0) - 1), ...(sheet.blankStyles ?? []).slice(moved.r0, moved.r1 + 1).map((row) => (row?.length ?? 0) - 1)) : moved.c1;
  const rTo = shift === 'cols' ? Math.max(moved.r1, sheet.cells.length - 1, (sheet.blankStyles?.length ?? 0) - 1) : moved.r1;
  for (let r = moved.r0; r <= rTo; r++) for (let c = moved.c0; c <= cTo; c++) cells.push({ r: r + dr, c: c + dc, x: take(sheet.cells, r, c), b: take(sheet.blankStyles, r, c) });
  for (const { r, c, x, b } of cells) {
    if (x) (sheet.cells[r] ??= [])[c] = x;
    else if (sheet.cells[r]) delete sheet.cells[r][c];
    if (b) ((sheet.blankStyles ??= [])[r] ??= [])[c] = b;
    else if (sheet.blankStyles?.[r]) delete sheet.blankStyles[r][c];
  }
  const inMoved = (g: Range) => g.r0 >= moved.r0 && g.r1 <= moved.r1 && g.c0 >= moved.c0 && g.c1 <= moved.c1;
  const mvRange = (g: Range): Range => ({ r0: g.r0 + dr, r1: g.r1 + dr, c0: g.c0 + dc, c1: g.c1 + dc });
  sheet.merges = sheet.merges.map((m) => (inMoved(m) ? mvRange(m) : m));
  if (sheet.notes) sheet.notes = sheet.notes.map((x) => (inMoved({ r0: x.r, r1: x.r, c0: x.c, c1: x.c }) ? { ...x, r: x.r + dr, c: x.c + dc } : x));
  if (shift === 'rows') {
    const defs = (sheet.rowDefs ??= {});
    for (let i = 0; i < h; i++) {
      const d: RowDef = { ...defs[from + i] };
      delete d.level;
      delete d.collapsed;
      const level = defs[at + i]?.level;
      if (level) d.level = level;
      if (Object.keys(d).length) defs[at + i] = d;
      else delete defs[at + i];
    }
  } else if (shift === 'cols') {
    const old = splitDefs(sheet.colDefs ?? [], Math.min(from, at), Math.max(from, at) + w - 1);
    const defs = old.filter((d) => d.max < at || d.min > at + w - 1);
    for (const d of old) {
      if (d.min < from || d.max > from + w - 1) continue;
      const nd: ColDef = { ...d, min: d.min + dc, max: d.max + dc };
      delete nd.level;
      delete nd.collapsed;
      defs.push(nd);
    }
    sheet.colDefs = mergeDefs(defs);
  }
  // 公式：指向被移动格子的引用跟着走
  const whole: Range = shift === 'rows' ? { ...moved, c0: 0, c1: MAX_COL } : shift === 'cols' ? { ...moved, r0: 0, r1: MAX_ROW } : moved;
  tx.book.sheets.forEach((s, i) => {
    s.cells.forEach((row, r) => {
      row?.forEach((cell, c) => {
        if (!cell?.f) return;
        const f = retargetFormula(cell.f, s.name, sheet.name, whole, dr, dc);
        if (f === cell.f) return;
        if (i !== si) tx.touch(i, r, c);
        s.cells[r][c] = { ...cell, f };
      });
    });
  });
  // 3. 删掉原位置；目标在它后面时跟着前移
  insertDelete(tx, si, axis, from, -n, srcBand);
  const dest: Range = byRow ? { r0: at, r1: at + h - 1, c0: src.c0 + dc, c1: src.c1 + dc } : { r0: src.r0 + dr, r1: src.r1 + dr, c0: at, c1: at + w - 1 };
  const back = at > from && (!band || to === so) ? n : 0;
  return byRow ? { ...dest, r0: dest.r0 - back, r1: dest.r1 - back } : { ...dest, c0: dest.c0 - back, c1: dest.c1 - back };
}

/**
 * 插入复制的单元格（同 Excel）：先在 (r0, c0) 处腾出和复制块一样大的位置，再把块粘贴进去。
 * shift = 'rows' / 'cols'：复制的是整行 / 整列，插入同样多的整行 / 整列，行高 / 列宽跟着来源；
 * 'down' / 'right'：只把块覆盖的那几列 / 几行里的格子下移 / 右移（调用方先用 cellShiftBlocked 检查合并单元格）。
 */
export function insertCopied(tx: Tx, si: number, r0: number, c0: number, block: Block, shift: 'rows' | 'cols' | 'down' | 'right'): Range {
  const sheet = tx.book.sheets[si];
  if (shift === 'rows') {
    insertDelete(tx, si, 'row', r0, block.rows);
    // 新行不沿用上一行的格式，全用来源的
    for (let i = 0; i < block.rows; i++) delete sheet.blankStyles?.[r0 + i];
    const defs = (sheet.rowDefs ??= {});
    for (let i = 0; i < block.rows; i++) {
      // 分组级别沿用插入处；行高、整行格式、隐藏用来源的
      const src = block.rowDefs?.[i];
      const level = defs[r0 + i]?.level;
      const d: RowDef = { ...(src && { ...src }), ...(level ? { level } : {}) };
      if (!level) delete d.level;
      delete d.collapsed;
      if (Object.keys(d).length) defs[r0 + i] = d;
      else delete defs[r0 + i];
    }
    return pasteBlock(tx, si, r0, 0, block);
  }
  if (shift === 'cols') {
    insertDelete(tx, si, 'col', c0, block.cols);
    sheet.blankStyles?.forEach((row) => {
      for (let j = 0; j < block.cols; j++) delete row?.[c0 + j];
    });
    const defs = splitDefs(sheet.colDefs ?? [], c0, c0 + block.cols - 1).filter((d) => d.max < c0 || d.min > c0 + block.cols - 1);
    for (const d of block.colDefs ?? []) {
      const nd: ColDef = { ...d, min: d.min + c0, max: d.max + c0 };
      delete nd.level;
      delete nd.collapsed;
      defs.push(nd);
    }
    sheet.colDefs = mergeDefs(defs);
    return pasteBlock(tx, si, 0, c0, block);
  }
  if (shift === 'down') insertDelete(tx, si, 'row', r0, block.rows, [c0, c0 + block.cols - 1]);
  else insertDelete(tx, si, 'col', c0, block.cols, [r0, r0 + block.rows - 1]);
  return pasteBlock(tx, si, r0, c0, block);
}

export type PasteMode = 'all' | 'values' | 'formats' | 'formulas';
/** 选择性粘贴的运算：把来源数值与目标数值相加 / 相减 / 相乘 / 相除 */
export type PasteOp = 'add' | 'sub' | 'mul' | 'div';
export interface PasteOpts {
  op?: PasteOp;
  /** 来源为空的格子不覆盖目标 */
  skipBlanks?: boolean;
}

/** 运算：来源必须是数字，目标是数字或空（按 0）；除以 0 得 #DIV/0!（同 Excel）；其余情况返回 undefined = 目标不动 */
export function pasteOpValue(op: PasteOp, dest: unknown, src: unknown): number | string | undefined {
  const a = typeof dest === 'number' ? dest : dest === '' || dest === undefined || dest === null ? 0 : NaN;
  if (typeof src !== 'number' || Number.isNaN(a)) return undefined;
  if (op === 'add') return a + src;
  if (op === 'sub') return a - src;
  if (op === 'mul') return a * src;
  return src === 0 ? '#DIV/0!' : a / src;
}

/** 粘贴块到 (r0, c0)；transpose 转置。返回粘贴后的区域 */
export function pasteBlock(tx: Tx, si: number, r0: number, c0: number, block: Block, mode: PasteMode = 'all', transpose = false, opts: PasteOpts = {}): Range {
  const book = tx.book;
  const sheet = book.sheets[si];
  const rows = transpose ? block.cols : block.rows;
  const cols = transpose ? block.rows : block.cols;
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      const src = transpose ? block.cells[j]?.[i] : block.cells[i]?.[j];
      const r = r0 + i;
      const c = c0 + j;
      if (!block.internal) {
        setInput(tx, si, r, c, src ? String(src.v) : '');
        continue;
      }
      tx.touch(si, r, c);
      const old = sheet.cells[r]?.[c];
      const oldS = old ? old.s : styleAt(sheet, r, c);
      if (opts.skipBlanks && (!src || src.v === '') && !src?.f) continue;
      if (opts.op) {
        // 运算只改数值，格式不动；目标是公式、或双方不是数字时保持原样
        if (old?.f) continue;
        const v = pasteOpValue(opts.op, old?.v, src?.v);
        if (v === undefined) continue;
        const cell: SCell = { ...(old ?? { v: '', t: '' }), v, t: '' };
        delete cell.ss;
        if (oldS && !cell.s) cell.s = oldS;
        refreshText(book, cell);
        putCell(sheet, r, c, cell, undefined);
        continue;
      }
      if (mode === 'formats') {
        const s = src?.s;
        if (old) {
          const next = { ...old, s };
          if (!s) delete next.s;
          refreshText(book, next);
          putCell(sheet, r, c, next, undefined);
        } else putCell(sheet, r, c, undefined, s || undefined);
        continue;
      }
      const keepF = mode === 'all' || mode === 'formulas';
      const s = mode === 'values' || mode === 'formulas' ? oldS : src?.s;
      if (!src || (src.v === '' && !(src.f && keepF))) {
        putCell(sheet, r, c, undefined, s || undefined);
        continue;
      }
      const cell: SCell = { v: src.v, t: src.t };
      if (src.k) cell.k = src.k;
      if (src.f && keepF) {
        cell.f = moveFormula(src.f, r - src.r, c - src.c);
        if (src.dyn) cell.dyn = true;
      }
      if (s) cell.s = s;
      refreshText(book, cell);
      putCell(sheet, r, c, cell, undefined);
    }
  }
  const g = { r0, c0, r1: r0 + rows - 1, c1: c0 + cols - 1 };
  if (block.internal && (mode === 'all' || mode === 'formats') && !opts.op && !transpose) {
    // 目标区域里原有的合并去掉，换成来源的合并
    const keep = sheet.merges.filter((m) => m.r1 < g.r0 || m.r0 > g.r1 || m.c1 < g.c0 || m.c0 > g.c1);
    const add = (block.merges ?? []).map((m) => ({ r0: m.r0 + r0, c0: m.c0 + c0, r1: m.r1 + r0, c1: m.c1 + c0 }));
    if (keep.length !== sheet.merges.length || add.length) {
      tx.meta(si);
      sheet.merges = [...keep, ...add];
    }
  }
  return g;
}

/** Tab / 换行分隔的文字（从 Excel、WPS、记事本复制）→ 块 */
export function blockFromText(text: string): Block {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  const src = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"' && src[i + 1] === '"') {
        field += '"';
        i++;
      } else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"' && field === '') quoted = true;
    else if (ch === '\t') {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += ch;
  }
  row.push(field);
  rows.push(row);
  const cols = Math.max(...rows.map((x) => x.length));
  return {
    rows: rows.length,
    cols,
    cells: rows.map((x, r) => Array.from({ length: cols }, (_, c) => (x[c] ? { v: x[c], t: x[c], r, c } : undefined))),
    internal: false,
  };
}

// ---------- 行列 ----------

function shiftKeys<T>(rec: Record<number, T>, at: number, n: number): Record<number, T> {
  const out: Record<number, T> = {};
  for (const [k, v] of Object.entries(rec)) {
    const i = Number(k);
    if (n < 0 && i >= at && i < at - n) continue;
    out[i >= at ? i + n : i] = v;
  }
  return out;
}

function shiftColDefs(defs: ColDef[], at: number, n: number): ColDef[] {
  const out: ColDef[] = [];
  for (const d of defs) {
    const span = shiftSpan(d.min, d.max, at, n);
    if (!span) continue;
    // 插入点在中间的定义：插入的列沿用它（区域变长）
    out.push({ ...d, min: span[0], max: Math.min(span[1], 16383) });
  }
  return out.filter((d) => d.min <= d.max);
}

/** 一个区域插入 / 删除行列后的位置；整个被删掉返回 null */
export function shiftRange(g: Range, axis: 'row' | 'col', at: number, n: number, band?: [number, number]): Range | null {
  if (band && !(axis === 'row' ? withinBand(g.c0, g.c1, band) : withinBand(g.r0, g.r1, band))) return g;
  const span = axis === 'row' ? shiftSpan(g.r0, g.r1, at, n) : shiftSpan(g.c0, g.c1, at, n);
  if (!span) return null;
  return axis === 'row' ? { ...g, r0: span[0], r1: span[1] } : { ...g, c0: span[0], c1: span[1] };
}

function shiftRanges(list: Range[], axis: 'row' | 'col', at: number, n: number, band?: [number, number]): Range[] {
  const out: Range[] = [];
  for (const m of list) {
    if (band && !(axis === 'row' ? withinBand(m.c0, m.c1, band) : withinBand(m.r0, m.r1, band))) {
      out.push(m);
      continue;
    }
    const span = axis === 'row' ? shiftSpan(m.r0, m.r1, at, n) : shiftSpan(m.c0, m.c1, at, n);
    if (!span) continue;
    const g = axis === 'row' ? { ...m, r0: span[0], r1: span[1] } : { ...m, c0: span[0], c1: span[1] };
    if (g.r0 === g.r1 && g.c0 === g.c1) continue;
    out.push(g);
  }
  return out;
}

/**
 * 插入（n > 0）/ 删除（n < 0）行或列：格子、行列定义、合并区域跟着移动；
 * 所有工作表里指向这些位置的公式引用跟着改（删掉的变 #REF!）；插入的行 / 列沿用上一行 / 左一列的格式（同 Excel）。
 */
/** 插入 / 删除单元格：band 里的格子（值和格式）沿 axis 方向从 at 起移动 n 格，删掉的格子丢弃 */
function shiftCells(sheet: SheetData, axis: 'row' | 'col', at: number, n: number, band: [number, number]) {
  const move = <T,>(grid: (T[] | undefined)[]) => {
    const moved: [number, number, T][] = [];
    if (axis === 'row') {
      for (let r = at; r < grid.length; r++) {
        const row = grid[r];
        if (!row) continue;
        for (let c = band[0]; c <= band[1] && c < row.length; c++) {
          if (row[c] === undefined) continue;
          const to = n < 0 && r < at - n ? -1 : r + n;
          if (to >= 0) moved.push([to, c, row[c]]);
          delete row[c];
        }
      }
    } else {
      for (let r = band[0]; r <= band[1] && r < grid.length; r++) {
        const row = grid[r];
        if (!row) continue;
        for (let c = at; c < row.length; c++) {
          if (row[c] === undefined) continue;
          const to = n < 0 && c < at - n ? -1 : c + n;
          if (to >= 0) moved.push([r, to, row[c]]);
          delete row[c];
        }
      }
    }
    for (const [r, c, v] of moved) (grid[r] ??= [])[c] = v;
  };
  move(sheet.cells);
  if (sheet.blankStyles) move(sheet.blankStyles);
}

/**
 * 插入 / 删除单元格能不能做：合并区域横跨 band 的边、或在移动方向上只被删掉一部分时不行（Excel 也拒绝）。
 * 返回原因，可以做时返回 null。
 */
export function cellShiftBlocked(sheet: SheetData, axis: 'row' | 'col', at: number, n: number, band: [number, number]): string | null {
  for (const m of sheet.merges) {
    const [lo, hi] = axis === 'row' ? [m.c0, m.c1] : [m.r0, m.r1];
    const [a, b] = axis === 'row' ? [m.r0, m.r1] : [m.c0, m.c1];
    const moves = b >= at; // 会被移动或删除的合并区域
    if (moves && hi >= band[0] && lo <= band[1] && (lo < band[0] || hi > band[1])) return '此操作会移动合并单元格的一部分，无法完成';
    if (n < 0 && hi >= band[0] && lo <= band[1] && a < at - n && b >= at && (a < at || b >= at - n)) return '此操作会删除合并单元格的一部分，无法完成';
  }
  return null;
}

/**
 * 插入（n > 0）/ 删除（n < 0）行列。band 有值时是插入 / 删除单元格：只有另一方向上 band 范围里的格子
 * 下移 / 上移（axis = 'row'）或右移 / 左移（axis = 'col'），行高列宽不变；引用、合并、条件格式等区域完全在 band 里的才跟着移动。
 */
export function insertDelete(tx: Tx, si: number, axis: 'row' | 'col', at: number, n: number, band?: [number, number]): void {
  const book = tx.book;
  const sheet = book.sheets[si];
  tx.snap(si);
  const spec: ShiftSpec = { sheet: sheet.name, axis, at, n, ...(band && { band }) };
  if (band) shiftCells(sheet, axis, at, n, band);
  else if (axis === 'row') {
    const fill = (arr: unknown[]) => {
      if (arr.length <= at) return;
      if (n > 0) arr.splice(at, 0, ...Array(n));
      else arr.splice(at, -n);
    };
    fill(sheet.cells);
    if (sheet.blankStyles) fill(sheet.blankStyles);
    sheet.rowDefs = shiftKeys(sheet.rowDefs ?? {}, at, n);
    if (n > 0 && at > 0) {
      // 沿用上一行的格式
      const above = at - 1;
      const rd = sheet.rowDefs[above];
      for (let i = 0; i < n; i++) {
        // 沿用上一行的格式和分组级别（在组里插入行，新行也在组里）
        if (rd?.s !== undefined || rd?.level) sheet.rowDefs[at + i] = { ...(rd.s !== undefined && { s: rd.s }), ...(rd.level && { level: rd.level }) };
        sheet.cells[above]?.forEach((x, c) => {
          if (x?.s) ((sheet.blankStyles ??= [])[at + i] ??= [])[c] = x.s;
        });
        sheet.blankStyles?.[above]?.forEach((s, c) => {
          if (s) ((sheet.blankStyles ??= [])[at + i] ??= [])[c] = s;
        });
      }
    }
  } else {
    const fill = (arr: unknown[] | undefined) => {
      if (!arr || arr.length <= at) return;
      if (n > 0) arr.splice(at, 0, ...Array(n));
      else arr.splice(at, -n);
    };
    sheet.cells.forEach((row) => fill(row));
    sheet.blankStyles?.forEach((row) => fill(row));
    sheet.colDefs = shiftColDefs(sheet.colDefs ?? [], at, n);
    if (n > 0 && at > 0) {
      const left = at - 1;
      sheet.cells.forEach((row, r) => {
        const s = row?.[left]?.s ?? sheet.blankStyles?.[r]?.[left];
        if (s) for (let i = 0; i < n; i++) ((sheet.blankStyles ??= [])[r] ??= [])[at + i] = s;
      });
    }
  }
  sheet.merges = shiftRanges(sheet.merges, axis, at, n, band);
  // 列属性跟着列走（插入 / 删除单元格只移动一部分格子，不动列属性）
  if (sheet.columns && !band) sheet.columns = shiftColumnMeta(sheet.columns, axis, at, n);
  if (sheet.autoFilter && !(band && !(axis === 'row' ? withinBand(sheet.autoFilter.ref.c0, sheet.autoFilter.ref.c1, band) : withinBand(sheet.autoFilter.ref.r0, sheet.autoFilter.ref.r1, band)))) {
    const af = sheet.autoFilter;
    const g = shiftRange(af.ref, axis, at, n);
    if (!g) delete sheet.autoFilter;
    else {
      const cols: AutoFilter['cols'] = {};
      for (const [k, v] of Object.entries(af.cols)) {
        const c = af.ref.c0 + Number(k);
        if (axis === 'col' && n < 0 && c >= at && c < at - n) continue;
        cols[(axis === 'col' && c >= at ? c + n : c) - g.c0] = v;
      }
      sheet.autoFilter = { ref: g, cols };
    }
  }
  if (sheet.page?.titleRows && axis === 'row' && !band) {
    const span = shiftSpan(sheet.page.titleRows[0], sheet.page.titleRows[1], at, n);
    if (span) sheet.page.titleRows = span;
    else delete sheet.page.titleRows;
  }
  // 手动分页符跟着行列移动（被删掉的行列上的分页符移到删除处）
  const bk = axis === 'row' ? 'rowBreaks' : 'colBreaks';
  if (sheet.page?.[bk] && !band) {
    const moved = sheet.page[bk]!.map((id) => (id < at ? id : n < 0 && id < at - n ? at : id + n)).filter((id) => id > 0);
    sheet.page = { ...sheet.page, [bk]: [...new Set(moved)].sort((a, b) => a - b) };
  }
  if (sheet.page?.area) {
    const g = shiftRange(sheet.page.area, axis, at, n, band);
    if (g) sheet.page.area = g;
    else delete sheet.page.area;
  }
  // 条件格式、数据验证的区域（显示用；文件里的由保存时的平移记录处理）
  const shiftAll = (gs: Range[]) => gs.map((g) => shiftRange(g, axis, at, n, band)).filter((g): g is Range => !!g);
  if (sheet.cf) sheet.cf = sheet.cf.map((x) => ({ ...x, ranges: shiftAll(x.ranges) })).filter((x) => x.ranges.length);
  if (sheet.validations) sheet.validations = sheet.validations.map((x) => ({ ...x, ranges: shiftAll(x.ranges) })).filter((x) => x.ranges.length);
  if (sheet.links) sheet.links = sheet.links.flatMap((x) => shiftAll([x.ref]).map((ref) => ({ ...x, ref })));
  const at1 = (i: number) => (n < 0 && i >= at && i < at - n ? -1 : i >= at ? i + n : i);
  const noteIn = (x: { r: number; c: number }) => !band || (axis === 'row' ? withinBand(x.c, x.c, band) : withinBand(x.r, x.r, band));
  if (sheet.notes) sheet.notes = sheet.notes.map((x) => (!noteIn(x) ? x : axis === 'row' ? { ...x, r: at1(x.r) } : { ...x, c: at1(x.c) })).filter((x) => x.r >= 0 && x.c >= 0);
  // 图片、图表随单元格移动（被删掉的行列里的锚点贴到删除处）
  const mv = (a: { c: number; r: number }) => {
    const k = axis === 'row' ? 'r' : 'c';
    const i = at1(a[k]);
    return { ...a, [k]: i < 0 ? at : i };
  };
  if (sheet.drawings && !band) sheet.drawings = sheet.drawings.map((d) => ({ ...d, from: { ...d.from, ...mv(d.from) }, ...(d.to && { to: { ...d.to, ...mv(d.to) } }) }));
  // 图表引用的区域跟着变（所有工作表里指向这张表的系列引用）；整个被删掉的引用去掉，图表里留着缓存值
  const shiftRef = (ref: string | undefined) => {
    if (!ref) return ref;
    const next = shiftFormula(ref, undefined, spec);
    return /#REF!/.test(next) ? undefined : next;
  };
  if (!band) {
    for (const s of book.sheets) {
      if (!s.drawings?.some((d) => d.chart)) continue;
      s.drawings = s.drawings.map((d) => (!d.chart ? d : { ...d, chart: { ...d.chart, series: d.chart.series.map((x) => ({ ...x, nameRef: shiftRef(x.nameRef), catRef: shiftRef(x.catRef), valRef: shiftRef(x.valRef) })) } }));
    }
  }
  (sheet.shifts ??= []).push({ axis, at, n, ...(band && { band }) });
  // 公式引用：本表用快照记录，其它表逐格记录
  book.sheets.forEach((s, i) => {
    s.cells.forEach((row, r) => {
      row?.forEach((cell, c) => {
        if (!cell?.f && !cell?.fa) return;
        const f = cell.f ? shiftFormula(cell.f, s.name, spec) : undefined;
        const fa = cell.fa && i === si ? shiftSqref(cell.fa, spec) || undefined : cell.fa;
        if (f === cell.f && fa === cell.fa) return;
        if (i !== si) tx.touch(i, r, c);
        const next: SCell = { ...cell };
        if (f) next.f = f;
        else delete next.f;
        if (fa) next.fa = fa;
        else delete next.fa;
        s.cells[r][c] = next;
      });
    });
  });
  tx.changed.push({ si, r: 0, c: 0 });
}

/** 设置列宽（像素，0 = 隐藏）或取消隐藏（px = undefined 且 unhide） */
export function setColWidth(tx: Tx, si: number, c0: number, c1: number, opt: { width?: number; hidden?: boolean }): void {
  const sheet = tx.book.sheets[si];
  tx.meta(si);
  const defs = splitDefs(sheet.colDefs ?? [], c0, c1);
  for (let c = c0; c <= c1; c++) {
    let d = defs.find((x) => x.min === c && x.max === c);
    if (!d) defs.push((d = { min: c, max: c }));
    if (opt.width !== undefined) {
      d.width = Math.round((opt.width / 7) * 100) / 100;
      d.hidden = undefined;
      d.extra = d.extra?.replace(/\s*bestFit="[^"]*"/, '') || undefined;
    }
    if (opt.hidden !== undefined) d.hidden = opt.hidden || undefined;
  }
  sheet.colDefs = mergeDefs(defs);
}

/** 把覆盖 [c0, c1] 的列定义拆成单列，方便逐列修改 */
export function splitDefs(defs: ColDef[], c0: number, c1: number): ColDef[] {
  const out: ColDef[] = [];
  for (const d of defs) {
    if (d.max < c0 || d.min > c1) {
      out.push({ ...d });
      continue;
    }
    if (d.min < c0) out.push({ ...d, max: c0 - 1 });
    for (let c = Math.max(d.min, c0); c <= Math.min(d.max, c1); c++) out.push({ ...d, min: c, max: c });
    if (d.max > c1) out.push({ ...d, min: c1 + 1 });
  }
  return out;
}

/** 相邻且属性完全一样的列定义合并成一段 */
export function mergeDefs(defs: ColDef[]): ColDef[] {
  const sorted = defs.filter((d) => d.width !== undefined || d.hidden || d.style !== undefined || d.extra || d.level || d.collapsed).sort((a, b) => a.min - b.min);
  const out: ColDef[] = [];
  const same = (a: ColDef, b: ColDef) => a.width === b.width && !!a.hidden === !!b.hidden && a.style === b.style && a.extra === b.extra && (a.level ?? 0) === (b.level ?? 0) && !!a.collapsed === !!b.collapsed;
  for (const d of sorted) {
    const last = out[out.length - 1];
    if (last && last.max + 1 === d.min && same(last, d)) last.max = d.max;
    else out.push({ ...d });
  }
  return out;
}

/** 设置行高（像素）或隐藏 / 取消隐藏 */
export function setRowHeight(tx: Tx, si: number, r0: number, r1: number, opt: { height?: number | null; hidden?: boolean }): void {
  const sheet = tx.book.sheets[si];
  tx.meta(si);
  const defs = (sheet.rowDefs ??= {});
  for (let r = r0; r <= r1; r++) {
    const d: RowDef = { ...defs[r] };
    if (opt.height === null) {
      delete d.ht;
      delete d.custom;
    } else if (opt.height !== undefined) {
      d.ht = Math.round(opt.height * 0.75 * 4) / 4;
      d.custom = true;
      delete d.hidden;
    }
    if (opt.hidden !== undefined) {
      if (opt.hidden) d.hidden = true;
      else delete d.hidden;
    }
    if (Object.keys(d).length) defs[r] = d;
    else delete defs[r];
  }
}

export function setFreeze(tx: Tx, si: number, rows: number, cols: number): void {
  tx.meta(si);
  tx.book.sheets[si].freeze = { rows, cols };
}

/** 合并单元格（只保留左上角的内容）/ 取消合并 */
export function mergeCells(tx: Tx, si: number, g: Range, merge: boolean): void {
  const sheet = tx.book.sheets[si];
  tx.meta(si);
  const rest = sheet.merges.filter((m) => m.r1 < g.r0 || m.r0 > g.r1 || m.c1 < g.c0 || m.c0 > g.c1);
  if (!merge) {
    sheet.merges = rest;
    return;
  }
  for (let r = g.r0; r <= g.r1; r++) {
    for (let c = g.c0; c <= g.c1; c++) {
      if ((r === g.r0 && c === g.c0) || !sheet.cells[r]?.[c]) continue;
      tx.touch(si, r, c);
      const s = sheet.cells[r][c]!.s;
      putCell(sheet, r, c, undefined, s || undefined);
    }
  }
  sheet.merges = [...rest, g];
}

// ---------- 填充柄 ----------

export { fillSeries } from './fill';

/** 区域内的值和公式是否有任何格子 */
export function hasContent(sheet: SheetData, g: Range): boolean {
  for (let r = g.r0; r <= Math.min(g.r1, sheet.rows - 1); r++) for (let c = g.c0; c <= Math.min(g.c1, sheet.cols - 1); c++) if (sheet.cells[r]?.[c]) return true;
  return false;
}

// ---------- 格式 ----------

/** 设置边框：每个格子按它在区域里的位置改对应的边（外侧框线只改外圈的边） */
export function applyBorders(tx: Tx, si: number, g: Range, kind: BorderKind, color?: string): void {
  const book = tx.book;
  if (!book.styleBook) return;
  const sheet = book.sheets[si];
  const cache = new Map<string, number>();
  for (let r = g.r0; r <= g.r1; r++) {
    for (let c = g.c0; c <= g.c1; c++) {
      const border = borderFor(kind, g, r, c, color);
      if (!border) continue;
      tx.touch(si, r, c);
      const cell = sheet.cells[r]?.[c];
      const base = cell ? cell.s : styleAt(sheet, r, c);
      const key = `${base}|${JSON.stringify(border)}`;
      let s = cache.get(key);
      if (s === undefined) cache.set(key, (s = deriveStyle(book.styleBook, base, { border })));
      if (cell) {
        const next: SCell = { ...cell, s };
        if (!s) delete next.s;
        putCell(sheet, r, c, next, undefined);
      } else putCell(sheet, r, c, undefined, s || undefined);
    }
  }
}

/** 字号变大后，没有手动设置行高的行跟着变高（同 Excel 的自动行高；只增不减） */
export function growRows(tx: Tx, si: number, r0: number, r1: number, pt: number): void {
  const sheet = tx.book.sheets[si];
  const defPt = Math.round(sheet.defaultRowHeight * 0.75 * 4) / 4;
  for (let r = r0; r <= r1; r++) {
    const d = sheet.rowDefs?.[r];
    if (d?.custom || d?.hidden || pt <= (d?.ht ?? defPt)) continue;
    tx.meta(si);
    (sheet.rowDefs ??= {})[r] = { ...d, ht: pt };
  }
}

/** 跨越合并：区域里每一行各自合并 */
export function mergeAcross(tx: Tx, si: number, g: Range): void {
  for (let r = g.r0; r <= g.r1; r++) mergeCells(tx, si, { r0: r, c0: g.c0, r1: r, c1: g.c1 }, true);
}

/** 区域里除左上角以外有内容的格子数（合并前提示会丢掉内容） */
export function lostByMerge(sheet: SheetData, g: Range): number {
  let n = 0;
  for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) if ((r !== g.r0 || c !== g.c0) && sheet.cells[r]?.[c]) n++;
  return n;
}

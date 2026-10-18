/**
 * 工作表管理：新增、删除、改名、复制、移动、隐藏、标签颜色。
 * 都是整本操作（Tx.wholeBook 记整本快照），引用这些表的公式与定义名称跟着改（改名换表名，删除变 #REF!）。
 */
import { mapRefs, renameSheetRefs } from './formula/refs';
import { DEFAULT_COL_WIDTH, MIN_ROW_HEIGHT, type SheetData, type Workbook } from './workbook';
import type { Tx } from './ops';

/** 表名是否合法（同 Excel：1–31 个字，不能有 \ / ? * [ ] :，首尾不能是 '，不重名），合法返回 null */
export function sheetNameError(book: Workbook, name: string, except = -1): string | null {
  const n = name.trim();
  if (!n) return '工作表名称不能为空';
  if ([...n].length > 31) return '工作表名称不能超过 31 个字';
  if (/[\\/?*[\]:]/.test(n)) return '工作表名称不能包含 \\ / ? * [ ] :';
  if (n.startsWith("'") || n.endsWith("'")) return "工作表名称不能以 ' 开头或结尾";
  if (n.toLowerCase() === 'history') return '“History”是 Excel 保留的名称';
  if (book.sheets.some((s, i) => i !== except && s.name.toLowerCase() === n.toLowerCase())) return '已有同名的工作表';
  return null;
}

/** 不重名的新表名：Sheet2、Sheet3 … 或“名称 (2)” */
export function uniqueSheetName(book: Workbook, base?: string): string {
  const taken = new Set(book.sheets.map((s) => s.name.toLowerCase()));
  if (base) {
    const stem = base.replace(/ \(\d+\)$/, '');
    for (let i = 2; ; i++) {
      const suffix = ` (${i})`;
      const name = [...stem].slice(0, 31 - suffix.length).join('') + suffix;
      if (!taken.has(name.toLowerCase())) return name;
    }
  }
  for (let i = book.sheets.length + 1; ; i++) if (!taken.has(`sheet${i}`)) return `Sheet${i}`;
}

function blankSheet(name: string): SheetData {
  return {
    name,
    cells: [],
    rows: 0,
    cols: 0,
    merges: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: DEFAULT_COL_WIDTH,
    defaultRowHeight: MIN_ROW_HEIGHT,
    freeze: { rows: 0, cols: 0 },
    colDefs: [],
    rowDefs: {},
    dirty: true,
  };
}

/** 所有公式（含定义名称）逐个改写 */
export function eachFormula(book: Workbook, fn: (f: string) => string) {
  for (const s of book.sheets) {
    s.cells.forEach((row) =>
      row?.forEach((cell) => {
        if (!cell?.f) return;
        const f = fn(cell.f);
        if (f !== cell.f) {
          cell.f = f;
          s.dirty = true;
        }
      }),
    );
  }
  for (const n of book.names ?? []) n.f = fn(n.f);
}

/** 定义名称的 localSheetId 跟着表的位置变（move = 旧位置 → 新位置，返回 -1 表示删除） */
function remapNames(book: Workbook, move: (i: number) => number) {
  if (!book.names) return;
  book.names = book.names
    .map((n) => (n.local === undefined ? n : { ...n, local: move(n.local) }))
    .filter((n) => n.local === undefined || n.local >= 0);
}

export function addSheet(tx: Tx, at: number, name?: string): number {
  const book = tx.book;
  tx.wholeBook();
  const i = Math.max(0, Math.min(at, book.sheets.length));
  book.sheets.splice(i, 0, blankSheet(name ?? uniqueSheetName(book)));
  remapNames(book, (k) => (k >= i ? k + 1 : k));
  return i;
}

export function deleteSheet(tx: Tx, i: number): boolean {
  const book = tx.book;
  if (book.sheets.filter((s) => !s.hidden).length <= 1 && !book.sheets[i].hidden) return false;
  tx.wholeBook();
  const name = book.sheets[i].name.toLowerCase();
  book.sheets.splice(i, 1);
  eachFormula(book, (f) => mapRefs(f, (ref) => (ref.sheet !== undefined && ref.sheet.toLowerCase() === name ? null : undefined)));
  remapNames(book, (k) => (k === i ? -1 : k > i ? k - 1 : k));
  return true;
}

export function renameSheet(tx: Tx, i: number, name: string): void {
  const book = tx.book;
  const old = book.sheets[i].name;
  const next = name.trim();
  if (old === next) return;
  tx.wholeBook();
  book.sheets[i].name = next;
  book.sheets[i].dirty = true;
  eachFormula(book, (f) => renameSheetRefs(f, old, next));
}

/** 复制工作表，放在它后面；返回新表的位置 */
export function copySheet(tx: Tx, i: number): number {
  const book = tx.book;
  tx.wholeBook();
  const src = book.sheets[i];
  const copy: SheetData = structuredClone(src);
  copy.name = uniqueSheetName(book, src.name);
  copy.copyOf = src.path ?? src.copyOf;
  // 复制出来的表没有表格部件：筛选当作普通工作表的自动筛选
  delete copy.tableFilter;
  delete copy.tables;
  if (copy.autoFilter) delete copy.autoFilter.table;
  delete copy.path;
  delete copy.rid;
  delete copy.sheetId;
  delete copy.origIndex;
  delete copy.hidden;
  copy.shifts = [];
  copy.dirty = true;
  // 复制的表里，指向原表名的引用改成指向新表（同 Excel）
  copy.cells.forEach((row) => row?.forEach((cell) => cell?.f && (cell.f = renameSheetRefs(cell.f, src.name, copy.name))));
  book.sheets.splice(i + 1, 0, copy);
  // 只属于原表的定义名称也复制一份
  const locals = (book.names ?? []).filter((n) => n.local === i).map((n) => ({ ...n, local: i + 1, f: renameSheetRefs(n.f, src.name, copy.name) }));
  remapNames(book, (k) => (k > i ? k + 1 : k));
  if (locals.length) book.names = [...(book.names ?? []), ...locals];
  return i + 1;
}

export function moveSheet(tx: Tx, from: number, to: number): void {
  const book = tx.book;
  if (from === to) return;
  tx.wholeBook();
  const [s] = book.sheets.splice(from, 1);
  book.sheets.splice(to, 0, s);
  remapNames(book, (k) => {
    if (k === from) return to;
    if (from < to && k > from && k <= to) return k - 1;
    if (from > to && k >= to && k < from) return k + 1;
    return k;
  });
}

export function setSheetHidden(tx: Tx, i: number, hidden: boolean): boolean {
  const book = tx.book;
  if (hidden && book.sheets.filter((s) => !s.hidden).length <= 1) return false;
  tx.wholeBook();
  if (hidden) book.sheets[i].hidden = 'hidden';
  else delete book.sheets[i].hidden;
  return true;
}

export function setTabColor(tx: Tx, i: number, color: string | null): void {
  tx.wholeBook();
  const s = tx.book.sheets[i];
  if (color) s.tabColor = color;
  else delete s.tabColor;
  s.dirty = true;
}

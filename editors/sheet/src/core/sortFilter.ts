/**
 * 排序与自动筛选（同 Excel）：
 * - 排序：按一列或多列，升序 / 降序；空白总在最后；数字 < 文字 < 布尔 < 错误；中文按拼音；格式跟着整行一起移动；
 *   公式按复制的规则平移（相对引用随行变化）。区域里有合并单元格时不排（Excel 也要求合并单元格大小一致）。
 * - 自动筛选：标题行每列一个条件（勾选值或自定义条件），不满足的行隐藏；条件和区域写回 <autoFilter>。
 */
import { moveFormula } from './formula/refs';
import { isDateFormat } from './numfmt';
import { formatCode } from './styles';
import { styleAt, type ColumnFilter, type DateGroup, type Range, type SCell, type SheetData, type Workbook } from './workbook';
import type { Tx } from './ops';

/** 筛选时要知道的格子信息：是不是日期（数字 + 日期格式）、底色、字体颜色 */
export interface FilterEnv {
  date(r: number, c: number): { y: number; m: number; d: number } | null;
  fill(r: number, c: number): string | null;
  color(r: number, c: number): string | null;
}

/** 日期序列号 → 年月日（1900 日期系统含 Excel 的 1900-2-29） */
export function serialYmd(serial: number, date1904 = false): { y: number; m: number; d: number } {
  let s = Math.floor(serial);
  if (!date1904) {
    if (s === 60) return { y: 1900, m: 2, d: 29 };
    if (s < 60) s += 1;
  }
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const dt = new Date(base + s * 86400000);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}

export function filterEnv(book: Workbook, si: number): FilterEnv {
  const sheet = book.sheets[si];
  const st = (r: number, c: number) => {
    const s = styleAt(sheet, r, c);
    return s === undefined ? undefined : book.styles[s];
  };
  return {
    date: (r, c) => {
      const cell = sheet.cells[r]?.[c];
      if (!cell || typeof cell.v !== 'number' || !book.styleBook) return null;
      return isDateFormat(formatCode(book.styleBook, cell.s)) ? serialYmd(cell.v, book.date1904) : null;
    },
    fill: (r, c) => st(r, c)?.fill?.toLowerCase() ?? null,
    color: (r, c) => st(r, c)?.color?.toLowerCase() ?? null,
  };
}

const inGroup = (d: { y: number; m: number; d: number }, g: DateGroup) => g.y === d.y && (g.m === undefined || g.m === d.m) && (g.d === undefined || g.d === d.d);

const collator = new Intl.Collator('zh-Hans-CN', { numeric: true, sensitivity: 'base' });

/** 排序比较：空白返回 null（由调用方放到最后） */
function rank(c: SCell | undefined): number {
  if (!c || c.v === '') return 9;
  if (c.k === 'e') return 3;
  if (typeof c.v === 'number') return 0;
  if (typeof c.v === 'boolean') return 2;
  return 1;
}

export function compareCells(a: SCell | undefined, b: SCell | undefined, desc = false): number {
  const ra = rank(a);
  const rb = rank(b);
  // 空白总在最后（升序降序都一样）
  if (ra === 9 || rb === 9) return ra === rb ? 0 : ra === 9 ? 1 : -1;
  let d: number;
  if (ra !== rb) d = ra - rb;
  else if (ra === 0) d = (a!.v as number) - (b!.v as number);
  else if (ra === 2) d = Number(a!.v) - Number(b!.v);
  else d = collator.compare(String(a!.v), String(b!.v));
  return desc ? -d : d;
}

/** 当前数据区域（同 Excel 的 Ctrl+A / 排序时自动选的区域）：从 p 出发向四周扩展到被空行空列包围 */
export function currentRegion(sheet: SheetData, r: number, c: number): Range {
  const has = (rr: number, cc: number) => rr >= 0 && cc >= 0 && !!sheet.cells[rr]?.[cc];
  const g = { r0: r, c0: c, r1: r, c1: c };
  for (let changed = true, guard = 0; changed && guard < 10000; guard++) {
    changed = false;
    const rowHas = (rr: number) => {
      for (let cc = g.c0 - 1; cc <= g.c1 + 1; cc++) if (has(rr, cc)) return true;
      return false;
    };
    const colHas = (cc: number) => {
      for (let rr = g.r0 - 1; rr <= g.r1 + 1; rr++) if (has(rr, cc)) return true;
      return false;
    };
    if (g.r0 > 0 && rowHas(g.r0 - 1)) (g.r0--, (changed = true));
    if (rowHas(g.r1 + 1)) (g.r1++, (changed = true));
    if (g.c0 > 0 && colHas(g.c0 - 1)) (g.c0--, (changed = true));
    if (colHas(g.c1 + 1)) (g.c1++, (changed = true));
  }
  // 区域边上全空的行列不算（起点本身是空格时）
  return g;
}

/** 第一行像标题：全是文字，且下一行同列里有不是文字的（数字、日期），或第一行是粗体 */
export function guessHeader(sheet: SheetData, g: Range, bold?: (r: number, c: number) => boolean): boolean {
  if (g.r1 <= g.r0) return false;
  let allText = true;
  let otherBelow = false;
  for (let c = g.c0; c <= g.c1; c++) {
    const h = sheet.cells[g.r0]?.[c];
    if (h && typeof h.v !== 'string') allText = false;
    const b = sheet.cells[g.r0 + 1]?.[c];
    if (b && typeof b.v !== 'string') otherBelow = true;
  }
  if (!allText) return false;
  if (otherBelow) return true;
  for (let c = g.c0; c <= g.c1; c++) if (bold?.(g.r0, c) && !bold?.(g.r0 + 1, c)) return true;
  return false;
}

export interface SortKey {
  /** 列号（工作表里的绝对列号） */
  c: number;
  desc: boolean;
}

/** 排序区域 g（header = 第一行是标题不参与）；有合并单元格时不排，返回 false */
export function sortRange(tx: Tx, si: number, g: Range, keys: SortKey[], header: boolean): boolean {
  const sheet = tx.book.sheets[si];
  const r0 = header ? g.r0 + 1 : g.r0;
  if (r0 >= g.r1 || !keys.length) return r0 <= g.r1;
  if (sheet.merges.some((m) => m.r1 >= r0 && m.r0 <= g.r1 && m.c1 >= g.c0 && m.c0 <= g.c1)) return false;
  tx.snap(si);
  const rows: { r: number; cells: (SCell | undefined)[]; blanks: (number | undefined)[] }[] = [];
  for (let r = r0; r <= g.r1; r++) {
    const cells: (SCell | undefined)[] = [];
    const blanks: (number | undefined)[] = [];
    for (let c = g.c0; c <= g.c1; c++) {
      cells.push(sheet.cells[r]?.[c]);
      blanks.push(sheet.blankStyles?.[r]?.[c]);
    }
    rows.push({ r, cells, blanks });
  }
  const sorted = [...rows].sort((a, b) => {
    for (const k of keys) {
      const d = compareCells(a.cells[k.c - g.c0], b.cells[k.c - g.c0], k.desc);
      if (d) return d;
    }
    return a.r - b.r;
  });
  sorted.forEach((row, i) => {
    const r = r0 + i;
    row.cells.forEach((cell, j) => {
      const c = g.c0 + j;
      const next = cell ? { ...cell, ...(cell.f && { f: moveFormula(cell.f, r - row.r, 0) }) } : undefined;
      if (next) (sheet.cells[r] ??= [])[c] = next;
      else if (sheet.cells[r]) sheet.cells[r][c] = undefined;
      const bs = row.blanks[j];
      if (bs !== undefined) ((sheet.blankStyles ??= [])[r] ??= [])[c] = bs;
      else if (sheet.blankStyles?.[r]) sheet.blankStyles[r][c] = undefined;
    });
  });
  return true;
}

// ---------- 自动筛选 ----------

const WILD = /[*?]/;
function wildcard(pattern: string): RegExp {
  const re = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '[\\s\\S]');
  return new RegExp(`^${re}$`, 'i');
}

/** 一个格子是否满足一列的筛选条件（按日期分组、按颜色时要 env 和格子位置） */
export function passes(cell: SCell | undefined, f: ColumnFilter, env?: FilterEnv, r = 0, c = 0): boolean {
  const text = cell?.t ?? '';
  if (f.color) {
    if (!env) return true;
    const got = f.color.cell ? env.fill(r, c) : env.color(r, c);
    return (got ?? null) === (f.color.color?.toLowerCase() ?? null);
  }
  if (f.dates?.length || f.values) {
    // 日期格子按年 / 月 / 日分组勾选（同 Excel），其它格子按显示文字
    const d = f.dates?.length && env ? env.date(r, c) : null;
    if (d) return f.dates!.some((g) => inGroup(d, g));
    return !!f.values?.includes(text);
  }
  if (!f.custom?.rules.length) return true;
  const test = (rule: NonNullable<ColumnFilter['custom']>['rules'][number]) => {
    const num = Number(rule.val);
    const numeric = rule.val.trim() !== '' && Number.isFinite(num) && typeof cell?.v === 'number';
    if (rule.op === 'equal' || rule.op === 'notEqual') {
      const eq = numeric ? cell!.v === num : WILD.test(rule.val) ? wildcard(rule.val).test(text) : text.toLowerCase() === rule.val.toLowerCase();
      return rule.op === 'equal' ? eq : !eq;
    }
    if (!numeric && (typeof cell?.v === 'number' || !cell)) return false;
    const d = numeric ? (cell!.v as number) - num : collator.compare(text, rule.val);
    return rule.op === 'greaterThan' ? d > 0 : rule.op === 'greaterThanOrEqual' ? d >= 0 : rule.op === 'lessThan' ? d < 0 : d <= 0;
  };
  return f.custom.and ? f.custom.rules.every(test) : f.custom.rules.some(test);
}

/** 筛选区域一列里出现的值（显示文字）与次数，排序后给下拉列表用；'' 表示空白 */
export function columnValues(sheet: SheetData, ref: Range, c: number, skip?: (r: number) => boolean): { text: string; count: number }[] {
  const map = new Map<string, { cell?: SCell; count: number }>();
  for (let r = ref.r0 + 1; r <= ref.r1; r++) {
    if (skip?.(r)) continue;
    const cell = sheet.cells[r]?.[c];
    const t = cell?.t ?? '';
    const x = map.get(t);
    if (x) x.count++;
    else map.set(t, { cell, count: 1 });
  }
  return [...map.entries()].sort((a, b) => compareCells(a[1].cell, b[1].cell)).map(([text, x]) => ({ text, count: x.count }));
}

/**
 * 被其它列的筛选条件隐藏的行（同 Excel：一列的下拉列表只列出满足其它列条件的行，所以第二次筛选是在第一次的结果里选）。
 * c 是绝对列号；没有自动筛选时返回永远为 false 的函数。
 */
export function hiddenByOthers(book: Workbook, si: number, c: number): (r: number) => boolean {
  const sheet = book.sheets[si];
  const af = sheet.autoFilter;
  const others = af ? Object.entries(af.cols).filter(([k]) => af.ref.c0 + Number(k) !== c) : [];
  if (!af || !others.length) return () => false;
  const env = filterEnv(book, si);
  return (r) => !others.every(([k, f]) => passes(sheet.cells[r]?.[af.ref.c0 + Number(k)], f, env, r, af.ref.c0 + Number(k)));
}

/** 筛选列里的日期（年 → 月 → 次数）和颜色（底色 / 字体颜色 → 次数），给下拉菜单用；skip 返回 true 的行不算 */
export function columnExtras(book: Workbook, si: number, ref: Range, c: number, skip?: (r: number) => boolean): { dates: Map<number, Map<number, number>>; fills: Map<string | null, number>; colors: Map<string | null, number> } {
  const env = filterEnv(book, si);
  const dates = new Map<number, Map<number, number>>();
  const fills = new Map<string | null, number>();
  const colors = new Map<string | null, number>();
  for (let r = ref.r0 + 1; r <= ref.r1; r++) {
    if (skip?.(r)) continue;
    const d = env.date(r, c);
    if (d) {
      const ym = dates.get(d.y) ?? dates.set(d.y, new Map()).get(d.y)!;
      ym.set(d.m, (ym.get(d.m) ?? 0) + 1);
    }
    const f = env.fill(r, c);
    fills.set(f, (fills.get(f) ?? 0) + 1);
    const k = env.color(r, c);
    colors.set(k, (colors.get(k) ?? 0) + 1);
  }
  return { dates, fills, colors };
}

/** 按全部列的条件重新决定哪些行隐藏（只改筛选区域里的数据行） */
export function refilter(tx: Tx, si: number): void {
  const sheet = tx.book.sheets[si];
  const af = sheet.autoFilter;
  if (!af) return;
  tx.meta(si);
  const defs = (sheet.rowDefs ??= {});
  const env = filterEnv(tx.book, si);
  for (let r = af.ref.r0 + 1; r <= af.ref.r1; r++) {
    const show = Object.entries(af.cols).every(([k, f]) => passes(sheet.cells[r]?.[af.ref.c0 + Number(k)], f, env, r, af.ref.c0 + Number(k)));
    const d = { ...defs[r] };
    if (show) delete d.hidden;
    else d.hidden = true;
    if (Object.keys(d).length) defs[r] = d;
    else delete defs[r];
  }
}

/** 打开自动筛选（区域 g 第一行为标题）；ref = null 关闭并显示全部行 */
export function setAutoFilter(tx: Tx, si: number, ref: Range | null): void {
  const sheet = tx.book.sheets[si];
  tx.meta(si);
  if (sheet.autoFilter) {
    sheet.autoFilter.cols = {};
    refilter(tx, si);
  }
  // 表格自带的筛选：区域就是表格，不能换成选区；取消筛选也只是清掉条件（表格的筛选按钮是表格的属性）
  if (sheet.tableFilter && sheet.autoFilter?.table) return;
  if (ref) sheet.autoFilter = { ref, cols: {} };
  else delete sheet.autoFilter;
}

/** 设置一列（绝对列号 c）的条件，null 清除该列条件 */
export function setColumnFilter(tx: Tx, si: number, c: number, f: ColumnFilter | null): void {
  const sheet = tx.book.sheets[si];
  const af = sheet.autoFilter;
  if (!af) return;
  tx.meta(si);
  const cols = { ...af.cols };
  if (f) cols[c - af.ref.c0] = f;
  else delete cols[c - af.ref.c0];
  sheet.autoFilter = { ...af, cols };
  refilter(tx, si);
}

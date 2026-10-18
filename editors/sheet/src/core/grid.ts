/** 表格查看器的纯逻辑：行列坐标轴、合并单元格、选区移动、复制内容（TSV / HTML / Markdown）、统计与搜索。 */
import { colName } from '@wz/editor-shared/xlsx';
import type { BorderSide, CellBorder } from './color';
import type { CellStyle, Range, SCell, SheetData } from './workbook';

export interface Pos {
  r: number;
  c: number;
}

// ---------- 坐标轴 ----------

/** 一个方向上的尺寸：pos[i] 是第 i 行 / 列的起点，pos[n] 是总长 */
export interface Axis {
  n: number;
  pos: Float64Array;
}

export function buildAxis(n: number, sizes: Record<number, number>, def: number, zoom = 1): Axis {
  const pos = new Float64Array(n + 1);
  for (let i = 0; i < n; i++) pos[i + 1] = pos[i] + Math.round((sizes[i] ?? def) * zoom);
  return { n, pos };
}

export const sizeOf = (a: Axis, i: number) => a.pos[i + 1] - a.pos[i];

/** 偏移量所在的行 / 列（超出范围时夹到两端） */
export function indexAt(a: Axis, offset: number): number {
  if (a.n === 0) return 0;
  let lo = 0;
  let hi = a.n - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (a.pos[mid] <= offset) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/**
 * 偏移量是否落在某条行 / 列分隔线上（左右各 tol 像素）：返回分隔线左 / 上侧那一行 / 列，不在分隔线上返回 -1。
 * 分隔线处有隐藏（尺寸为 0）的行列时，取它前面最近的可见行列，和 Excel 拖动分隔线调整的对象一致。
 */
export function edgeAt(a: Axis, offset: number, tol = 3): number {
  if (a.n === 0 || offset < 0) return -1;
  const i = indexAt(a, offset);
  // 离哪条线近就选哪条：i 的起点（即 i-1 的终点）或 i 的终点
  const before = offset - a.pos[i];
  const after = Math.abs(a.pos[i + 1] - offset);
  let k = after <= tol && after <= before ? i : before <= tol ? i - 1 : -1;
  while (k >= 0 && sizeOf(a, k) === 0) k--;
  return k;
}

/** 单元格左右内边距（px，与 .sc 的 padding 一致） */
export const CELL_PAD = 6;

/**
 * 单元格的左右内边距：内边距 + 1px 右网格线不能超过列宽，否则 border-box 的格子会被撑得比列宽，
 * 右网格线压到下一格底下看不见（窄列只显示网格线，和 Excel 一样）。
 */
export function cellPadding(width: number): number {
  return Math.max(0, Math.min(CELL_PAD, Math.floor((width - 1) / 2)));
}

/**
 * 单元格的底色 / 网格线 / 文字颜色（内联样式）：
 * - 有底色时网格线也用底色（Excel 的填充盖住网格线），选区内把底色压暗一些；
 * - 有底色而字体是自动色时文字用黑色（Excel 的自动色），不随深色主题变浅；
 * - 没有底色时黑色文字按自动色处理（跟随主题的文字色），其它颜色照原样显示。
 */
export function cellColors(st: CellStyle | undefined, inSel = false): { background?: string; borderColor?: string; color?: string } {
  const out: { background?: string; borderColor?: string; color?: string } = {};
  if (st?.fill) {
    out.background = inSel ? `color-mix(in srgb, ${st.fill} 80%, #000)` : st.fill;
    out.borderColor = st.fill;
    out.color = st.color ?? '#000000';
  } else if (st?.color && st.color !== '#000000') out.color = st.color;
  return out;
}

/**
 * 单元格要画的边框：合并区域里只画落在合并区域外沿上的边（Excel 合并后按外沿各格的边框显示）。
 */
export function cellBorder(st: CellStyle | undefined, r: number, c: number, merge?: Range): CellBorder | undefined {
  const b = st?.border;
  if (!b) return undefined;
  if (!merge) return b;
  const out: CellBorder = {};
  if (b.l && c === merge.c0) out.l = b.l;
  if (b.r && c === merge.c1) out.r = b.r;
  if (b.t && r === merge.r0) out.t = b.t;
  if (b.b && r === merge.r1) out.b = b.b;
  return out.l || out.r || out.t || out.b ? out : undefined;
}

/** 自动色 / 黑色边框在无底色的格子上跟随主题的文字色（和文字颜色的处理一致） */
export function borderLine(side: BorderSide, fill?: string): string {
  const color = side.color && side.color !== '#000000' ? side.color : fill ? '#000000' : 'var(--text)';
  return `${side.w}px ${side.style} ${color}`;
}

/**
 * 一段边框线：竖线（v）在第 c 列的左边线（= 第 c - 1 列的右边线）上、占第 r 行；横线在第 r 行的上边线上、占第 c 列。
 * 相邻两格共用的边只出一段，两格都设置时取较粗的（一样粗取先出现的）。
 * 每段线都按网格线的方式摆放（压在前一格的右 / 下网格线上），在非整数缩放的屏幕上与网格线取整到同一像素，
 * 不会因为两格各画一条而错开一个像素、看起来变粗。
 */
export interface BorderEdge {
  v: boolean;
  r: number;
  c: number;
  w: number;
  css: string;
}

export function borderEdges(cells: { r: number; c: number; b: CellBorder; fill?: string }[]): BorderEdge[] {
  const map = new Map<string, BorderEdge>();
  const put = (v: boolean, r: number, c: number, side: BorderSide | undefined, fill?: string) => {
    if (!side) return;
    const k = `${v ? 'v' : 'h'}${r}:${c}`;
    const old = map.get(k);
    if (!old || side.w > old.w) map.set(k, { v, r, c, w: side.w, css: borderLine(side, fill) });
  };
  for (const { r, c, b, fill } of cells) {
    put(true, r, c, b.l, fill);
    put(true, r, c + 1, b.r, fill);
    put(false, r, c, b.t, fill);
    put(false, r + 1, c, b.b, fill);
  }
  return [...map.values()];
}

/** 显示的行列数：数据范围之外多留一些空白格，像 Excel 一样 */
export function gridExtent(sheet: SheetData): { rows: number; cols: number } {
  return {
    rows: Math.max(sheet.rows + 20, sheet.freeze.rows + 1, 60),
    cols: Math.max(sheet.cols + 4, sheet.freeze.cols + 1, 16),
  };
}

// ---------- 区域与合并单元格 ----------

export function rangeOf(a: Pos, b: Pos): Range {
  return { r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) };
}

export const contains = (g: Range, r: number, c: number) => r >= g.r0 && r <= g.r1 && c >= g.c0 && c <= g.c1;
const overlaps = (a: Range, b: Range) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** 每个被合并覆盖的单元格 → 所在合并区域 */
export type MergeIndex = Map<number, Range>;
export const cellKey = (r: number, c: number) => r * 16384 + c;

export function buildMergeIndex(merges: Range[]): MergeIndex {
  const idx: MergeIndex = new Map();
  for (const m of merges) for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) idx.set(cellKey(r, c), m);
  return idx;
}

/** 把区域扩大到完整包含与它相交的合并单元格 */
export function expandToMerges(range: Range, merges: Range[]): Range {
  let g = { ...range };
  for (let changed = true; changed; ) {
    changed = false;
    for (const m of merges) {
      if (!overlaps(g, m) || (contains(g, m.r0, m.c0) && contains(g, m.r1, m.c1))) continue;
      g = { r0: Math.min(g.r0, m.r0), c0: Math.min(g.c0, m.c0), r1: Math.max(g.r1, m.r1), c1: Math.max(g.c1, m.c1) };
      changed = true;
    }
  }
  return g;
}

/** 合并单元格内的位置 → 合并区域的左上角 */
export function snap(p: Pos, idx: MergeIndex): Pos {
  const m = idx.get(cellKey(p.r, p.c));
  return m ? { r: m.r0, c: m.c0 } : p;
}

// ---------- 移动 ----------

export interface MoveCtx {
  sheet: SheetData;
  rows: number;
  cols: number;
  merges: MergeIndex;
  hiddenRow: (r: number) => boolean;
  hiddenCol: (c: number) => boolean;
}

const has = (s: SheetData, r: number, c: number) => !!s.cells[r]?.[c];

/** 方向键移动一格（跳过隐藏行列和同一个合并单元格）；jump = Ctrl+方向键，跳到数据块边缘 */
export function moveBy(ctx: MoveCtx, from: Pos, dr: number, dc: number, jump = false): Pos {
  const inb = (r: number, c: number) => r >= 0 && c >= 0 && r < ctx.rows && c < ctx.cols;
  const hidden = (r: number, c: number) => ctx.hiddenRow(r) || ctx.hiddenCol(c);
  // 从合并单元格移出时，从它的边缘开始走
  const m = ctx.merges.get(cellKey(from.r, from.c));
  let r = m ? (dr > 0 ? m.r1 : dr < 0 ? m.r0 : from.r) : from.r;
  let c = m ? (dc > 0 ? m.c1 : dc < 0 ? m.c0 : from.c) : from.c;
  const stepOnce = (): boolean => {
    let nr = r + dr;
    let nc = c + dc;
    while (inb(nr, nc) && hidden(nr, nc)) {
      nr += dr;
      nc += dc;
    }
    if (!inb(nr, nc)) return false;
    r = nr;
    c = nc;
    return true;
  };
  if (!jump) {
    stepOnce();
    return snap({ r, c }, ctx.merges);
  }
  const s = ctx.sheet;
  const startFilled = has(s, r, c);
  if (!stepOnce()) return snap({ r, c }, ctx.merges);
  if (startFilled && has(s, r, c)) {
    // 在数据块里：走到块的最后一个非空格
    for (;;) {
      const pr = r;
      const pc = c;
      if (!stepOnce()) break;
      if (!has(s, r, c)) {
        r = pr;
        c = pc;
        break;
      }
    }
  } else {
    // 走到下一个非空格，没有就到边缘
    while (!has(s, r, c) && stepOnce());
  }
  return snap({ r, c }, ctx.merges);
}

// ---------- 地址 ----------

export const cellName = (p: Pos) => `${colName(p.c)}${p.r + 1}`;

export function rangeName(g: Range): string {
  const a = cellName({ r: g.r0, c: g.c0 });
  return g.r0 === g.r1 && g.c0 === g.c1 ? a : `${a}:${cellName({ r: g.r1, c: g.c1 })}`;
}

// ---------- 复制 ----------

export interface OutCell {
  text: string;
  num: boolean;
  bold: boolean;
  rowspan: number;
  colspan: number;
}

export interface CopyOptions {
  /** 复制原始值（不带千分位 / 日期格式） */
  raw?: boolean;
  hiddenRow?: (r: number) => boolean;
  hiddenCol?: (c: number) => boolean;
}

function rawText(cell: SCell): string {
  if (typeof cell.v === 'number') return String(cell.v);
  if (typeof cell.v === 'boolean') return cell.v ? 'TRUE' : 'FALSE';
  return cell.v;
}

/**
 * 区域内（跳过隐藏行列）的单元格矩阵；被合并覆盖的非左上角格为 null，
 * 左上角格带上合并跨度（按区域和隐藏行列裁剪）。
 */
export function collect(sheet: SheetData, styles: CellStyle[], range: Range, opts: CopyOptions = {}): (OutCell | null)[][] {
  const hr = opts.hiddenRow ?? (() => false);
  const hc = opts.hiddenCol ?? (() => false);
  const rows: number[] = [];
  const cols: number[] = [];
  for (let r = range.r0; r <= range.r1; r++) if (!hr(r)) rows.push(r);
  for (let c = range.c0; c <= range.c1; c++) if (!hc(c)) cols.push(c);
  const idx = buildMergeIndex(sheet.merges.filter((m) => overlaps(m, range)));
  return rows.map((r) =>
    cols.map((c) => {
      const m = idx.get(cellKey(r, c));
      const top = m ? { r: Math.max(m.r0, range.r0), c: Math.max(m.c0, range.c0) } : null;
      // 合并区域在本区域内第一个可见格才输出
      if (m && top) {
        const firstR = rows.find((x) => x >= top.r && x <= m.r1);
        const firstC = cols.find((x) => x >= top.c && x <= m.c1);
        if (r !== firstR || c !== firstC) return null;
      }
      const cell = m ? sheet.cells[m.r0]?.[m.c0] : sheet.cells[r]?.[c];
      const span = (list: number[], lo: number, hi: number) => list.filter((x) => x >= lo && x <= hi).length;
      return {
        text: cell ? (opts.raw ? rawText(cell) : cell.t) : '',
        num: typeof cell?.v === 'number',
        bold: !!(cell?.s !== undefined && styles[cell.s]?.bold),
        rowspan: m ? span(rows, r, m.r1) : 1,
        colspan: m ? span(cols, c, m.c1) : 1,
      };
    }),
  );
}

const plain = (m: (OutCell | null)[][]) => m.map((row) => row.map((x) => x?.text ?? ''));

/** 粘贴到 Excel / 记事本：Tab 分列、换行分行；单个单元格不加引号 */
export function toTsv(m: (OutCell | null)[][]): string {
  const t = plain(m);
  if (t.length === 1 && t[0].length === 1) return t[0][0];
  const q = (s: string) => (/[\t\r\n"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  return t.map((row) => row.map(q).join('\t')).join('\r\n');
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** 粘贴到 Word / 邮件：保留表格、合并、粗体、数字右对齐 */
export function toHtml(m: (OutCell | null)[][]): string {
  const td = (x: OutCell) => {
    const attrs = [
      x.rowspan > 1 ? ` rowspan="${x.rowspan}"` : '',
      x.colspan > 1 ? ` colspan="${x.colspan}"` : '',
      ` style="border:1px solid #bfbfbf;padding:2px 6px;${x.num ? 'text-align:right;' : ''}${x.bold ? 'font-weight:bold;' : ''}"`,
    ].join('');
    return `<td${attrs}>${esc(x.text).replace(/\r?\n/g, '<br>')}</td>`;
  };
  const body = m.map((row) => `<tr>${row.map((x) => (x ? td(x) : '')).join('')}</tr>`).join('');
  return `<table style="border-collapse:collapse" cellspacing="0">${body}</table>`;
}

/** Markdown 表格：第一行作表头 */
export function toMarkdown(m: (OutCell | null)[][]): string {
  const t = plain(m).map((row) => row.map((s) => s.replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>')));
  if (!t.length) return '';
  const line = (row: string[]) => `| ${row.join(' | ')} |`;
  return [line(t[0]), line(t[0].map(() => '---')), ...t.slice(1).map(line)].join('\n');
}

// ---------- 统计 ----------

export interface Stats {
  /** 非空单元格数 */
  count: number;
  /** 数值个数 */
  nums: number;
  sum: number;
  avg: number;
  min: number;
  max: number;
}

/** 选区里的非空格数与数值统计（隐藏行列不算） */
export function statsOf(sheet: SheetData, range: Range, hiddenRow?: (r: number) => boolean, hiddenCol?: (c: number) => boolean): Stats {
  let count = 0;
  let nums = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  const r1 = Math.min(range.r1, sheet.rows - 1);
  for (let r = range.r0; r <= r1; r++) {
    const row = sheet.cells[r];
    if (!row || hiddenRow?.(r)) continue;
    const c1 = Math.min(range.c1, row.length - 1);
    for (let c = range.c0; c <= c1; c++) {
      const cell = row[c];
      if (!cell || hiddenCol?.(c)) continue;
      count++;
      if (typeof cell.v === 'number') {
        nums++;
        sum += cell.v;
        if (cell.v < min) min = cell.v;
        if (cell.v > max) max = cell.v;
      }
    }
  }
  // 去掉浮点累加误差（0.1 + 0.2）
  sum = Number(sum.toPrecision(15));
  return { count, nums, sum, avg: nums ? Number((sum / nums).toPrecision(15)) : 0, min: nums ? min : 0, max: nums ? max : 0 };
}

/** 状态栏里的数字：千分位，最多 4 位小数 */
export function fmtStat(v: number): string {
  return v.toLocaleString('zh-CN', { maximumFractionDigits: 4 });
}

// ---------- 搜索 ----------

/** 显示文字包含关键字（不区分大小写）的单元格，按行优先排列 */
export function findAll(sheet: SheetData, query: string): Pos[] {
  const q = query.trim().toLowerCase();
  if (!q) return [];
  const out: Pos[] = [];
  sheet.cells.forEach((row, r) => {
    row?.forEach((cell, c) => {
      if (cell && cell.t.toLowerCase().includes(q)) out.push({ r, c });
    });
  });
  return out;
}

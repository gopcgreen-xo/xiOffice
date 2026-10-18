/**
 * 表格编辑器的数据模型与解析：.xlsx / .xlsm（值、公式、数字格式、字体 / 对齐 / 底色 / 边框、合并单元格、行高列宽、
 * 隐藏行列与工作表、冻结窗格）和 .csv。解析在 Web Worker 里进行（`parse.worker.ts`）。
 *
 * 模型只记录编辑器认识的内容；保存时（`save.ts`）在原文件上修补，其它内容（图表、条件格式等）原样保留。
 */
import { strFromU8, unzipSync } from 'fflate';
import type { SheetColumns } from '@wz/editor-shared';
import type { PivotDef } from './pivot';
import { colIndex, resolvePart, unescapeXml, xmlAttr, xmlText } from '@wz/editor-shared/xlsx';
import { formatGeneral, formatNumber } from './numfmt';
import { allCellStyles, formatCode, readStyleBook, type CellStyle, type StyleBook } from './styles';
import { resolveColor } from './color';
import { readDrawings, type SheetDrawing, type SheetLink, type SheetNote } from './drawing';
import type { SheetTable } from './tableStyle';
import { moveFormula } from './formula/refs';
import { fromFileFormula } from './formula/parse';

export type { CellStyle } from './styles';
export type Raw = string | number | boolean;

/** 只有动态数组才有的函数开头的公式（旧版 Excel 里没有这些函数，所以文件里的数组公式一定是动态数组） */
const DYNAMIC_TOP = /^(FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|TEXTSPLIT|HSTACK|VSTACK)\(/i;

export interface SCell {
  /** 原始值（数字、文字、布尔）；公式格子是计算结果 */
  v: Raw;
  /** 按数字格式显示的文字 */
  t: string;
  /** 样式序号（`Workbook.styles`） */
  s?: number;
  /** 公式（不带 =） */
  f?: string;
  /** 数组公式的区域（`<f t="array" ref>`） */
  fa?: string;
  /** 编辑器不改写的公式（如模拟运算表），保存时原样写回整个 `<f>` */
  fraw?: string;
  /** 值来自共享字符串第几项（文字没改过时保存沿用，保留富文本格式） */
  ss?: number;
  /** 特殊值类型：e = 错误值（#N/A 等），d = ISO 日期文字 */
  k?: 'e' | 'd';
  /** `<c>` 上编辑器不认识的属性（cm、vm、ph），原样写回 */
  xa?: string;
  /** 动态数组公式：结果是数组时溢出到相邻格子（`fa` 记当前溢出区域）。不写进文件 */
  dyn?: boolean;
  /** 这个格子是某个动态数组公式溢出来的结果（不是用户输入的）。不写进文件 */
  sp?: boolean;
}

/** 单元格区域（含两端，r0 ≤ r1，c0 ≤ c1） */
export interface Range {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** `<col>`：一段列（从 0 开始，含两端）的宽度（字符数）、隐藏、整列格式与其它属性 */
export interface ColDef {
  min: number;
  max: number;
  width?: number;
  hidden?: boolean;
  style?: number;
  /** 分组（大纲）级别 1–7；没有分组不写 */
  level?: number;
  /** 这一列是汇总列，它管着的那组列已折叠 */
  collapsed?: boolean;
  /** 其它属性原文（bestFit 等） */
  extra?: string;
}

/** `<row>` 的行高（磅）、隐藏、整行格式与其它属性 */
export interface RowDef {
  ht?: number;
  custom?: boolean;
  hidden?: boolean;
  s?: number;
  /** 分组（大纲）级别 1–7；没有分组不写 */
  level?: number;
  /** 这一行是汇总行，它管着的那组行已折叠 */
  collapsed?: boolean;
  extra?: string;
}

export interface SheetData {
  name: string;
  /** 按行排列的稀疏单元格 */
  cells: (SCell | undefined)[][];
  /** 数据范围：最后一个有值 / 合并的行列之后 */
  rows: number;
  cols: number;
  merges: Range[];
  /** 列宽 / 行高（像素，100% 缩放）；隐藏的行列为 0；未列出的用默认值（由 colDefs / rowDefs 推出） */
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  defaultColWidth: number;
  defaultRowHeight: number;
  /** 冻结的行数 / 列数 */
  freeze: { rows: number; cols: number };
  /** 没有值的 `<c>` 的样式序号（只为显示底色），按行排列的稀疏数组 */
  blankStyles?: (number | undefined)[][];
  /** 整行格式（`<row s customFormat="1">`）与整列格式（`<col style>`），用于文件里没有 `<c>` 的格子 */
  rowStyles?: Record<number, number>;
  colStyles?: Record<number, number>;
  /** 文件里的列定义与行属性（保存用） */
  colDefs?: ColDef[];
  rowDefs?: Record<number, RowDef>;
  /** 数据透视表的定义（这张表是透视结果）。只在编辑期间有效，不写进文件 */
  pivot?: PivotDef;
  /** 分组汇总项的位置（`<outlinePr>`）：默认汇总行在明细下方、汇总列在明细右侧 */
  outline?: { summaryBelow: boolean; summaryRight: boolean };
  /** 隐藏的工作表 */
  hidden?: 'hidden' | 'veryHidden';
  /** 在 xlsx 包里的路径（xl/worksheets/sheet1.xml）与 workbook.xml 里的 r:id；新加的表没有 */
  path?: string;
  rid?: string;
  /** 内容改过，保存时要重写 */
  dirty?: boolean;
  /** 插入 / 删除行列的记录（按顺序），保存时用来平移编辑器不建模的内容（条件格式、数据验证、定义名称等） */
  /** 插入 / 删除行列（band 有值时为插入 / 删除单元格：只移动垂直方向上 band 范围内的格子） */
  shifts?: { axis: 'row' | 'col'; at: number; n: number; band?: [number, number] }[];
  /** workbook.xml 里的 sheetId 与原来的位置（定义名称的 localSheetId 按位置算）；新加的表没有 */
  sheetId?: number;
  origIndex?: number;
  /** 文件里的表名（改名后保存时据此改写定义名称里的引用） */
  fileName?: string;
  /** 复制出来的表：保存时以这张表的 XML 为底（去掉图表等关系） */
  copyOf?: string;
  /** 工作表标签颜色（#rrggbb） */
  tabColor?: string;
  /** 工作表保护（sheetProtection sheet="1"）：只有未锁定的单元格能改 */
  protected?: boolean;
  /** `<sheetProtection>` 的全部属性（密码散列、各项权限），原样保存 */
  protection?: { attrs: Record<string, string> };
  /** 自动筛选 */
  autoFilter?: AutoFilter;
  /** 页面设置（打印用） */
  page?: PageSetup;
  /** 条件格式；文件里的规则原样保留，只有新建的写入、删除的去掉 */
  cf?: CfRule[];
  /** 删掉的文件里的规则（CfRule.idx），保存时从原 XML 里去掉 */
  cfRemoved?: number[];
  /** 数据验证（下拉列表、数值范围） */
  validations?: Validation[];
  /** 列属性（工作包里才有，见 docs/WBP_FORMAT.md）：单选 / 多选 / 日期 / 复选框等字段类型。只在编辑期间有效，不写进 xlsx，由宿主存进工作包 */
  columns?: SheetColumns;
  /** 图片、图表、批注（笔记）、超链接：只读显示 */
  drawings?: SheetDrawing[];
  /** 文件里有图表 / 图片（删光后保存也要重新读一遍） */
  hadDrawings?: boolean;
  notes?: SheetNote[];
  links?: SheetLink[];
  /** 套用表格格式的表格（显示样式；定义原样保留，只有自带的筛选条件会写回） */
  tables?: SheetTable[];
  /** 表格自带的筛选所在的表格部件路径（xl/tables/tableN.xml）、文件里原来的筛选区域（A1:G20 这样的文字） */
  tableFilter?: { path: string; ref: string };
}

/** 一条条件格式规则 */
export interface CfRule {
  /** 在文件里是第几条 cfRule（按出现顺序）；编辑器新建的规则没有，保存时才写入 */
  idx?: number;
  ranges: Range[];
  type: string;
  priority: number;
  stopIfTrue?: boolean;
  /** 满足时套用的差异格式（styleBook.dxfs 的序号） */
  dxf?: number;
  operator?: string;
  /** 公式（不带 =），相对引用以第一块区域的左上角为准 */
  formulas: string[];
  /** containsText 等的文字 */
  text?: string;
  /** top10：前 / 后几项（或百分比） */
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  /** aboveAverage：false = 低于平均 */
  above?: boolean;
  /** 色阶 / 数据条：颜色（#rrggbb）与阈值（min / max / num / percent / percentile + 值） */
  colors?: string[];
  cfvo?: { type: string; val?: string }[];
}

/** 一条数据验证 */
export interface Validation {
  ranges: Range[];
  /** list / whole / decimal / date / time / textLength / custom */
  type: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  allowBlank: boolean;
  showError: boolean;
  /** stop（拒绝） / warning / information */
  errorStyle?: string;
  error?: string;
  errorTitle?: string;
  prompt?: string;
  promptTitle?: string;
  /** 下拉箭头：showDropDown="1" 在 OOXML 里反而表示“不显示” */
  noDropDown?: boolean;
}

/** 一列的筛选条件：values = 勾选的显示文字（'' 表示空白）；custom = 自定义条件（最多两个） */
export interface ColumnFilter {
  /** 按值：显示的文字（'' = 空白）；日期格子不看它，看 dates */
  values?: string[];
  /** 按日期分组勾选（同一个 filters 里的 dateGroupItem）：年、年月或年月日 */
  dates?: DateGroup[];
  /** 按颜色：cell = 单元格底色，否则字体颜色；color 为 CSS #rrggbb，null = 无填充 / 自动色；dxf = 文件里引用的差异格式 */
  color?: { cell: boolean; color: string | null; dxf?: number };
  custom?: { and: boolean; rules: { op: 'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual'; val: string }[] };
}

export interface DateGroup {
  y: number;
  m?: number;
  d?: number;
}

export interface AutoFilter {
  /** 筛选区域（第一行是标题） */
  ref: Range;
  /** 这个筛选是表格（套用表格格式）自带的：条件写回表格部件，不写工作表的 autoFilter */
  table?: boolean;
  /** 列号（相对 ref.c0）→ 条件 */
  cols: Record<number, ColumnFilter>;
}

export interface PageSetup {
  landscape?: boolean;
  /** paperSize：9 = A4，8 = A3，11 = A5，1 = Letter */
  paper?: number;
  /** 缩放百分比（10–400） */
  scale?: number;
  /** 调整为 N 页宽 × M 页高（0 = 不限）；有它时不用 scale */
  fitW?: number;
  fitH?: number;
  gridLines?: boolean;
  /** 每页重复的标题行（从 0 开始，含两端） */
  titleRows?: [number, number];
  /** 打印区域 */
  area?: Range;
  /** 页边距（英寸，同文件里的 pageMargins）；没有时按 Excel 默认 左右 0.7、上下 0.75、页眉页脚 0.3 */
  margins?: PageMargins;
  /** 水平 / 垂直居中 */
  centerH?: boolean;
  centerV?: boolean;
  /** 页眉 / 页脚（Excel 的代码：&L &C &R 分节，&P 页码，&N 总页数，&D 日期，&A 表名，&F 文件名） */
  header?: string;
  footer?: string;
  /** 手动分页符：新的一页从这些行 / 列开始（0 起，同 brk id） */
  rowBreaks?: number[];
  colBreaks?: number[];
}

export interface PageMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

export const DEFAULT_MARGINS: PageMargins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

/**
 * 格子用哪个样式（OOXML：有 `<c>` 就用它的 s，没写 s 即 0；没有 `<c>` 时先看整行格式，再看整列格式）。
 * 返回 undefined 表示缺省样式。
 */
export function styleAt(sheet: SheetData, r: number, c: number): number | undefined {
  const cell = sheet.cells[r]?.[c];
  if (cell) return cell.s;
  const blank = sheet.blankStyles?.[r]?.[c];
  if (blank !== undefined) return blank;
  return sheet.rowStyles?.[r] ?? sheet.colStyles?.[c];
}

export interface Workbook {
  sheets: SheetData[];
  /** 每个样式序号的显示样式（与 styleBook.xfs 一一对应） */
  styles: CellStyle[];
  styleBook?: StyleBook;
  /** 1904 日期系统 */
  date1904?: boolean;
  /** 来源：xlsx（含 xlsm）、csv、xls（只能另存为 xlsx）、new（新建） */
  kind?: 'xlsx' | 'csv' | 'xls' | 'new';
  /** CSV 的分隔符与编码 */
  csv?: { delim: string; encoding: string; bom: boolean };
  /** 文件里有、编辑器原样保留但不能编辑的内容（提示用） */
  keeps?: string[];
  /** 定义名称（local = 只在第几张表里有效）；打印区域等内置名称不在这里 */
  names?: { name: string; f: string; local?: number }[];
  /** 定义名称改过（保存时要重写工作簿里的 definedNames）；保存成功后清掉 */
  namesDirty?: boolean;
  /** `<workbookProtection>` 的属性（lockStructure 等） */
  wbProtection?: { attrs: Record<string, string> };
  /** 原文件字节（保存时在它上面修补；界面线程设置，不进撤销记录） */
  source?: Uint8Array;
  /** 工作包里的表格：读包里别的表格的区域（`[条目名]工作表!A2:A20`），给列属性的选项引用用。运行期设置，不写进文件 */
  externalValues?: (ref: string) => string[];
}

/** 至少这么高，保证 13px 文字不挤 */
export const MIN_ROW_HEIGHT = 22;
export const DEFAULT_COL_WIDTH = 72;
/** 最多展开这么多列的列宽设置（`<col max="16384">`） */
const MAX_COLS = 2048;

export const colPx = (chars: number) => Math.round(chars * 7);
export const rowPx = (pt: number) => Math.round((pt * 4) / 3);
/** 像素 → 列宽字符数 / 行高磅数（保存用） */
export const pxToChars = (px: number) => Math.round((px / 7) * 100) / 100;
export const pxToPt = (px: number) => Math.round(px * 0.75 * 4) / 4;

/** A1:C3 → 区域（单个单元格也可以） */
export function parseRef(ref: string): Range | null {
  const m = ref.trim().toUpperCase().match(/^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/);
  if (!m) return null;
  const r0 = Number(m[2]) - 1;
  const c0 = colIndex(m[1]);
  const r1 = m[4] ? Number(m[4]) - 1 : r0;
  const c1 = m[3] ? colIndex(m[3]) : c0;
  if (r0 < 0 || r1 < 0) return null;
  return { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
}

/** 由 colDefs / rowDefs 推出显示用的列宽、行高、整行整列格式 */
export function deriveAxes(sheet: SheetData): void {
  const colWidths: Record<number, number> = {};
  const colStyles: Record<number, number> = {};
  for (const d of sheet.colDefs ?? []) {
    for (let c = d.min; c <= Math.min(d.max, MAX_COLS - 1); c++) {
      if (d.style !== undefined && d.style > 0) colStyles[c] = d.style;
      if (d.hidden) colWidths[c] = 0;
      else if (d.width !== undefined && d.width > 0) colWidths[c] = colPx(d.width);
    }
  }
  const rowHeights: Record<number, number> = {};
  const rowStyles: Record<number, number> = {};
  for (const [k, d] of Object.entries(sheet.rowDefs ?? {})) {
    const r = Number(k);
    if (d.hidden) rowHeights[r] = 0;
    else if (d.ht !== undefined && d.ht > 0) rowHeights[r] = d.custom ? Math.max(4, rowPx(d.ht)) : Math.max(MIN_ROW_HEIGHT, rowPx(d.ht));
    if (d.s !== undefined) rowStyles[r] = d.s;
  }
  sheet.colWidths = colWidths;
  sheet.colStyles = colStyles;
  sheet.rowHeights = rowHeights;
  sheet.rowStyles = rowStyles;
}

/** 重新计算数据范围（最后一个有值 / 合并的行列） */
export function updateExtent(sheet: SheetData): void {
  let rows = 0;
  let cols = 0;
  sheet.cells.forEach((row, r) => {
    if (!row) return;
    for (let c = row.length - 1; c >= 0; c--) {
      if (row[c]) {
        rows = r + 1;
        if (c + 1 > cols) cols = c + 1;
        break;
      }
    }
  });
  for (const m of sheet.merges) {
    rows = Math.max(rows, m.r1 + 1);
    cols = Math.max(cols, m.c1 + 1);
  }
  sheet.rows = rows;
  sheet.cols = cols;
}

/** 按数字格式显示数值 */
export function displayText(v: Raw, code: string, date1904 = false): string {
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  try {
    return formatNumber(v, code, date1904);
  } catch {
    return formatGeneral(v);
  }
}

/** 单元格的显示文字（值或样式改变后重算） */
export function refreshText(book: Workbook, cell: SCell): void {
  cell.t = cell.k ? String(cell.v) : displayText(cell.v, book.styleBook ? formatCode(book.styleBook, cell.s) : 'General', book.date1904);
  if (cell.k === 'd') cell.t = String(cell.v).replace('T', ' ');
}

/** styleBook 追加了格式后，补齐显示样式 */
export function syncStyles(book: Workbook): void {
  if (book.styleBook && book.styles.length !== book.styleBook.xfs.length) book.styles = allCellStyles(book.styleBook);
}

const EXTRA_COL = new Set(['min', 'max', 'width', 'hidden', 'style', 'customWidth', 'outlineLevel', 'collapsed']);
const EXTRA_ROW = new Set(['r', 'ht', 'customHeight', 'hidden', 's', 'customFormat', 'spans', 'outlineLevel', 'collapsed']);
const EXTRA_C = new Set(['r', 's', 't']);

/** 标签里的全部属性（保持原来的顺序，值是解码后的文字） */
export function attrsOf(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of tag.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = unescapeXml(m[2]);
  return out;
}

function extraAttrs(attrs: string, known: Set<string>): string | undefined {
  const out: string[] = [];
  for (const m of attrs.matchAll(/([\w:]+)="([^"]*)"/g)) if (!known.has(m[1])) out.push(`${m[1]}="${m[2]}"`);
  return out.length ? out.join(' ') : undefined;
}

const truthy = (v: string | undefined) => v === '1' || v === 'true';

function parseSheet(name: string, xml: string, shared: string[], book: StyleBook, date1904: boolean): SheetData {
  const cells: (SCell | undefined)[][] = [];
  const blankStyles: (number | undefined)[][] = [];
  const fmt = xml.match(/<sheetFormatPr\b[^>]*>/)?.[0] ?? '';
  const defCol = Number(xmlAttr(fmt, 'defaultColWidth'));
  const baseCol = Number(xmlAttr(fmt, 'baseColWidth'));
  const defaultColWidth = defCol > 0 ? colPx(defCol) : baseCol > 0 ? colPx(baseCol) + 5 : DEFAULT_COL_WIDTH;
  const defRow = Number(xmlAttr(fmt, 'defaultRowHeight'));
  const defaultRowHeight = Math.max(MIN_ROW_HEIGHT, defRow > 0 ? rowPx(defRow) : 0);

  const colDefs: ColDef[] = [];
  for (const m of (xml.match(/<cols>([\s\S]*?)<\/cols>/)?.[1] ?? '').matchAll(/<col\b([^>]*)>/g)) {
    const a = m[1];
    const w = Number(xmlAttr(a, 'width'));
    const style = Number(xmlAttr(a, 'style') ?? 0);
    const d: ColDef = { min: Number(xmlAttr(a, 'min')) - 1, max: Number(xmlAttr(a, 'max')) - 1 };
    if (!(d.min >= 0) || !(d.max >= d.min)) continue;
    if (w > 0) d.width = w;
    if (truthy(xmlAttr(a, 'hidden'))) d.hidden = true;
    if (style > 0) d.style = style;
    const lv = Number(xmlAttr(a, 'outlineLevel'));
    if (lv > 0) d.level = Math.min(7, Math.trunc(lv));
    if (truthy(xmlAttr(a, 'collapsed'))) d.collapsed = true;
    const extra = extraAttrs(a.replace(/\/$/, ''), EXTRA_COL);
    if (extra) d.extra = extra;
    colDefs.push(d);
  }

  let freeze = { rows: 0, cols: 0 };
  const pane = xml.match(/<pane\b[^>]*>/)?.[0];
  if (pane && /frozen/i.test(xmlAttr(pane, 'state') ?? '')) {
    freeze = { rows: Math.floor(Number(xmlAttr(pane, 'ySplit') ?? 0)), cols: Math.floor(Number(xmlAttr(pane, 'xSplit') ?? 0)) };
  }

  const rowDefs: Record<number, RowDef> = {};
  /** 共享公式：si → 主公式与它所在的格子 */
  const sharedF = new Map<string, { f: string; r: number; c: number }>();
  let rows = 0;
  let cols = 0;
  let nextRow = 0;
  for (const r of xml.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const ri = Number(xmlAttr(r[1], 'r') ?? nextRow + 1) - 1;
    nextRow = ri + 1;
    const rd: RowDef = {};
    const ht = Number(xmlAttr(r[1], 'ht'));
    if (ht > 0) rd.ht = ht;
    if (truthy(xmlAttr(r[1], 'customHeight'))) rd.custom = true;
    if (truthy(xmlAttr(r[1], 'hidden'))) rd.hidden = true;
    const rs = xmlAttr(r[1], 's');
    if (rs !== undefined && truthy(xmlAttr(r[1], 'customFormat'))) rd.s = Number(rs);
    const rlv = Number(xmlAttr(r[1], 'outlineLevel'));
    if (rlv > 0) rd.level = Math.min(7, Math.trunc(rlv));
    if (truthy(xmlAttr(r[1], 'collapsed'))) rd.collapsed = true;
    const extra = extraAttrs(r[1], EXTRA_ROW);
    if (extra) rd.extra = extra;
    if (Object.keys(rd).length) rowDefs[ri] = rd;
    let row: (SCell | undefined)[] | undefined;
    let nextCol = 0;
    for (const c of (r[2] ?? '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = xmlAttr(c[1], 'r');
      const ci = ref ? colIndex(ref) : nextCol;
      nextCol = ci + 1;
      const t = xmlAttr(c[1], 't');
      const sAttr = xmlAttr(c[1], 's');
      const s = sAttr === undefined ? undefined : Number(sAttr);
      const inner = c[2] ?? '';
      const v = inner.match(/<v>([\s\S]*?)<\/v>/)?.[1];
      let cell: SCell | undefined;
      if (t === 'inlineStr') {
        const text = xmlText(inner);
        if (text) cell = { v: text, t: text };
      } else if (v === undefined) cell = undefined;
      else if (t === 's') {
        const text = shared[Number(v)] ?? '';
        if (text) cell = { v: text, t: text, ss: Number(v) };
      } else if (t === 'str') {
        const text = unescapeXml(v);
        cell = { v: text, t: text };
      } else if (t === 'e') {
        const text = unescapeXml(v);
        cell = { v: text, t: text, k: 'e' };
      } else if (t === 'b') cell = { v: v === '1', t: v === '1' ? 'TRUE' : 'FALSE' };
      else if (t === 'd') {
        const text = unescapeXml(v).replace('T00:00:00', '').replace(/(\.\d+)?Z?$/, '');
        cell = { v: text, t: text.replace('T', ' '), k: 'd' };
      } else {
        const n = Number(v);
        if (Number.isFinite(n)) cell = { v: n, t: displayText(n, formatCode(book, s), date1904) };
        else {
          const text = unescapeXml(v);
          if (text) cell = { v: text, t: text };
        }
      }
      // 公式
      const fm = inner.match(/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/);
      if (fm) {
        const ft = xmlAttr(fm[1], 't');
        const text = fromFileFormula(unescapeXml(fm[2] ?? ''));
        let formula: string | undefined;
        if (ft === 'shared') {
          const si = xmlAttr(fm[1], 'si') ?? '';
          if (text) {
            sharedF.set(si, { f: text, r: ri, c: ci });
            formula = text;
          } else {
            const master = sharedF.get(si);
            if (master) formula = moveFormula(master.f, ri - master.r, ci - master.c);
          }
        } else if (ft === 'dataTable') {
          cell ??= { v: '', t: '' };
          cell.fraw = fm[0];
        } else formula = text;
        if (formula) {
          cell ??= { v: '', t: '' };
          cell.f = formula;
          if (ft === 'array') {
            cell.fa = xmlAttr(fm[1], 'ref') ?? ref;
            // 动态数组（Excel 365 写成带 cm 的数组公式）和编辑器自己保存的溢出公式：重新打开后继续按动态数组计算
            if (/\scm="/.test(c[1]) || DYNAMIC_TOP.test(formula)) cell.dyn = true;
          }
        }
      }
      if (cell && cell.v === '' && !cell.f && !cell.fraw) cell = undefined;
      if (!cell) {
        // 空格子也要记住样式：它可能有底色，且会盖过整行 / 整列格式
        (blankStyles[ri] ??= [])[ci] = s ?? 0;
        continue;
      }
      if (s !== undefined && s > 0) cell.s = s;
      const xa = extraAttrs(c[1], EXTRA_C);
      if (xa) cell.xa = xa;
      row ??= cells[ri] = [];
      row[ci] = cell;
      if (ci + 1 > cols) cols = ci + 1;
      if (ri + 1 > rows) rows = ri + 1;
    }
  }

  const merges: Range[] = [];
  for (const m of (xml.match(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/)?.[1] ?? '').matchAll(/<mergeCell\b[^>]*>/g)) {
    const range = parseRef(xmlAttr(m[0], 'ref') ?? '');
    if (!range || (range.r0 === range.r1 && range.c0 === range.c1)) continue;
    merges.push(range);
    rows = Math.max(rows, range.r1 + 1);
    cols = Math.max(cols, range.c1 + 1);
  }
  const sheet: SheetData = {
    name,
    cells,
    rows,
    cols,
    merges,
    colWidths: {},
    rowHeights: {},
    defaultColWidth,
    defaultRowHeight,
    freeze,
    blankStyles,
    colDefs,
    rowDefs,
  };
  const tab = xml.match(/<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/)?.[0].match(/<tabColor\b[^>]*>/)?.[0];
  const tabColor = resolveColor(tab, book.palette);
  if (tabColor) sheet.tabColor = tabColor;
  const sp = xml.match(/<sheetProtection\b[^>]*>/)?.[0];
  if (sp) {
    sheet.protection = { attrs: attrsOf(sp) };
    if (/\bsheet="(1|true)"/.test(sp)) sheet.protected = true;
  }
  const op = xml.match(/<outlinePr\b[^>]*>/)?.[0];
  if (op) sheet.outline = { summaryBelow: !/summaryBelow="(0|false)"/.test(op), summaryRight: !/summaryRight="(0|false)"/.test(op) };
  const af = parseAutoFilter(xml, book);
  if (af) sheet.autoFilter = af;
  const page = parsePage(xml);
  if (page) sheet.page = page;
  const cf = parseCf(xml, book);
  if (cf.length) sheet.cf = cf;
  const dv = parseValidations(xml);
  if (dv.length) sheet.validations = dv;
  deriveAxes(sheet);
  return sheet;
}

/** sqref（空格分隔的多块区域） → 区域列表 */
export function parseSqref(sqref: string | undefined): Range[] {
  return (sqref ?? '')
    .split(/\s+/)
    .map((x) => parseRef(x))
    .filter((g): g is Range => !!g);
}

function parseCf(xml: string, book: StyleBook): CfRule[] {
  const out: CfRule[] = [];
  let idx = 0;
  for (const m of xml.matchAll(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g)) {
    const ranges = parseSqref(xmlAttr(m[1], 'sqref'));
    if (!ranges.length) continue;
    for (const r of m[2].matchAll(/<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g)) {
      const a = r[1];
      const inner = r[2] ?? '';
      const rule: CfRule = {
        idx: idx++,
        ranges,
        type: xmlAttr(a, 'type') ?? '',
        priority: Number(xmlAttr(a, 'priority') ?? 0),
        formulas: [...inner.matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map((x) => fromFileFormula(unescapeXml(x[1]))),
      };
      const dxf = xmlAttr(a, 'dxfId');
      if (dxf !== undefined) rule.dxf = Number(dxf);
      if (truthy(xmlAttr(a, 'stopIfTrue'))) rule.stopIfTrue = true;
      const op = xmlAttr(a, 'operator');
      if (op) rule.operator = op;
      const text = xmlAttr(a, 'text');
      if (text !== undefined) rule.text = unescapeXml(text);
      const rank = xmlAttr(a, 'rank');
      if (rank !== undefined) rule.rank = Number(rank);
      if (truthy(xmlAttr(a, 'percent'))) rule.percent = true;
      if (truthy(xmlAttr(a, 'bottom'))) rule.bottom = true;
      if (rule.type === 'aboveAverage') rule.above = xmlAttr(a, 'aboveAverage') !== '0';
      const scale = inner.match(/<(colorScale|dataBar)\b[^>]*>([\s\S]*?)<\/\1>/);
      if (scale) {
        rule.cfvo = [...scale[2].matchAll(/<cfvo\b[^>]*>/g)].map((x) => ({ type: xmlAttr(x[0], 'type') ?? 'num', val: xmlAttr(x[0], 'val') }));
        rule.colors = [...scale[2].matchAll(/<color\b[^>]*>/g)].map((x) => resolveColor(x[0], book.palette) ?? '#638ec6');
      }
      out.push(rule);
    }
  }
  return out.sort((x, y) => x.priority - y.priority);
}

function parseValidations(xml: string): Validation[] {
  const out: Validation[] = [];
  for (const m of xml.matchAll(/<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g)) {
    const a = m[1];
    const inner = m[2] ?? '';
    const ranges = parseSqref(xmlAttr(a, 'sqref'));
    if (!ranges.length) continue;
    const f = (n: string) => {
      const v = inner.match(new RegExp(`<${n}>([\\s\\S]*?)</${n}>`))?.[1];
      return v === undefined ? undefined : fromFileFormula(unescapeXml(v));
    };
    const v: Validation = {
      ranges,
      type: xmlAttr(a, 'type') ?? 'any',
      allowBlank: truthy(xmlAttr(a, 'allowBlank')),
      showError: truthy(xmlAttr(a, 'showErrorMessage')),
    };
    const op = xmlAttr(a, 'operator');
    if (op) v.operator = op;
    const f1 = f('formula1');
    const f2 = f('formula2');
    if (f1 !== undefined) v.formula1 = f1;
    if (f2 !== undefined) v.formula2 = f2;
    for (const k of ['errorStyle', 'error', 'errorTitle', 'prompt', 'promptTitle'] as const) {
      const x = xmlAttr(a, k);
      if (x !== undefined) v[k] = unescapeXml(x);
    }
    if (truthy(xmlAttr(a, 'showDropDown'))) v.noDropDown = true;
    out.push(v);
  }
  return out;
}

const FILTER_OPS = new Set(['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual']);
type FilterOp = NonNullable<ColumnFilter['custom']>['rules'][number]['op'];

function parseAutoFilter(xml: string, styles?: StyleBook): AutoFilter | undefined {
  const m = xml.match(/<autoFilter\b([^>]*?)(?:\/>|>([\s\S]*?)<\/autoFilter>)/);
  const ref = m ? parseRef(xmlAttr(m[1], 'ref') ?? '') : null;
  if (!m || !ref) return undefined;
  const cols: Record<number, ColumnFilter> = {};
  for (const fc of (m[2] ?? '').matchAll(/<filterColumn\b([^>]*)>([\s\S]*?)<\/filterColumn>/g)) {
    const id = Number(xmlAttr(fc[1], 'colId'));
    const body = fc[2];
    const filters = body.match(/<filters\b([^>]*?)(?:\/>|>([\s\S]*?)<\/filters>)/);
    if (filters) {
      const values = [...(filters[2] ?? '').matchAll(/<filter\b[^>]*>/g)].map((x) => unescapeXml(xmlAttr(x[0], 'val') ?? ''));
      if (truthy(xmlAttr(filters[1], 'blank'))) values.push('');
      // 日期分组：dateTimeGrouping 到哪一级就取到哪一级（时分秒按天处理）
      const dates: DateGroup[] = [...(filters[2] ?? '').matchAll(/<dateGroupItem\b[^>]*>/g)].map((x) => {
        const g = xmlAttr(x[0], 'dateTimeGrouping') ?? 'year';
        const y = Number(xmlAttr(x[0], 'year'));
        if (g === 'year') return { y };
        if (g === 'month') return { y, m: Number(xmlAttr(x[0], 'month')) };
        return { y, m: Number(xmlAttr(x[0], 'month')), d: Number(xmlAttr(x[0], 'day')) };
      });
      cols[id] = dates.length ? { values, dates } : { values };
      continue;
    }
    const cf = body.match(/<colorFilter\b[^>]*>/)?.[0];
    if (cf) {
      const dxf = Number(xmlAttr(cf, 'dxfId'));
      const cell = xmlAttr(cf, 'cellColor') === undefined || truthy(xmlAttr(cf, 'cellColor'));
      const st = Number.isFinite(dxf) ? styles?.dxfs?.[dxf] : undefined;
      cols[id] = { color: { cell, color: (cell ? st?.fill : st?.color) ?? null, ...(Number.isFinite(dxf) && { dxf }) } };
      continue;
    }
    const custom = body.match(/<customFilters\b([^>]*?)>([\s\S]*?)<\/customFilters>/);
    if (custom) {
      const rules = [...custom[2].matchAll(/<customFilter\b[^>]*>/g)].map((x) => {
        const op = xmlAttr(x[0], 'operator') ?? 'equal';
        return { op: (FILTER_OPS.has(op) ? op : 'equal') as FilterOp, val: unescapeXml(xmlAttr(x[0], 'val') ?? '') };
      });
      cols[id] = { custom: { and: truthy(xmlAttr(custom[1], 'and')), rules } };
    }
  }
  return { ref, cols };
}

function parsePage(xml: string): PageSetup | undefined {
  const ps = xml.match(/<pageSetup\b[^>]*>/)?.[0];
  const po = xml.match(/<printOptions\b[^>]*>/)?.[0];
  const fit = truthy(xmlAttr(xml.match(/<pageSetUpPr\b[^>]*>/)?.[0] ?? '', 'fitToPage'));
  const pm = xml.match(/<pageMargins\b[^>]*>/)?.[0];
  const hf = xml.match(/<headerFooter\b[^>]*\/>|<headerFooter\b[^>]*>[\s\S]*?<\/headerFooter>/)?.[0];
  const breaks = (tag: string) =>
    [...(xml.match(new RegExp(`<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`))?.[0] ?? '').matchAll(/<brk\b[^>]*>/g)]
      .filter((m) => truthy(xmlAttr(m[0], 'man'))) // 没写 man 的是自动分页符
      .map((m) => Number(xmlAttr(m[0], 'id')))
      .filter((n) => n > 0);
  const rowBreaks = breaks('rowBreaks');
  const colBreaks = breaks('colBreaks');
  const hfText = (tag: string) => {
    const m = hf?.match(new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`));
    return m ? unescapeXml(m[1]) : undefined;
  };
  const header = hfText('oddHeader');
  const footer = hfText('oddFooter');
  if (!ps && !po && !pm && !header && !footer && !rowBreaks.length && !colBreaks.length) return undefined;
  const page: PageSetup = {};
  if (pm) {
    const v = (k: keyof PageMargins) => {
      const n = Number(xmlAttr(pm, k));
      return Number.isFinite(n) && xmlAttr(pm, k) !== undefined ? n : DEFAULT_MARGINS[k];
    };
    page.margins = { left: v('left'), right: v('right'), top: v('top'), bottom: v('bottom'), header: v('header'), footer: v('footer') };
  }
  if (header) page.header = header;
  if (footer) page.footer = footer;
  if (rowBreaks.length) page.rowBreaks = rowBreaks;
  if (colBreaks.length) page.colBreaks = colBreaks;
  if (po && truthy(xmlAttr(po, 'horizontalCentered'))) page.centerH = true;
  if (po && truthy(xmlAttr(po, 'verticalCentered'))) page.centerV = true;
  if (ps) {
    if (xmlAttr(ps, 'orientation') === 'landscape') page.landscape = true;
    const paper = Number(xmlAttr(ps, 'paperSize'));
    if (paper) page.paper = paper;
    const scale = Number(xmlAttr(ps, 'scale'));
    if (scale) page.scale = scale;
    if (fit) {
      page.fitW = Number(xmlAttr(ps, 'fitToWidth') ?? 1);
      page.fitH = Number(xmlAttr(ps, 'fitToHeight') ?? 1);
    }
  }
  if (po && truthy(xmlAttr(po, 'gridLines'))) page.gridLines = true;
  return page;
}

/** 编辑器不能编辑、但保存时原样保留的内容 */
function keepsOf(files: Record<string, Uint8Array>, sheetXml: string[]): string[] {
  const out = new Set<string>();
  const names = Object.keys(files);
  const all = sheetXml.join('');
  if (names.some((n) => n.startsWith('xl/charts/'))) out.add('图表');
  if (names.some((n) => n.startsWith('xl/media/'))) out.add('图片');
  if (names.some((n) => /^xl\/(comments|threadedComments)/.test(n))) out.add('批注');
  if (names.some((n) => n.startsWith('xl/pivotTables/'))) out.add('数据透视表');
  if (names.some((n) => n.startsWith('xl/tables/'))) out.add('套用表格格式的表格范围和样式');
  // 普通条件格式可以在编辑器里新建 / 删除；图标集和 x14 扩展（新式数据条等）只能原样保留
  if (/<iconSet\b|<x14:conditionalFormatting\b/.test(all)) out.add('条件格式（图标集等）');
  if (/<dataValidations?\b/.test(all)) out.add('数据验证');
  if (/<hyperlinks\b/.test(all)) out.add('超链接');
  if (/<sheetProtection\b/.test(all)) out.add('工作表保护');
  if (names.includes('xl/vbaProject.bin')) out.add('宏');
  return [...out];
}

export function readWorkbook(data: Uint8Array): Workbook {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(data);
  } catch {
    throw new Error('不是有效的 .xlsx 文件（旧版 .xls 请用 Excel 打开）');
  }
  const text = (p: string) => (files[p] ? strFromU8(files[p]) : '');
  const wb = text('xl/workbook.xml');
  if (!wb) throw new Error('不是有效的 .xlsx 文件');
  const rels = new Map<string, string>();
  for (const m of text('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b[^>]*>/g)) {
    const id = xmlAttr(m[0], 'Id');
    const target = xmlAttr(m[0], 'Target');
    if (id && target) rels.set(id, resolvePart('xl/workbook.xml', target));
  }
  const sharedPath = [...rels.values()].find((p) => /sharedStrings\.xml$/i.test(p)) ?? 'xl/sharedStrings.xml';
  const shared = [...text(sharedPath).matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g)].map((m) => xmlText(m[1] ?? ''));
  const themePath = [...text('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b[^>]*>/g)]
    .filter((m) => /\/theme$/.test(xmlAttr(m[0], 'Type') ?? ''))
    .map((m) => resolvePart('xl/workbook.xml', xmlAttr(m[0], 'Target') ?? ''))[0];
  const stylesPath = [...rels.values()].find((p) => /styles\.xml$/i.test(p)) ?? 'xl/styles.xml';
  const styleBook = readStyleBook(text(stylesPath), themePath ? text(themePath) : '');
  const pr = wb.match(/<workbookPr\b[^>]*>/)?.[0] ?? '';
  const date1904 = truthy(xmlAttr(pr, 'date1904'));

  const sheets: SheetData[] = [];
  const xmls: string[] = [];
  let origIndex = 0;
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const idx = origIndex++;
    const state = xmlAttr(m[0], 'state');
    const path = rels.get(xmlAttr(m[0], 'r:id') ?? '');
    if (!path || !files[path]) continue;
    const xml = text(path);
    xmls.push(xml);
    const sheet = parseSheet(unescapeXml(xmlAttr(m[0], 'name') ?? ''), xml, shared, styleBook, date1904);
    sheet.path = path;
    const extra = readDrawings(files, path);
    if (extra.drawings.length) {
      sheet.drawings = extra.drawings;
      sheet.hadDrawings = true;
    }
    if (extra.notes.length) sheet.notes = extra.notes;
    if (extra.links.length) sheet.links = extra.links;
    if (extra.tables.length) sheet.tables = extra.tables;
    // 表格自带的筛选：像工作表的自动筛选一样显示和使用（区域按数据范围截短，免得整列 100 万行都去判断）；只取第一个
    const tf = extra.tables.find((t) => t.filter && t.headerRows > 0);
    if (tf && !sheet.autoFilter) {
      const af = parseAutoFilter(text(tf.path), styleBook);
      if (af) {
        sheet.tableFilter = { path: tf.path, ref: xmlAttr(text(tf.path).match(/<autoFilter\b[^>]*>/)?.[0] ?? '', 'ref') ?? '' };
        const last = Math.max(af.ref.r0, Math.min(af.ref.r1, sheet.rows - 1));
        sheet.autoFilter = { ...af, ref: { ...af.ref, r1: last }, table: true };
      }
    }
    sheet.rid = xmlAttr(m[0], 'r:id');
    sheet.sheetId = Number(xmlAttr(m[0], 'sheetId'));
    sheet.origIndex = idx;
    sheet.fileName = sheet.name;
    if (state === 'hidden' || state === 'veryHidden') sheet.hidden = state;
    sheets.push(sheet);
  }
  if (!sheets.some((s) => !s.hidden)) throw new Error('工作簿里没有可显示的工作表');
  const names: NonNullable<Workbook['names']> = [];
  for (const m of wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const name = unescapeXml(xmlAttr(m[1], 'name') ?? '');
    const local = xmlAttr(m[1], 'localSheetId');
    if (name === '_xlnm.Print_Titles' || name === '_xlnm.Print_Area') {
      const sheet = sheets.find((x) => x.origIndex === Number(local));
      if (sheet) applyPrintName(sheet, name, unescapeXml(m[2]));
      continue;
    }
    if (!name || name.startsWith('_xlnm.')) continue;
    names.push({ name, f: fromFileFormula(unescapeXml(m[2])), ...(local !== undefined && { local: Number(local) }) });
  }
  const wp = wb.match(/<workbookProtection\b[^>]*>/)?.[0];
  return { sheets, styles: allCellStyles(styleBook), styleBook, date1904, kind: 'xlsx', keeps: keepsOf(files, xmls), names, ...(wp && { wbProtection: { attrs: attrsOf(wp) } }) };
}

/** 打印标题（Sheet1!$1:$2）、打印区域（Sheet1!$A$1:$H$30）→ 页面设置；多段的只取第一段 */
function applyPrintName(sheet: SheetData, name: string, f: string) {
  const parts = f.split(',').map((x) => x.replace(/^.*!/, '').replace(/\$/g, ''));
  if (name === '_xlnm.Print_Titles') {
    const rows = parts.find((x) => /^\d+:\d+$/.test(x));
    if (!rows) return;
    const [a, b] = rows.split(':').map((x) => Number(x) - 1);
    (sheet.page ??= {}).titleRows = [Math.min(a, b), Math.max(a, b)];
  } else {
    const g = parseRef(parts[0]);
    if (g) (sheet.page ??= {}).area = g;
  }
}

/** 猜 CSV 分隔符：第一行（引号外）出现最多的 , Tab ; */
export function guessDelimiter(text: string): string {
  const counts: Record<string, number> = { ',': 0, '\t': 0, ';': 0 };
  let quoted = false;
  for (const ch of text.slice(0, 4096)) {
    if (ch === '"') quoted = !quoted;
    else if (!quoted && (ch === '\n' || ch === '\r')) break;
    else if (!quoted && ch in counts) counts[ch]++;
  }
  const [best, n] = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
  return n > 0 ? best : ',';
}

const NUM_RE = /^[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/;

export function readCsv(text: string, name = 'Sheet1'): Workbook {
  const src = text.replace(/^﻿/, '');
  const delim = guessDelimiter(src);
  const cells: (SCell | undefined)[][] = [];
  let row: (SCell | undefined)[] = [];
  let field = '';
  let quoted = false;
  let wasQuoted = false;
  let rows = 0;
  let cols = 0;
  const pushField = () => {
    const ci = row.length;
    if (field !== '') {
      // 前导 0（编号）和超长数字（身份证号等）保持文字
      const numeric = !wasQuoted && NUM_RE.test(field) && !/^[-+]?0\d/.test(field) && field.replace(/\D/g, '').length <= 15;
      row.push({ v: numeric ? Number(field) : field, t: field });
      cols = Math.max(cols, ci + 1);
      rows = cells.length + 1;
    } else row.push(undefined);
    field = '';
    wasQuoted = false;
  };
  const pushRow = () => {
    pushField();
    cells.push(row);
    row = [];
  };
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          field += '"';
          i++;
        } else quoted = false;
      } else field += ch;
    } else if (ch === '"' && field === '') {
      quoted = true;
      wasQuoted = true;
    } else if (ch === delim) pushField();
    else if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && src[i + 1] === '\n') i++;
      pushRow();
    } else field += ch;
  }
  if (field !== '' || row.length) pushRow();
  cells.length = rows;
  const styleBook = readStyleBook('');
  return {
    sheets: [
      {
        name,
        cells,
        rows,
        cols,
        merges: [],
        colWidths: {},
        rowHeights: {},
        defaultColWidth: 96,
        defaultRowHeight: MIN_ROW_HEIGHT + 2,
        freeze: { rows: 0, cols: 0 },
      },
    ],
    styles: allCellStyles(styleBook),
    styleBook,
    kind: 'csv',
    csv: { delim, encoding: 'UTF-8', bom: false },
  };
}

/** 新建的空白工作簿（一个 Sheet1） */
export function blankWorkbook(): Workbook {
  const styleBook = readStyleBook('');
  return {
    sheets: [
      {
        name: 'Sheet1',
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
      },
    ],
    styles: allCellStyles(styleBook),
    styleBook,
    kind: 'new',
  };
}

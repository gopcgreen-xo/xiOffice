/**
 * 数据透视表（第一阶段）：把一片带标题行的数据按“行字段 / 列字段 / 值字段”汇总，结果写成普通单元格放进新工作表。
 *
 * - 汇总在 `buildPivot`（纯函数）：按字段值分组，求和 / 计数 / 平均 / 最大 / 最小；总计（行、列）单独累计，平均值、最大最小都是对原始记录算的。
 * - 工作表上保存透视定义（`SheetData.pivot`，只在编辑期间有效，不写进文件），改了源数据后可以“刷新”。
 * - 写进文件的是普通表格（值 + 样式），不是 Excel 的 pivotTable 部件，所以在 Excel 里不能拖拽字段；
 *   文件里原有的透视表部件照旧原样保留。
 */
import { compare, type Scalar } from './formula/eval';
import { applyStyle, setColWidth, setParsed, type Tx } from './ops';
import { addSheet } from './sheets';
import type { Range, SheetData, Workbook } from './workbook';

export type Agg = 'sum' | 'count' | 'counta' | 'avg' | 'max' | 'min';
// 名称同 Excel 中文版的默认数据字段名（求和项 / 计数项 / 数值计数项 / 平均值项 / 最大值项 / 最小值项）
export const AGG_LABEL: Record<Agg, string> = { sum: '求和', count: '数值计数', counta: '计数', avg: '平均值', max: '最大值', min: '最小值' };
export const BLANK_LABEL = '(空白)';

export interface PivotValue {
  field: number;
  agg: Agg;
}

export interface PivotConfig {
  /** 行字段（列号，相对数据区域第一列），顺序 = 嵌套顺序 */
  rows: number[];
  cols: number[];
  values: PivotValue[];
  /** 是否显示总计行 / 总计列 */
  totals: { rows: boolean; cols: boolean };
}

/** 工作表上记着的透视定义：数据在哪、怎么汇总、上次写到哪 */
/** 保存过的 Excel 透视表部件的位置（再保存时原地覆盖） */
export interface PivotParts {
  table: string;
  cacheDef: string;
  cacheRec: string;
  cacheId: number;
}

export interface PivotDef {
  source: { sheet: string; range: Range };
  /** 写进文件后记下部件位置；没有 = 还没写过 */
  parts?: PivotParts;
  config: PivotConfig;
  /** 上次写出的区域（刷新时先清掉） */
  extent: Range;
}

export interface PivotResult {
  cells: Scalar[][];
  /** 标题行数 */
  headerRows: number;
  /** 左边的行标签列数 */
  labelCols: number;
  /** 总计行 / 总计列在 cells 里的位置（没有 = -1） */
  totalRow: number;
  totalCol: number;
}

export interface Acc {
  sum: number;
  n: number;
  nonBlank: number;
  max: number;
  min: number;
}
export const newAcc = (): Acc => ({ sum: 0, n: 0, nonBlank: 0, max: -Infinity, min: Infinity });
const isBlank = (v: Scalar) => v === null || v === '';

function add(a: Acc, v: Scalar) {
  if (isBlank(v)) return;
  a.nonBlank++;
  if (typeof v === 'number') {
    a.n++;
    a.sum += v;
    if (v > a.max) a.max = v;
    if (v < a.min) a.min = v;
  }
}

export function result(a: Acc | undefined, agg: Agg): Scalar {
  if (!a) return '';
  const fin = (x: number) => Number(x.toPrecision(15));
  switch (agg) {
    case 'sum':
      return a.n ? fin(a.sum) : 0;
    case 'count':
      return a.n;
    case 'counta':
      return a.nonBlank;
    case 'avg':
      return a.n ? fin(a.sum / a.n) : '';
    case 'max':
      return a.n ? a.max : '';
    case 'min':
      return a.n ? a.min : '';
  }
}

/** 分组键里的标签：空白显示成 (空白)，数字保持数字（排序时数字在文字前） */
export const labelOf = (v: Scalar): string | number | boolean => (isBlank(v) || (typeof v === 'object' && v !== null) ? BLANK_LABEL : (v as string | number | boolean));

/** 排序：数字 < 文字 < 逻辑值，(空白) 最后；文字按拼音顺序 */
export function cmpLabel(a: string | number | boolean, b: string | number | boolean): number {
  if (a === BLANK_LABEL) return b === BLANK_LABEL ? 0 : 1;
  if (b === BLANK_LABEL) return -1;
  if (typeof a === 'string' && typeof b === 'string') return a.localeCompare(b, 'zh-Hans-CN');
  return compare(a, b);
}

export type Label = string | number | boolean;

/** 汇总前的分析：记录、字段名、排好序的行 / 列分组、各层级的累计（给布局和 Excel 部件共用） */
export interface PivotAnalysis {
  fieldNames: string[];
  /** 非空记录（全空的行已去掉） */
  records: Scalar[][];
  /** 有效的值字段（没选值字段时按第一个字段做“计数”，Excel 的透视表至少要有一个值字段） */
  values: PivotValue[];
  rowKeys: Label[][];
  colKeys: Label[][];
  grid: Map<string, Acc[]>;
  rowTot: Map<string, Acc[]>;
  colTot: Map<string, Acc[]>;
  grand: Acc[];
}

export const keyId = (key: Label[]) => key.map((k) => `${typeof k}:${k}`).join('\u0001');

export const effectiveValues = (cfg: PivotConfig, fieldCount = 0): PivotValue[] => {
  if (cfg.values.length) return cfg.values;
  // Excel 不允许同一个字段既在行 / 列上又做值字段：没选值字段时，数第一个没被用上的字段
  const used = new Set([...cfg.rows, ...cfg.cols]);
  let f = 0;
  while (f < fieldCount - 1 && used.has(f)) f++;
  return [{ field: f, agg: 'counta' }];
};

export function analyzePivot(data: Scalar[][], cfg: PivotConfig): PivotAnalysis {
  const head = data[0] ?? [];
  const records = data.slice(1).filter((row) => row.some((v) => !isBlank(v)));
  const width = Math.max(head.length, ...data.map((r) => r.length));
  const fieldNames = Array.from({ length: width }, (_, f) => (isBlank(head[f]) ? `列${f + 1}` : String(head[f])));
  const values = effectiveValues(cfg, width);
  const nv = values.length;
  const fresh = () => Array.from({ length: nv }, newAcc);
  const grid = new Map<string, Acc[]>();
  const rowTot = new Map<string, Acc[]>();
  const colTot = new Map<string, Acc[]>();
  const grand = fresh();
  const rows = new Map<string, Label[]>();
  const cols = new Map<string, Label[]>();
  const slot = (m: Map<string, Acc[]>, k: string): Acc[] => m.get(k) ?? m.set(k, fresh()).get(k)!;
  for (const row of records) {
    const rk = cfg.rows.map((f) => labelOf(row[f] ?? null));
    const ck = cfg.cols.map((f) => labelOf(row[f] ?? null));
    const rid = keyId(rk);
    const cid = keyId(ck);
    rows.set(rid, rk);
    cols.set(cid, ck);
    const cells = [slot(grid, `${rid}\u0002${cid}`), slot(rowTot, rid), slot(colTot, cid), grand];
    values.forEach((v, i) => {
      const x: Scalar = row[v.field] ?? null;
      for (const c of cells) add(c[i], x);
    });
  }
  const sorted = (m: Map<string, Label[]>): Label[][] =>
    [...m.values()].sort((a, b) => {
      for (let i = 0; i < a.length; i++) {
        const d = cmpLabel(a[i], b[i]);
        if (d) return d;
      }
      return 0;
    });
  // 没有行（列）字段时只有一个空分组：整张表汇总成一行（列）
  const rowKeys = cfg.rows.length ? sorted(rows) : [[]];
  const colKeys = cfg.cols.length ? sorted(cols) : [[]];
  return { fieldNames, records, values, rowKeys, colKeys, grid, rowTot, colTot, grand };
}

/** 值字段的标题：求和项:金额（同 Excel 的默认名） */
export const valueName = (v: PivotValue, names: string[]) => `${AGG_LABEL[v.agg]}项:${names[v.field] ?? `列${v.field + 1}`}`;

/**
 * 汇总并排成 Excel “表格形式”（无分类汇总）的样子：和 Excel 里同样设置的透视表逐格一致（验收时用 Excel 对拍）。
 * 行字段为空时只有一行汇总；列字段最多按第一个排版（多个时标签用 “ / ” 连起来）。
 */
export function buildPivot(data: Scalar[][], cfg: PivotConfig): PivotResult {
  const an = analyzePivot(data, cfg);
  const { fieldNames: names, values, rowKeys, colKeys } = an;
  const R = cfg.rows.length;
  const K = values.length;
  const hasCols = cfg.cols.length > 0;
  const showTotCol = hasCols && cfg.totals.cols;
  const showTotRow = R > 0 && cfg.totals.rows;
  const blank = (n: number): Scalar[] => Array(Math.max(0, n)).fill('');
  const capRows: Scalar[] = cfg.rows.map((f) => names[f]);
  const colCap = cfg.cols.map((f) => names[f]).join(' / ');
  const vn = values.map((v) => valueName(v, names));
  const blockLabel = (k: Label[]) => k.join(' / ');
  const empty = !an.records.length;

  // ---- 标题行（四种形态，对照 Excel）----
  let header: Scalar[][];
  if (!hasCols) {
    header = K === 1 ? [[...capRows, vn[0]]] : [[...blank(R), '值', ...blank(K - 1)], [...capRows, ...vn]];
  } else if (K === 1) {
    // 没有行字段、只有列字段和一个值字段时，Excel 在最左边多出一列，值字段名写在数据行的第一格
    const lead = R === 0 ? 1 : 0;
    const top: Scalar[] = R === 0 ? ['', colCap] : [vn[0], ...blank(R - 1), colCap];
    header = [top, [...blank(lead), ...capRows, ...(empty ? [] : colKeys.map(blockLabel)), ...(showTotCol ? ['总计'] : [])]];
  } else {
    const row1: Scalar[] = [...blank(R)];
    const row2: Scalar[] = [...capRows];
    if (!empty) {
      for (const k of colKeys) {
        row1.push(blockLabel(k), ...blank(K - 1));
        row2.push(...vn);
      }
    }
    if (showTotCol) {
      row1.push(...vn.map((n) => `${n}汇总`));
      row2.push(...blank(K));
    }
    header = [[...blank(R), colCap, '值'], row1, row2];
  }

  // ---- 数据 ----
  const out = (accs: Acc[] | undefined): Scalar[] => values.map((v, i) => result(accs?.[i], v.agg));
  const headerRows = header.length;
  const cells: Scalar[][] = [...header];
  let prev: Label[] = [];
  const usedRows = empty && R > 0 ? [] : rowKeys;
  for (const rk of usedRows) {
    const rid = keyId(rk);
    // 外层标签和上一行相同时留空（Excel 表格形式只在每组第一行写外层标签）
    let same = true;
    const labels: Scalar[] = rk.map((l, j) => {
      same = same && prev[j] === l;
      return same ? '' : l;
    });
    prev = rk;
    const line: Scalar[] = [...labels];
    if (hasCols) {
      for (const ck of colKeys) line.push(...out(an.grid.get(`${rid}\u0002${keyId(ck)}`)));
      if (showTotCol) line.push(...out(an.rowTot.get(rid)));
    } else line.push(...out(an.grid.get(`${rid}\u0002`)));
    if (R === 0 && hasCols && K === 1) line.unshift(vn[0]);
    cells.push(line);
  }
  let totalRow = -1;
  if (showTotRow && usedRows.length) {
    const line: Scalar[] = ['总计', ...blank(R - 1)];
    if (hasCols) {
      for (const ck of colKeys) line.push(...out(an.colTot.get(keyId(ck))));
      if (showTotCol) line.push(...out(an.grand));
    } else line.push(...out(an.grand));
    totalRow = cells.length;
    cells.push(line);
  }
  const width = Math.max(...cells.map((r) => r.length));
  for (const r of cells) while (r.length < width) r.push('');
  return { cells, headerRows, labelCols: R === 0 && hasCols && K === 1 ? 1 : R, totalRow, totalCol: showTotCol ? width - K : -1 };
}

// ---------- 写进工作表 ----------

/** 读源数据：值（公式取计算结果），错误值当空白 */
export function readSource(sheet: SheetData, g: Range): Scalar[][] {
  const out: Scalar[][] = [];
  for (let r = g.r0; r <= g.r1; r++) {
    const row: Scalar[] = [];
    for (let c = g.c0; c <= g.c1; c++) {
      const x = sheet.cells[r]?.[c];
      // 逻辑值按文字处理（Excel 的缓存里是 TRUE / FALSE 项）
      row.push(!x || x.k === 'e' ? null : typeof x.v === 'boolean' ? (x.v ? 'TRUE' : 'FALSE') : x.v);
    }
    out.push(row);
  }
  return out;
}

/** 数据区域第一行的字段名（空的用“列 N”） */
export function fieldNames(sheet: SheetData, g: Range): string[] {
  return Array.from({ length: g.c1 - g.c0 + 1 }, (_, i) => {
    const x = sheet.cells[g.r0]?.[g.c0 + i];
    return x && x.v !== '' ? String(x.t || x.v) : `列${i + 1}`;
  });
}

/** 默认设置：第一个文字列做行字段，最后一个数字列求和 */
export function defaultConfig(sheet: SheetData, g: Range): PivotConfig {
  const n = g.c1 - g.c0 + 1;
  const sample = (c: number) => sheet.cells[g.r0 + 1]?.[g.c0 + c]?.v;
  const numeric = Array.from({ length: n }, (_, c) => c).filter((c) => typeof sample(c) === 'number');
  const text = Array.from({ length: n }, (_, c) => c).filter((c) => typeof sample(c) === 'string');
  return {
    rows: text.length ? [text[0]] : [],
    cols: [],
    values: numeric.length ? [{ field: numeric[numeric.length - 1], agg: 'sum' }] : [],
    totals: { rows: true, cols: true },
  };
}

/** 把汇总结果写到工作表 si 的左上角（先清掉上次写的区域）；返回写出的区域 */
function writeResult(tx: Tx, si: number, res: PivotResult, prev?: Range): Range {
  tx.snap(si);
  const sheet = tx.book.sheets[si];
  if (prev) {
    for (let r = prev.r0; r <= prev.r1; r++) {
      if (sheet.cells[r]) sheet.cells[r].fill(undefined, prev.c0, prev.c1 + 1);
      if (sheet.blankStyles?.[r]) sheet.blankStyles[r].fill(undefined, prev.c0, prev.c1 + 1);
    }
  }
  res.cells.forEach((row, r) => row.forEach((v, c) => v !== '' && setParsed(tx, si, r, c, { kind: 'value', v: v as string | number | boolean })));
  const width = res.cells[0]?.length ?? 1;
  const extent: Range = { r0: 0, c0: 0, r1: Math.max(0, res.cells.length - 1), c1: width - 1 };
  // 只给有内容的格子加样式：不写“只有样式的空白格”（Excel 对透视表区域里的空白带样式格子很挑剔）
  const style = (g: Range, ch: Parameters<typeof applyStyle>[3]) => {
    for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) if (res.cells[r]?.[c] !== '' && res.cells[r]?.[c] !== undefined) applyStyle(tx, si, { r0: r, c0: c, r1: r, c1: c }, ch);
  };
  style({ r0: 0, c0: 0, r1: res.headerRows - 1, c1: width - 1 }, { bold: true, fill: '#EEF2F7' });
  if (res.totalRow >= 0) style({ r0: res.totalRow, c0: 0, r1: res.totalRow, c1: width - 1 }, { bold: true });
  if (res.totalCol >= 0) style({ r0: res.headerRows, c0: res.totalCol, r1: extent.r1, c1: width - 1 }, { bold: true });
  if (res.labelCols) setColWidth(tx, si, 0, res.labelCols - 1, { width: 120 });
  if (width > res.labelCols) setColWidth(tx, si, res.labelCols, width - 1, { width: 110 });
  return extent;
}

/** 新建工作表放透视结果。返回新工作表的序号 */
export function createPivot(tx: Tx, srcSi: number, range: Range, cfg: PivotConfig): number {
  const book = tx.book;
  const src = book.sheets[srcSi];
  const res = buildPivot(readSource(src, range), cfg);
  const taken = new Set(book.sheets.map((s) => s.name.toLowerCase()));
  let name = '透视表';
  for (let i = 2; taken.has(name.toLowerCase()); i++) name = `透视表 ${i}`;
  const at = addSheet(tx, srcSi + 1, name);
  const extent = writeResult(tx, at, res);
  book.sheets[at].pivot = { source: { sheet: src.name, range }, config: cfg, extent };
  return at;
}

/** 按当前源数据重新汇总。返回 false = 找不到源工作表 */
export function refreshPivot(tx: Tx, si: number): boolean {
  const book: Workbook = tx.book;
  const def = book.sheets[si].pivot;
  if (!def) return false;
  const src = book.sheets.find((s) => s.name === def.source.sheet);
  if (!src) return false;
  const res = buildPivot(readSource(src, def.source.range), def.config);
  tx.snap(si);
  const extent = writeResult(tx, si, res, def.extent);
  book.sheets[si].pivot = { ...def, extent };
  return true;
}

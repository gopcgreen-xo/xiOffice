/**
 * 列属性的纯逻辑（不改模型）：取某格的列属性、选项、插删行列时的位移、打开时按表头重新绑定、
 * 按现有内容推断列类型、必填检查。存放在工作包里，见 docs/WBP_FORMAT.md；改模型的操作在 columns.ts。
 */
import { checkColumnValue, columnAt, splitMulti, isCheckOn, type ColumnMeta, type ColumnType, type SelectOption, type SheetColumns } from '@wz/editor-shared';
import { colName } from '@wz/editor-shared/xlsx';
import { isDateFormat } from './numfmt';
import { formatCode } from './styles';
import { parseRef, type SheetData, type Workbook } from './workbook';

/** 该格所在列的列属性（表头行及以上不算，只有数据行才有） */
export function metaAt(sheet: SheetData, r: number, c: number): ColumnMeta | undefined {
  const sc = sheet.columns;
  if (!sc || r < sc.header) return undefined;
  return columnAt(sc, c);
}

/** 显示的文字（没有格子为空串） */
export const cellText = (sheet: SheetData, r: number, c: number): string => sheet.cells[r]?.[c]?.t ?? '';

/** 区域引用（`A2:A20` 或 `Sheet2!A2:A20`）里的文字，去空去重，保持出现顺序 */
function refValues(book: Workbook, sheet: SheetData, ref: string): string[] {
  // 引用包里别的表格：由宿主提供读取
  if (ref.trim().startsWith('[')) return book.externalValues?.(ref) ?? [];
  const bang = ref.lastIndexOf('!');
  const sheetName = bang > 0 ? ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") : sheet.name;
  const g = parseRef(bang > 0 ? ref.slice(bang + 1) : ref);
  const target = book.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
  if (!g || !target) return [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (let r = g.r0; r <= Math.min(g.r1, g.r0 + 2000); r++) {
    for (let c = g.c0; c <= g.c1; c++) {
      const t = cellText(target, r, c).trim();
      if (t && !seen.has(t)) {
        seen.add(t);
        out.push(t);
      }
    }
  }
  return out;
}

/** 单选 / 多选 / 复选框的可选值：引用区域优先，否则是手填的选项 */
export function optionValues(book: Workbook, sheet: SheetData, c: ColumnMeta): string[] {
  if (c.type === 'check') return ['是', '否'];
  if (c.optionsRef) return refValues(book, sheet, c.optionsRef);
  return (c.options ?? []).map((o) => o.v);
}

/** 带颜色的选项（引用区域时，颜色按手填选项里同名的取） */
export function optionList(book: Workbook, sheet: SheetData, c: ColumnMeta): SelectOption[] {
  const known = new Map((c.options ?? []).map((o) => [o.v, o]));
  return optionValues(book, sheet, c).map((v) => known.get(v) ?? { v });
}

/** 校验一格的文字（选项来自引用区域时按区域内容检查）；null = 通过，空文字也通过（必填另算） */
export function checkCell(book: Workbook, sheet: SheetData, c: ColumnMeta, text: string): string | null {
  if (!text.trim()) return null;
  const withOptions = c.optionsRef && (c.type === 'select' || c.type === 'multi') ? { ...c, options: optionValues(book, sheet, c).map((v) => ({ v })) } : c;
  return checkColumnValue(withOptions, text);
}

// ---------- 插入 / 删除行列时的位移 ----------

/** 插入（n > 0）/ 删除（n < 0）行列后的列属性：列号跟着移动，被删掉的列去掉，表头行跟着移动 */
export function shiftColumnMeta(sc: SheetColumns, axis: 'row' | 'col', at: number, n: number): SheetColumns {
  if (axis === 'col') {
    const columns: ColumnMeta[] = [];
    for (const c of sc.columns) {
      if (n < 0 && c.col >= at && c.col < at - n) continue;
      columns.push(c.col >= at ? { ...c, col: c.col + n } : c);
    }
    return { ...sc, columns };
  }
  const h0 = sc.header - 1;
  let h = h0;
  if (n > 0) {
    if (at <= h0) h = h0 + n;
  } else if (h0 >= at - n) h = h0 + n;
  else if (h0 >= at) h = at;
  return h === h0 ? sc : { ...sc, header: h + 1 };
}

// ---------- 打开时重新绑定 ----------

/**
 * 列属性按“列号 + 表头文字”绑定到表格的列。xlsx 被别的程序改过（插了列、改了表头）时：
 * 表头文字对不上的列，按表头文字在表头行里找，唯一一列就更新列号；找不到或不唯一，标为“未绑定”，
 * 不套用到任何列（不静默套到别的列上）。
 */
export function bindColumns(sheet: SheetData, sc: SheetColumns): { sc: SheetColumns; moved: string[]; unbound: string[] } {
  const hr = sc.header - 1;
  const textAt = (c: number) => cellText(sheet, hr, c).trim();
  const width = Math.max(sheet.cols, ...sc.columns.map((c) => c.col + 1));
  const matches = (c: ColumnMeta) => !c.title || textAt(c.col) === c.title.trim();
  const moved: string[] = [];
  const unbound: string[] = [];
  const next: ColumnMeta[] = sc.columns.map((c) => ({ ...c }));
  // 表头文字没变的列原地不动；其余的按表头文字重新找
  const taken = new Set(next.filter(matches).map((c) => c.col));
  for (const c of next) {
    if (matches(c)) continue;
    const hits: number[] = [];
    for (let k = 0; k < width; k++) if (textAt(k) === c.title.trim() && !taken.has(k)) hits.push(k);
    if (hits.length === 1) {
      moved.push(c.title);
      c.col = hits[0];
      taken.add(c.col);
    } else {
      c.unbound = true;
      unbound.push(c.title);
    }
  }
  return { sc: { ...sc, columns: next }, moved, unbound };
}

// ---------- 按内容推断列类型 ----------

export interface Inference {
  col: number;
  head: string;
  type: ColumnType;
  options?: string[];
  integer?: boolean;
  fmt?: string;
  /** 推断依据（给用户看） */
  reason: string;
  /** 补充说明（格式、范围等） */
  detail: string;
  warn?: string;
}

const MONEY_HEAD = /价|金额|费|元|款|额|税|总计|合计|price|amount|cost|fee/i;
const MAX_OPTIONS_INFER = 12;

/** 按表头下面各格的内容推断每一列的类型。header 是表头所在行（1 起）；没有表头文字的列跳过 */
export function inferColumns(book: Workbook, sheet: SheetData, header: number): Inference[] {
  const out: Inference[] = [];
  const hr = header - 1;
  const last = sheet.rows - 1;
  for (let c = 0; c < sheet.cols; c++) {
    const head = cellText(sheet, hr, c).trim();
    if (!head) continue;
    const cells: { r: number; t: string; v: unknown; fmt: string }[] = [];
    for (let r = header; r <= last; r++) {
      const cell = sheet.cells[r]?.[c];
      const t = (cell?.t ?? '').trim();
      if (t) cells.push({ r, t, v: cell?.v, fmt: book.styleBook ? formatCode(book.styleBook, cell?.s) : 'General' });
    }
    const total = last - hr;
    const base = { col: c, head };
    if (!cells.length) {
      out.push({ ...base, type: 'text', reason: total > 0 ? `${total} 行都是空的` : '没有数据', detail: '不设置' });
      continue;
    }
    const empty = total - cells.length;
    const distinct = [...new Set(cells.map((x) => x.t))];

    // 复选框：只有「是 / 否」
    if (distinct.every((t) => /^(是|否|true|false)$/i.test(t))) {
      out.push({ ...base, type: 'check', reason: '只有「是 / 否」', detail: '是 = 勾选，否 = 空框' });
      continue;
    }
    // 日期：多数格子是日期（日期格式的数字，或能解析的日期文字）
    const dateLike = (x: (typeof cells)[number]) => (typeof x.v === 'number' && isDateFormat(x.fmt)) || /^\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?$/.test(x.t);
    const dates = cells.filter(dateLike);
    if (dates.length / cells.length >= 0.8) {
      const bad = cells.find((x) => !dateLike(x));
      out.push({
        ...base,
        type: 'date',
        fmt: 'yyyy-mm-dd',
        reason: dates.length === cells.length ? '全部是日期' : `${dates.length} / ${cells.length} 为日期`,
        detail: 'yyyy-mm-dd',
        ...(bad && { warn: `${colName(c)}${bad.r + 1}「${bad.t}」不是日期，应用后标为不合规` }),
      });
      continue;
    }
    // 数字 / 金额 / 百分比
    if (cells.every((x) => typeof x.v === 'number')) {
      const pct = cells.every((x) => x.fmt.includes('%'));
      const ints = cells.every((x) => Number.isInteger(x.v as number));
      if (pct) out.push({ ...base, type: 'percent', fmt: '0.00%', reason: '全部是百分数', detail: '0.00%' });
      else if (MONEY_HEAD.test(head) && !ints) out.push({ ...base, type: 'money', fmt: '#,##0.00', reason: '全部是数字，表头像金额', detail: '#,##0.00' });
      else if (ints) out.push({ ...base, type: 'number', integer: true, reason: '全部为整数', detail: '只允许整数' });
      else out.push({ ...base, type: 'number', reason: '全部是数字', detail: '不限小数位' });
      continue;
    }
    // 多选：不少格子带分隔符，拆开后的取值很少
    const sep = cells.filter((x) => /[,，、;；]/.test(x.t));
    if (sep.length / cells.length >= 0.3) {
      const tokens = [...new Set(cells.flatMap((x) => splitMulti(x.t)))];
      if (tokens.length <= MAX_OPTIONS_INFER && tokens.length >= 2) {
        out.push({ ...base, type: 'multi', options: tokens, reason: `${sep.length} 格含多个取值，共 ${tokens.length} 种`, detail: '' });
        continue;
      }
    }
    // 单选：取值种类少且有重复
    if (distinct.length <= MAX_OPTIONS_INFER && distinct.length <= Math.max(1, Math.floor(cells.length * 0.6)) && distinct.length >= 2) {
      out.push({ ...base, type: 'select', options: distinct, reason: `${distinct.length} 种取值（≤ ${MAX_OPTIONS_INFER}）`, detail: '' });
      continue;
    }
    out.push({
      ...base,
      type: 'text',
      reason: empty > cells.length ? `${empty} / ${total} 为空` : distinct.length === cells.length ? `${cells.length} 行各不相同` : '取值较多',
      detail: '不设置',
    });
  }
  return out;
}

/** 表头下面这一列里不同的取值（按出现顺序，最多 limit 个）；多选按分隔符拆开 */
export function columnValues(sheet: SheetData, header: number, col: number, split = false, limit = 200): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (let r = header; r < sheet.rows && out.length < limit; r++) {
    const t = cellText(sheet, r, col).trim();
    for (const v of t ? (split ? splitMulti(t) : [t]) : []) {
      if (!seen.has(v)) {
        seen.add(v);
        out.push(v);
      }
    }
  }
  return out;
}

// ---------- 必填 ----------

/** 必填列里空着的格子（只看已经有内容的行，整行空的不算；最多返回 limit 个） */
export function missingRequired(sheet: SheetData, limit = 1000): { r: number; c: number }[] {
  const sc = sheet.columns;
  const out: { r: number; c: number }[] = [];
  if (!sc) return out;
  const req = sc.columns.filter((c) => c.required && !c.unbound);
  if (!req.length) return out;
  for (let r = sc.header; r < sheet.rows && out.length < limit; r++) {
    const row = sheet.cells[r];
    if (!row?.some((x) => x && x.t !== '')) continue;
    for (const c of req) if (!(row[c.col]?.t ?? '').trim()) out.push({ r, c: c.col });
  }
  return out;
}

/** 复选框值 → 是否勾选 */
export const isChecked = (text: string): boolean => isCheckOn(text);

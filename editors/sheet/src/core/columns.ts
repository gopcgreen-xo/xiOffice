/**
 * 列属性的操作（改模型，进撤销栈）：设置 / 清除某一列的字段类型、按类型给数据设数字格式、批量应用推断结果、
 * 录入时补默认值。纯逻辑（取选项、位移、推断）在 colmeta.ts。列属性存在工作包里，不写进 xlsx。
 */
import { newColumnId, resolveDefault, splitMulti, type ColumnMeta, type ColumnType, type SheetColumns } from '@wz/editor-shared';
import { cellText } from './colmeta';
import type { Inference } from './colmeta';
import { parseInput } from './input';
import { applyStyle, setInput, type Tx } from './ops';
import { formatCode } from './styles';
import type { SheetData } from './workbook';

/** 某类型的数字格式（列属性 fmt 优先）；没有格式要求的类型返回 undefined */
export function typeFormat(c: Pick<ColumnMeta, 'type' | 'fmt'>): string | undefined {
  if (c.fmt) return c.fmt;
  switch (c.type) {
    case 'money':
      return '#,##0.00';
    case 'percent':
      return '0.00%';
    case 'date':
    case 'created':
      return 'yyyy-mm-dd';
    default:
      return undefined;
  }
}

/** 新的列属性（id 自动取；表头文字取表头行里这一列的文字） */
export function makeMeta(sheet: SheetData, col: number, type: ColumnType, extra: Partial<ColumnMeta> = {}): ColumnMeta {
  const header = sheet.columns?.header ?? 1;
  return {
    id: newColumnId(sheet.columns?.columns ?? []),
    title: cellText(sheet, header - 1, col).trim(),
    col,
    type,
    ...(type === 'select' || type === 'multi' ? { options: [] } : {}),
    ...extra,
  };
}

/** 把单元格里要写的文字变成 setInput 能原样存成文字的形式（「1,2」这类会被当成数字的加前导 '） */
export function literalInput(text: string): string {
  if (!text) return text;
  const p = parseInput(text);
  return p.kind === 'value' && typeof p.v !== 'string' ? `'${text}` : p.kind === 'formula' ? `'${text}` : text;
}

/** 给这一列的已有数据（表头行以下）设数字格式；之后在这一列输入数字时沿用它（空格子也带上格式） */
export function applyColumnFormat(tx: Tx, si: number, c: ColumnMeta, header: number): void {
  const fmt = typeFormat(c);
  const sheet = tx.book.sheets[si];
  if (!fmt || sheet.rows <= header) return;
  applyStyle(tx, si, { r0: header, c0: c.col, r1: sheet.rows - 1, c1: c.col }, { numFmt: fmt });
}

/** 设置（meta）或清除（meta = null）一列的列属性；一次改动 = 一步撤销 */
export function setColumn(tx: Tx, si: number, col: number, meta: ColumnMeta | null): void {
  const sheet = tx.book.sheets[si];
  tx.columns(si);
  const sc: SheetColumns = sheet.columns ?? { header: 1, columns: [] };
  const rest = sc.columns.filter((c) => c.col !== col || c.unbound);
  const next: SheetColumns = { ...sc, columns: meta ? [...rest, { ...meta, col }].sort((a, b) => a.col - b.col) : rest };
  if (next.columns.length || next.ext) sheet.columns = next;
  else delete sheet.columns;
  if (meta) applyColumnFormat(tx, si, { ...meta, col }, next.header);
}

/** 表头所在行（1 起）；已有列属性时不动 */
export function setHeaderRow(tx: Tx, si: number, header: number): void {
  const sheet = tx.book.sheets[si];
  tx.columns(si);
  sheet.columns = { ...(sheet.columns ?? { columns: [] }), header };
}

/** 批量应用推断结果：文本列不设置；选项取推断出的取值 */
export function applyInferences(tx: Tx, si: number, header: number, list: Inference[]): number {
  const sheet = tx.book.sheets[si];
  tx.columns(si);
  sheet.columns = { ...(sheet.columns ?? { columns: [] }), header };
  let n = 0;
  for (const x of list) {
    if (x.type === 'text') continue;
    const meta = makeMeta(sheet, x.col, x.type, {
      ...(x.options && { options: x.options.map((v) => ({ v })) }),
      ...(x.integer && { integer: true }),
      ...(x.fmt && { fmt: x.fmt }),
    });
    setColumn(tx, si, x.col, meta);
    n++;
  }
  return n;
}

/** 在列属性的格式下写入一个格子：先按输入写，再把数字格式套成这一列的格式 */
export function setColumnInput(tx: Tx, si: number, r: number, c: number, meta: ColumnMeta | undefined, text: string): void {
  const sheet = tx.book.sheets[si];
  const literal = meta && (meta.type === 'select' || meta.type === 'multi' || meta.type === 'check' || meta.type === 'link' || meta.type === 'email' || meta.type === 'phone');
  setInput(tx, si, r, c, literal ? literalInput(meta.type === 'multi' ? splitMulti(text).join(',') : text) : text);
  const fmt = meta && typeFormat(meta);
  const cell = sheet.cells[r]?.[c];
  if (fmt && typeof cell?.v === 'number' && tx.book.styleBook && formatCode(tx.book.styleBook, cell.s) !== fmt) {
    applyStyle(tx, si, { r0: r, c0: c, r1: r, c1: c }, { numFmt: fmt });
  }
}

/** 这一行刚开始录入：给其它有默认值、还空着的列填上默认值（整行没填过任何东西才填，不覆盖已有内容） */
export function fillDefaults(tx: Tx, si: number, r: number, skipCol: number, user: string | null, now = new Date()): number {
  const sheet = tx.book.sheets[si];
  const sc = sheet.columns;
  if (!sc || r < sc.header) return 0;
  let n = 0;
  for (const c of sc.columns) {
    if (c.unbound || c.col === skipCol || c.default === undefined || c.default === '') continue;
    if (cellText(sheet, r, c.col).trim()) continue;
    setColumnInput(tx, si, r, c.col, c, resolveDefault(c.default, now, user));
    n++;
  }
  return n;
}

/** 这一行是不是还没有任何内容（除了 exceptCol） */
export function rowIsBlank(sheet: SheetData, r: number, exceptCol: number): boolean {
  return !(sheet.cells[r] ?? []).some((cell, c) => c !== exceptCol && cell && cell.t !== '');
}

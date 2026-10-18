/**
 * 查找与替换（同 Excel 的 Ctrl+F / Ctrl+H）：在当前工作表或整个工作簿里按行查找，
 * 可以区分大小写、单元格完全匹配、查找“公式”（编辑时看到的原文）或“值”（显示的文字）。
 * 替换改的是格子的原文（公式里的文字也会替换），替换后按输入重新识别。
 */
import { formatCode } from './styles';
import { formatNumber, isDateFormat } from './numfmt';
import { parseInput } from './input';
import { setInput, type Tx } from './ops';
import type { SCell, Workbook } from './workbook';

export interface FindOptions {
  query: string;
  matchCase?: boolean;
  /** 单元格完全匹配 */
  whole?: boolean;
  /** 查找范围：formulas = 原文（公式带 =），values = 显示的文字 */
  lookIn?: 'formulas' | 'values';
  /** 只在这张表里找；undefined = 整个工作簿 */
  si?: number;
}

export interface Hit {
  si: number;
  r: number;
  c: number;
}

/** 编辑时显示的原文（同 Excel：公式带 =，日期按短日期，百分比带 %；看起来像数字的文字前加 '） */
export function editText(book: Workbook, cell: SCell | undefined): string {
  if (!cell) return '';
  const fmt = book.styleBook ? formatCode(book.styleBook, cell.s) : 'General';
  if (cell.f) return `=${cell.f}`;
  if (typeof cell.v === 'boolean') return cell.v ? 'TRUE' : 'FALSE';
  if (typeof cell.v === 'number') {
    if (isDateFormat(fmt)) {
      const v = cell.v;
      if (v < 1) return formatNumber(v, 'h:mm:ss');
      return formatNumber(v, v - Math.floor(v) > 1e-9 ? 'yyyy/m/d h:mm:ss' : 'yyyy/m/d');
    }
    if (/%/.test(fmt.replace(/"[^"]*"/g, ''))) return `${Number((cell.v * 100).toPrecision(15))}%`;
    return String(cell.v);
  }
  const p = parseInput(cell.v, fmt, book.date1904);
  return cell.k || (p.kind === 'value' && typeof p.v === 'string') || fmt === '@' ? String(cell.v) : `'${cell.v}`;
}

function matcher(o: FindOptions): ((text: string) => boolean) | null {
  if (!o.query) return null;
  const q = o.matchCase ? o.query : o.query.toLowerCase();
  return (text) => {
    const t = o.matchCase ? text : text.toLowerCase();
    return o.whole ? t === q : t.includes(q);
  };
}

const target = (book: Workbook, cell: SCell, o: FindOptions) => (o.lookIn === 'values' ? cell.t : editText(book, cell).replace(/^'/, ''));

/** 所有匹配的格子：按表、行、列的顺序 */
export function findAll(book: Workbook, o: FindOptions): Hit[] {
  const test = matcher(o);
  if (!test) return [];
  const out: Hit[] = [];
  book.sheets.forEach((s, si) => {
    if (o.si !== undefined && si !== o.si) return;
    if (o.si === undefined && s.hidden) return;
    s.cells.forEach((row, r) =>
      row?.forEach((cell, c) => {
        if (cell && test(target(book, cell, o))) out.push({ si, r, c });
      }),
    );
  });
  return out;
}

/** 从 (si, r, c) 之后（不含）找下一个，找到最后从头再来 */
export function findNext(hits: Hit[], from: Hit, back = false): Hit | null {
  if (!hits.length) return null;
  const key = (h: Hit) => h.si * 1e12 + h.r * 16384 + h.c;
  const k = key(from);
  if (back) return [...hits].reverse().find((h) => key(h) < k) ?? hits[hits.length - 1];
  return hits.find((h) => key(h) > k) ?? hits[0];
}

/** 文字里的替换（不区分大小写时按原文位置替换；完全匹配时整个换掉） */
export function replaceText(text: string, o: FindOptions, repl: string): string {
  if (o.whole) return repl;
  if (o.matchCase) return text.split(o.query).join(repl);
  const low = text.toLowerCase();
  const q = o.query.toLowerCase();
  let out = '';
  let i = 0;
  for (let j = low.indexOf(q); j >= 0; j = low.indexOf(q, i)) {
    out += text.slice(i, j) + repl;
    i = j + q.length;
  }
  return out + text.slice(i);
}

/** 替换一个格子（改它的原文后按输入重新识别）；不匹配时返回 false */
export function replaceCell(tx: Tx, hit: Hit, o: FindOptions, repl: string): boolean {
  const book = tx.book;
  const cell = book.sheets[hit.si].cells[hit.r]?.[hit.c];
  const test = matcher({ ...o, lookIn: 'formulas' });
  if (!cell || !test) return false;
  const raw = editText(book, cell);
  const body = raw.replace(/^'/, '');
  if (!test(body)) return false;
  const next = replaceText(body, o, repl);
  setInput(tx, hit.si, hit.r, hit.c, raw.startsWith("'") ? `'${next}` : next);
  return true;
}

/** 全部替换，返回替换的格子数 */
export function replaceAll(tx: Tx, o: FindOptions, repl: string): number {
  let n = 0;
  for (const h of findAll(tx.book, { ...o, lookIn: 'formulas' })) if (replaceCell(tx, h, o, repl)) n++;
  return n;
}

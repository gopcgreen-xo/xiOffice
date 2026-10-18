/**
 * 数据工具（同 Excel “数据”选项卡）：
 * - 删除重复项：按选中的列比较（文字不区分大小写，比较显示的值），保留第一次出现的行；只动选区里的格子，
 *   后面的行在选区内上移，空出的行清空（选区外的格子不动）。
 * - 分列：一列文字按分隔符拆到右边几列，每段按输入识别（数字、日期、前导 0 的编号保持文字）。
 */
import { moveFormula } from './formula/refs';
import { setInput, type Tx } from './ops';
import type { Range, SCell, SheetData } from './workbook';

/** 比较用的键：显示文字（没有时用值），不区分大小写 */
const keyOf = (c: SCell | undefined) => (c ? String(c.t ?? c.v ?? '') : '').toLowerCase();

/** 删除重复项。cols 为参与比较的列（绝对列号）；返回 [删除的行数, 留下的行数]；有合并单元格时返回 null */
export function removeDuplicates(tx: Tx, si: number, g: Range, cols: number[], header: boolean): [number, number] | null {
  const sheet = tx.book.sheets[si];
  const r0 = header ? g.r0 + 1 : g.r0;
  if (r0 > g.r1 || !cols.length) return [0, Math.max(0, g.r1 - r0 + 1)];
  if (sheet.merges.some((m) => m.r1 >= r0 && m.r0 <= g.r1 && m.c1 >= g.c0 && m.c0 <= g.c1)) return null;
  const seen = new Set<string>();
  const keep: number[] = [];
  for (let r = r0; r <= g.r1; r++) {
    const k = cols.map((c) => keyOf(sheet.cells[r]?.[c])).join('\u0000');
    if (seen.has(k)) continue;
    seen.add(k);
    keep.push(r);
  }
  const removed = g.r1 - r0 + 1 - keep.length;
  if (!removed) return [0, keep.length];
  tx.snap(si);
  const rows = keep.map((r) => ({
    r,
    cells: Array.from({ length: g.c1 - g.c0 + 1 }, (_, j) => sheet.cells[r]?.[g.c0 + j]),
    blanks: Array.from({ length: g.c1 - g.c0 + 1 }, (_, j) => sheet.blankStyles?.[r]?.[g.c0 + j]),
  }));
  for (let i = 0; i <= g.r1 - r0; i++) {
    const r = r0 + i;
    const src = rows[i];
    for (let j = 0; j <= g.c1 - g.c0; j++) {
      const c = g.c0 + j;
      const cell = src?.cells[j];
      const next = cell ? { ...cell, ...(cell.f && { f: moveFormula(cell.f, r - src.r, 0) }) } : undefined;
      if (next) (sheet.cells[r] ??= [])[c] = next;
      else if (sheet.cells[r]) sheet.cells[r][c] = undefined;
      const bs = src ? src.blanks[j] : undefined;
      if (bs !== undefined) ((sheet.blankStyles ??= [])[r] ??= [])[c] = bs;
      else if (sheet.blankStyles?.[r]) sheet.blankStyles[r][c] = undefined;
    }
  }
  return [removed, keep.length];
}

export interface SplitOptions {
  /** 分隔符（每个都是一个字符） */
  delims: string[];
  /** 连续分隔符视为一个 */
  merge: boolean;
}

/** 按分隔符拆开一段文字（双引号里的分隔符不算，同 Excel 的“文本识别符号”） */
export function splitText(text: string, o: SplitOptions): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      if (quoted && text[i + 1] === '"') {
        cur += '"';
        i++;
      } else quoted = !quoted;
      continue;
    }
    if (!quoted && o.delims.includes(ch)) {
      out.push(cur);
      cur = '';
      if (o.merge) while (i + 1 < text.length && o.delims.includes(text[i + 1])) i++;
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

/** 分列后最多几列（预览、检查右边是否有数据用） */
export function splitWidth(sheet: SheetData, g: Range, o: SplitOptions): number {
  let n = 1;
  for (let r = g.r0; r <= g.r1; r++) {
    const c = sheet.cells[r]?.[g.c0];
    if (c && typeof c.v === 'string') n = Math.max(n, splitText(c.v, o).length);
  }
  return n;
}

/** 分列：只处理选区第一列里的文字格子；第 k 段写到右边第 k 列（覆盖原内容）。返回拆出来的最多列数 */
export function textToColumns(tx: Tx, si: number, g: Range, o: SplitOptions): number {
  const sheet = tx.book.sheets[si];
  let width = 1;
  for (let r = g.r0; r <= g.r1; r++) {
    const cell = sheet.cells[r]?.[g.c0];
    if (!cell || typeof cell.v !== 'string' || cell.f) continue;
    const parts = splitText(cell.v, o);
    width = Math.max(width, parts.length);
    if (parts.length < 2) continue;
    parts.forEach((p, k) => setInput(tx, si, r, g.c0 + k, p.trim() === '' ? '' : p));
  }
  return width;
}

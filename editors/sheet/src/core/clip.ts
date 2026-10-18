/** 剪贴板 HTML 表格（从网页、Word、WPS 复制）→ 粘贴块；CSV 字节解码（编码与 BOM 保存时原样写回）。 */
import type { Block, ClipCell } from './ops';

/** HTML 里第一个表格 → 只含文字的块（合并单元格的文字放在左上角）；没有表格返回 null */
export function blockFromHtml(html: string, doc?: Document): Block | null {
  if (!/<table\b/i.test(html)) return null;
  const d = doc ?? new DOMParser().parseFromString(html, 'text/html');
  const table = d.querySelector('table');
  if (!table) return null;
  const grid: (ClipCell | undefined)[][] = [];
  const taken = new Set<string>();
  let cols = 0;
  table.querySelectorAll('tr').forEach((tr, r) => {
    let c = 0;
    grid[r] ??= [];
    tr.querySelectorAll('td, th').forEach((cell) => {
      while (taken.has(`${r}:${c}`)) c++;
      const rs = Math.max(1, Number(cell.getAttribute('rowspan')) || 1);
      const cs = Math.max(1, Number(cell.getAttribute('colspan')) || 1);
      cell.querySelectorAll('br').forEach((br) => br.replaceWith('\n'));
      const text = (cell.textContent ?? '').replace(/ /g, ' ').replace(/[ \t]*\n[ \t]*/g, '\n').trim();
      if (text) grid[r][c] = { v: text, t: text, r, c };
      for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) taken.add(`${r + i}:${c + j}`);
      c += cs;
      cols = Math.max(cols, c);
    });
  });
  const rows = grid.length;
  if (!rows || !cols) return null;
  for (let r = 0; r < rows; r++) grid[r] = Array.from({ length: cols }, (_, c) => grid[r]?.[c]);
  return { rows, cols, cells: grid, internal: false };
}

/** CSV 文件字节 → 文字：BOM → UTF-8 → GBK（中文 Excel 存的 CSV 多是 GBK） */
export function decodeCsvBytes(bytes: Uint8Array): { text: string; encoding: string; bom: boolean } {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { text: new TextDecoder('utf-8').decode(bytes.subarray(3)), encoding: 'UTF-8', bom: true };
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(bytes.subarray(2)), encoding: 'UTF-8', bom: true };
  try {
    return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes), encoding: 'UTF-8', bom: false };
  } catch {
    return { text: new TextDecoder('gbk').decode(bytes), encoding: 'GBK', bom: false };
  }
}

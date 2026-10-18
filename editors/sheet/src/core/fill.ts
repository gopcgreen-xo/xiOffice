/**
 * 填充柄（拖动选区右下角）与 Ctrl+D / Ctrl+R，规则同 Excel：
 * - 一个数字：复制；两个以上的数字：按等差（最小二乘直线）延续；
 * - 日期（日期格式的数字）：一个时按天递增，多个时按直线延续；
 * - 结尾带数字的文字（第1项、A-01）：数字递增，补零位数保持；源里全是同一模式时按它们的步长；
 * - 其它文字、布尔：按源循环复制；公式：相对引用按位移平移；格式按源循环复制。
 */
import { formatCode } from './styles';
import { isDateFormat } from './numfmt';
import { moveFormula } from './formula/refs';
import { refreshText, styleAt, type Range, type SCell } from './workbook';
import type { Tx } from './ops';

type Dir = 'down' | 'up' | 'right' | 'left';

const TEXT_NUM = /^(.*?)(\d+)(\D*)$/;

function linear(ys: number[]): (x: number) => number {
  const n = ys.length;
  const mx = (n - 1) / 2;
  const my = ys.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  ys.forEach((y, x) => {
    num += (x - mx) * (y - my);
    den += (x - mx) * (x - mx);
  });
  const k = den ? num / den : 0;
  return (x) => Number((my + k * (x - mx)).toPrecision(15));
}

/** 源区域 src 扩展到 to（to 包含 src，只在一个方向上变大） */
export function fillSeries(tx: Tx, si: number, src: Range, to: Range): void {
  const book = tx.book;
  const sheet = book.sheets[si];
  const dir: Dir | null = to.r1 > src.r1 ? 'down' : to.r0 < src.r0 ? 'up' : to.c1 > src.c1 ? 'right' : to.c0 < src.c0 ? 'left' : null;
  if (!dir) return;
  const vertical = dir === 'down' || dir === 'up';
  const lines = vertical ? [src.c0, src.c1] : [src.r0, src.r1];
  const n = vertical ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1;
  const count = vertical ? to.r1 - to.r0 + 1 - n : to.c1 - to.c0 + 1 - n;
  for (let line = lines[0]; line <= lines[1]; line++) {
    const pos = (i: number) => (vertical ? { r: src.r0 + i, c: line } : { r: line, c: src.c0 + i });
    const items = Array.from({ length: n }, (_, i) => {
      const p = pos(i);
      return { ...p, cell: sheet.cells[p.r]?.[p.c], s: styleAt(sheet, p.r, p.c) };
    });
    const isNum = (x: (typeof items)[number]) => typeof x.cell?.v === 'number' && !x.cell.f;
    const allNum = items.every(isNum);
    const isDate = allNum && book.styleBook && isDateFormat(formatCode(book.styleBook, items[0].cell?.s));
    const fit = allNum && (n > 1 || isDate) ? linear(n > 1 ? items.map((x) => x.cell!.v as number) : [items[0].cell!.v as number, (items[0].cell!.v as number) + 1]) : null;
    // 全是同一模式的“文字+数字”：数字按直线延续
    const tm = items.map((x) => (typeof x.cell?.v === 'string' && !x.cell.f ? x.cell.v.match(TEXT_NUM) : null));
    const sameText = n > 1 && tm.every((m) => m && m[1] === tm[0]![1] && m[3] === tm[0]![3]);
    const textFit = sameText ? linear(tm.map((m) => Number(m![2]))) : null;

    for (let k = 1; k <= count; k++) {
      // 向下 / 右：源后面第 k 个；向上 / 左：源前面第 k 个
      const forward = dir === 'down' || dir === 'right';
      const x = forward ? n - 1 + k : -k;
      const i = ((x % n) + n) % n;
      const cycle = Math.floor(x / n);
      const item = items[i];
      const tp = vertical ? { r: forward ? src.r1 + k : src.r0 - k, c: line } : { r: line, c: forward ? src.c1 + k : src.c0 - k };
      tx.touch(si, tp.r, tp.c);
      const s = item.s;
      let cell: SCell | undefined;
      const c0 = item.cell;
      if (!c0) cell = undefined;
      else if (c0.f) cell = { v: c0.v, t: c0.t, f: moveFormula(c0.f, tp.r - item.r, tp.c - item.c) };
      else if (fit) cell = { v: fit(x), t: '' };
      else if (textFit && tm[i]) {
        const m = tm[i]!;
        const v = Math.max(0, Math.round(textFit(x)));
        cell = { v: `${m[1]}${String(v).padStart(m[2].length, '0')}${m[3]}`, t: '' };
      } else if (tm[i] && !c0.f) {
        const m = tm[i]!;
        const v = Math.max(0, Number(m[2]) + cycle);
        cell = { v: `${m[1]}${String(v).padStart(m[2].length, '0')}${m[3]}`, t: '' };
      } else cell = { v: c0.v, t: '', ...(c0.k && { k: c0.k }) };
      const row = (sheet.cells[tp.r] ??= []);
      if (cell) {
        if (s) cell.s = s;
        refreshText(book, cell);
        row[tp.c] = cell;
        if (sheet.blankStyles?.[tp.r]) sheet.blankStyles[tp.r][tp.c] = undefined;
      } else {
        row[tp.c] = undefined;
        if (s) ((sheet.blankStyles ??= [])[tp.r] ??= [])[tp.c] = s;
        else if (sheet.blankStyles?.[tp.r]) sheet.blankStyles[tp.r][tp.c] = undefined;
      }
    }
  }
}

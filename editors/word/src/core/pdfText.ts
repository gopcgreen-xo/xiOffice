/**
 * 导出 PDF 的文字层：从排好版的页面里取出每一行文字的位置，给 PDF 写成不可见文字（可选中、复制、查找）。
 * 位置、字号都取自画页面用的同一份排版，所以和看到的文字对得上。
 *
 * 一行在图片、公式、带圈字符、目录前导符、制表符处断开成几段，各段单独定位；
 * 暂不处理：文本框 / 浮动对象里的文字、竖排文字（这些文字在 PDF 里只有图像）。
 */
import type { Glyph, PageLayout, PlacedLine } from './layout';
import { FIELD_CHAR } from './model';

export interface PdfTextLine {
  text: string;
  /** 基线起点（PDF 用户空间，原点左下，磅） */
  origin: [number, number];
  /** 字号（磅） */
  size: number;
  /** 这一段的宽度（磅） */
  width: number;
}

/** 不算文字内容的字形：图片、公式、带圈字符、前导符、域占位符、制表符、符号字体的私用区字符 */
function isBreak(g: Glyph): boolean {
  if (g.image || g.math || g.eq || g.leader) return true;
  const c = g.ch.charCodeAt(0);
  return g.ch === FIELD_CHAR || g.ch === '\t' || (c >= 0xf000 && c <= 0xf0ff);
}

function addLine(out: PdfTextLine[], pl: PlacedLine, pageH: number): void {
  const base = pl.y + pl.line.baseline;
  let seg: Glyph[] = [];
  const flush = () => {
    const text = seg.map((g) => g.ch).join('');
    if (text.trim()) {
      const first = seg[0]!;
      const last = seg[seg.length - 1]!;
      out.push({
        text: text.replace(/[\r\n\u000b]/g, ' '),
        origin: [pl.x + first.x, pageH - base],
        size: Math.max(1, Math.max(...seg.map((g) => g.props.size))),
        width: Math.max(0, last.x + last.w - first.x),
      });
    }
    seg = [];
  };
  for (const g of pl.line.glyphs) {
    if (isBreak(g)) flush();
    else seg.push(g);
  }
  flush();
}

/** 一页的文字段（正文、页眉、页脚、脚注） */
export function pageTextLines(page: PageLayout): PdfTextLine[] {
  const out: PdfTextLine[] = [];
  const h = page.setup.height;
  for (const pl of [...page.header, ...page.body, ...page.footer, ...(page.notes ?? [])]) addLine(out, pl, h);
  return out;
}

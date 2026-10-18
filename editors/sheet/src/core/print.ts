/**
 * 打印与导出 PDF：按页面设置（纸张、方向、缩放 / 调整为几页、打印区域、每页重复标题行、网格线）分页，
 * 再把每一页画到画布上（底色、文字、对齐、自动换行、合并单元格、边框）。
 * 分页规则同 Excel：页边距左右 0.7 英寸、上下 0.75 英寸；先向下再向右；隐藏的行列不打印。
 */
import { buildMergeIndex, cellBorder, cellKey } from './grid';
import { DEFAULT_MARGINS, styleAt, type CellStyle, type PageSetup, type Range, type SheetData } from './workbook';

/** 纸张大小（磅，纵向） */
export const PAPERS: Record<number, { name: string; w: number; h: number }> = {
  9: { name: 'A4', w: 595.3, h: 841.9 },
  8: { name: 'A3', w: 841.9, h: 1190.6 },
  11: { name: 'A5', w: 419.5, h: 595.3 },
  1: { name: 'Letter', w: 612, h: 792 },
};

/** 页边距（磅）：文件里是英寸 */
function marginsPt(page: PageSetup) {
  const m = page.margins ?? DEFAULT_MARGINS;
  return { l: m.left * 72, r: m.right * 72, t: m.top * 72, b: m.bottom * 72, header: m.header * 72, footer: m.footer * 72 };
}

/** 页眉页脚代码里的变量 */
export interface HFVars {
  page: number;
  pages: number;
  sheet: string;
  file: string;
  date: Date;
}

/**
 * 解析 Excel 页眉页脚代码：&L / &C / &R 切换左中右三段（没写时是中间），&P 页码、&N 总页数、&D 日期、&T 时间、
 * &A 表名、&F 文件名、&& 一个 &；字体（&"字体,样式"）、字号（&数字）、&B &I &U &S 等格式代码去掉。
 */
export function headerFooterText(code: string, v: HFVars): { left: string; center: string; right: string } {
  const out = { left: '', center: '', right: '' };
  let cur: keyof typeof out = 'center';
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch !== '&') {
      out[cur] += ch;
      continue;
    }
    const k = code[++i];
    if (k === undefined) break;
    if (k === '&') out[cur] += '&';
    else if (k === 'L') cur = 'left';
    else if (k === 'C') cur = 'center';
    else if (k === 'R') cur = 'right';
    else if (k === 'P') {
      // &P+1 / &P-1
      const m = code.slice(i + 1).match(/^([+-]\d+)/);
      out[cur] += String(v.page + (m ? Number(m[1]) : 0));
      if (m) i += m[1].length;
    } else if (k === 'N') out[cur] += String(v.pages);
    else if (k === 'D') out[cur] += `${v.date.getFullYear()}/${v.date.getMonth() + 1}/${v.date.getDate()}`;
    else if (k === 'T') out[cur] += `${v.date.getHours()}:${String(v.date.getMinutes()).padStart(2, '0')}`;
    else if (k === 'A') out[cur] += v.sheet;
    else if (k === 'F') out[cur] += v.file;
    else if (k === 'Z') continue; // 文件路径：不显示
    else if (k === '"') {
      const end = code.indexOf('"', i + 1);
      i = end < 0 ? code.length : end;
    } else if (/\d/.test(k)) {
      while (/\d/.test(code[i + 1] ?? '')) i++;
    }
    // 其它（&B &I &U &S &X &Y &E &G &K 颜色…）是格式代码，忽略
    else if (k === 'K') i += 6;
  }
  return out;
}

/** 按分页符切开：新的一页从 breaks 里的行 / 列开始 */
function splitAt(items: number[], breaks: number[] | undefined): number[][] {
  if (!breaks?.length) return [items];
  const out: number[][] = [];
  let cur: number[] = [];
  for (const i of items) {
    if (cur.length && breaks.some((b) => b <= i && b > cur[cur.length - 1])) {
      out.push(cur);
      cur = [];
    }
    cur.push(i);
  }
  if (cur.length) out.push(cur);
  return out;
}

export interface PrintPage {
  /** 这一页打印的行（含重复的标题行，在最前面）与列 */
  rows: number[];
  cols: number[];
}

export interface PrintLayout {
  /** 纸张（磅，已按方向转好） */
  w: number;
  h: number;
  scale: number;
  pages: PrintPage[];
  colPt: (c: number) => number;
  rowPt: (r: number) => number;
  area: Range;
  /** 页边距（磅） */
  margin: { l: number; r: number; t: number; b: number; header: number; footer: number };
  page: PageSetup;
}

/** 已用区域：有内容、有样式的格子与合并区域 */
export function usedRange(sheet: SheetData): Range | null {
  let r1 = -1;
  let c1 = -1;
  sheet.cells.forEach((row, r) =>
    row?.forEach((cell, c) => {
      if (!cell) return;
      r1 = Math.max(r1, r);
      c1 = Math.max(c1, c);
    }),
  );
  for (const m of sheet.merges) {
    r1 = Math.max(r1, m.r1);
    c1 = Math.max(c1, m.c1);
  }
  return r1 < 0 ? null : { r0: 0, c0: 0, r1, c1 };
}

/** 一组尺寸按可用长度切段（单个超长的也自成一段） */
function chunks(items: number[], size: (i: number) => number, avail: number): number[][] {
  const out: number[][] = [];
  let cur: number[] = [];
  let used = 0;
  for (const i of items) {
    const s = size(i);
    if (cur.length && used + s > avail + 0.01) {
      out.push(cur);
      cur = [];
      used = 0;
    }
    cur.push(i);
    used += s;
  }
  if (cur.length) out.push(cur);
  return out;
}

export function layoutPrint(sheet: SheetData, page: PageSetup = {}, area?: Range): PrintLayout {
  const paper = PAPERS[page.paper ?? 9] ?? PAPERS[9];
  const w = page.landscape ? paper.h : paper.w;
  const h = page.landscape ? paper.w : paper.h;
  const g = area ?? page.area ?? usedRange(sheet) ?? { r0: 0, c0: 0, r1: 0, c1: 0 };
  const colPt = (c: number) => (sheet.colWidths[c] ?? sheet.defaultColWidth) * 0.75;
  const rowPt = (r: number) => (sheet.rowHeights[r] ?? sheet.defaultRowHeight) * 0.75;
  const cols: number[] = [];
  const rows: number[] = [];
  for (let c = g.c0; c <= g.c1; c++) if (colPt(c) > 0) cols.push(c);
  for (let r = g.r0; r <= g.r1; r++) if (rowPt(r) > 0) rows.push(r);
  const MARGIN = marginsPt(page);
  const availW = Math.max(36, w - MARGIN.l - MARGIN.r);
  const availH = Math.max(36, h - MARGIN.t - MARGIN.b);
  const totalW = cols.reduce((a, c) => a + colPt(c), 0);
  // 标题行在打印区域外面时也打印（同 Excel）
  const titles: number[] = [];
  if (page.titleRows) for (let r = page.titleRows[0]; r <= page.titleRows[1]; r++) if (rowPt(r) > 0) titles.push(r);
  const titleH = titles.reduce((a, r) => a + rowPt(r), 0);
  const totalH = rows.reduce((a, r) => a + rowPt(r), 0);
  let scale = Math.max(10, Math.min(400, page.scale ?? 100)) / 100;
  if (page.fitW !== undefined || page.fitH !== undefined) {
    const sw = page.fitW ? (page.fitW * availW) / Math.max(1, totalW) : Infinity;
    const sh = page.fitH ? (page.fitH * availH) / Math.max(1, totalH) : Infinity;
    scale = Math.max(0.1, Math.min(1, sw, sh));
  }
  // 手动分页符先切开，每段再按纸张大小自动分页（“调整为一页”时 Excel 也仍按手动分页符分页）
  const colGroups = splitAt(cols, page.colBreaks).flatMap((part) => chunks(part, colPt, availW / scale));
  // 每页 = 标题行 + 接下来的正文行（标题行本来就在区域顶部时也一样，只打印一次）
  const rest = rows.filter((r) => !titles.includes(r));
  const rowGroups: number[][] = splitAt(rest, page.rowBreaks)
    .flatMap((part) => chunks(part, rowPt, availH / scale - titleH))
    .map((grp) => [...titles, ...grp]);
  if (!rowGroups.length) rowGroups.push([]);
  if (!colGroups.length) colGroups.push([]);
  const pages: PrintPage[] = [];
  for (const cg of colGroups) for (const rg of rowGroups) pages.push({ rows: rg, cols: cg });
  return { w, h, scale, pages, colPt, rowPt, area: g, margin: MARGIN, page };
}

// ---------- 画一页 ----------

function fontOf(st: CellStyle | undefined, defSize: number, defFont: string): string {
  return `${st?.italic ? 'italic ' : ''}${st?.bold ? 'bold ' : ''}${st?.fontSize ?? defSize}pt "${st?.fontName ?? defFont}", "Microsoft YaHei", sans-serif`;
}

/** 按宽度折行（中文逐字，西文尽量在空格处断） */
function wrapLines(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    let line = '';
    for (const ch of [...para]) {
      if (line && ctx.measureText(line + ch).width > width) {
        const sp = line.lastIndexOf(' ');
        if (sp > 0 && /\w/.test(ch)) {
          out.push(line.slice(0, sp));
          line = line.slice(sp + 1) + ch;
        } else {
          out.push(line);
          line = ch;
        }
      } else line += ch;
    }
    out.push(line);
  }
  return out;
}

export interface RenderOptions {
  styles: CellStyle[];
  defSize: number;
  defFont: string;
  gridLines?: boolean;
  /** 每磅多少像素（2 = 144 dpi） */
  pxPerPt?: number;
  /** 页眉页脚里的表名、文件名 */
  sheetName?: string;
  fileName?: string;
}

export function renderPage(sheet: SheetData, layout: PrintLayout, page: PrintPage, o: RenderOptions): HTMLCanvasElement {
  const k = o.pxPerPt ?? 2;
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(layout.w * k);
  canvas.height = Math.round(layout.h * k);
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.scale(k, k);
  // 页眉页脚：在纸张上（不缩放），距上 / 下边 header / footer 边距，左中右三段
  const M = layout.margin;
  const vars: HFVars = { page: layout.pages.indexOf(page) + 1, pages: layout.pages.length, sheet: o.sheetName ?? '', file: o.fileName ?? '', date: new Date() };
  const hf = (code: string | undefined, y: number, base: CanvasTextBaseline) => {
    if (!code) return;
    const t = headerFooterText(code, vars);
    ctx.font = `${o.defSize}pt "${o.defFont}", "Microsoft YaHei", sans-serif`;
    ctx.fillStyle = '#000';
    ctx.textBaseline = base;
    ctx.textAlign = 'left';
    ctx.fillText(t.left, M.l, y);
    ctx.textAlign = 'center';
    ctx.fillText(t.center, layout.w / 2, y);
    ctx.textAlign = 'right';
    ctx.fillText(t.right, layout.w - M.r, y);
  };
  hf(layout.page.header, M.header, 'top');
  hf(layout.page.footer, layout.h - M.footer, 'bottom');
  // 内容大小（居中用）
  const contentW = page.cols.reduce((a, c) => a + layout.colPt(c), 0) * layout.scale;
  const contentH = page.rows.reduce((a, r) => a + layout.rowPt(r), 0) * layout.scale;
  const offX = layout.page.centerH ? Math.max(0, (layout.w - M.l - M.r - contentW) / 2) : 0;
  const offY = layout.page.centerV ? Math.max(0, (layout.h - M.t - M.b - contentH) / 2) : 0;
  ctx.translate(M.l + offX, M.t + offY);
  ctx.scale(layout.scale, layout.scale);
  // 每行 / 列在本页的位置
  const x = new Map<number, number>();
  const y = new Map<number, number>();
  let acc = 0;
  for (const c of page.cols) {
    x.set(c, acc);
    acc += layout.colPt(c);
  }
  const width = acc;
  acc = 0;
  for (const r of page.rows) {
    y.set(r, acc);
    acc += layout.rowPt(r);
  }
  const height = acc;
  const merges = buildMergeIndex(sheet.merges);
  const st = (r: number, c: number) => {
    const s = styleAt(sheet, r, c);
    return s !== undefined ? o.styles[s] : undefined;
  };
  /** 合并区域在本页上的矩形（只算本页上有的行列） */
  const spanRect = (m: Range) => {
    let x0 = Infinity;
    let x1 = -Infinity;
    let y0 = Infinity;
    let y1 = -Infinity;
    for (let c = m.c0; c <= m.c1; c++) if (x.has(c)) ((x0 = Math.min(x0, x.get(c)!)), (x1 = Math.max(x1, x.get(c)! + layout.colPt(c))));
    for (let r = m.r0; r <= m.r1; r++) if (y.has(r)) ((y0 = Math.min(y0, y.get(r)!)), (y1 = Math.max(y1, y.get(r)! + layout.rowPt(r))));
    return { x0, y0, x1, y1 };
  };
  const firstOnPage = (m: Range, r: number, c: number) => {
    const fr = page.rows.find((rr) => rr >= m.r0 && rr <= m.r1);
    const fc = page.cols.find((cc) => cc >= m.c0 && cc <= m.c1);
    return fr === r && fc === c;
  };

  // 底色
  for (const r of page.rows) {
    for (const c of page.cols) {
      const m = merges.get(cellKey(r, c));
      if (m && !firstOnPage(m, r, c)) continue;
      const fill = st(m ? m.r0 : r, m ? m.c0 : c)?.fill;
      if (!fill) continue;
      const rc = m ? spanRect(m) : { x0: x.get(c)!, y0: y.get(r)!, x1: x.get(c)! + layout.colPt(c), y1: y.get(r)! + layout.rowPt(r) };
      ctx.fillStyle = fill;
      ctx.fillRect(rc.x0, rc.y0, rc.x1 - rc.x0, rc.y1 - rc.y0);
    }
  }
  // 网格线
  if (o.gridLines) {
    ctx.strokeStyle = '#c8c8c8';
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    for (const c of page.cols) (ctx.moveTo(x.get(c)!, 0), ctx.lineTo(x.get(c)!, height));
    ctx.moveTo(width, 0);
    ctx.lineTo(width, height);
    for (const r of page.rows) (ctx.moveTo(0, y.get(r)!), ctx.lineTo(width, y.get(r)!));
    ctx.moveTo(0, height);
    ctx.lineTo(width, height);
    ctx.stroke();
  }
  // 文字
  ctx.textBaseline = 'middle';
  for (const r of page.rows) {
    for (let ci = 0; ci < page.cols.length; ci++) {
      const c = page.cols[ci];
      const m = merges.get(cellKey(r, c));
      if (m && !firstOnPage(m, r, c)) continue;
      const cell = sheet.cells[m ? m.r0 : r]?.[m ? m.c0 : c];
      if (!cell?.t) continue;
      const s = st(m ? m.r0 : r, m ? m.c0 : c);
      const rc = m ? spanRect(m) : { x0: x.get(c)!, y0: y.get(r)!, x1: x.get(c)! + layout.colPt(c), y1: y.get(r)! + layout.rowPt(r) };
      ctx.font = fontOf(s, o.defSize, o.defFont);
      ctx.fillStyle = s?.color ?? '#000';
      const align = s?.align ?? (typeof cell.v === 'number' ? 'right' : typeof cell.v === 'boolean' || cell.k === 'e' ? 'center' : 'left');
      const pad = 2;
      let clipX1 = rc.x1;
      // 左对齐、不换行的文字可以伸进右边的空格子（同 Excel）
      if (!s?.wrap && !m && align === 'left' && ctx.measureText(cell.t).width > rc.x1 - rc.x0 - pad * 2) {
        for (let j = ci + 1; j < page.cols.length; j++) {
          const cc = page.cols[j];
          if (sheet.cells[r]?.[cc] || merges.has(cellKey(r, cc))) break;
          clipX1 = x.get(cc)! + layout.colPt(cc);
          if (ctx.measureText(cell.t).width <= clipX1 - rc.x0 - pad * 2) break;
        }
      }
      ctx.save();
      ctx.beginPath();
      ctx.rect(rc.x0, rc.y0, clipX1 - rc.x0, rc.y1 - rc.y0);
      ctx.clip();
      const size = (s?.fontSize ?? o.defSize) * 1.25;
      const lines = s?.wrap ? wrapLines(ctx, cell.t, rc.x1 - rc.x0 - pad * 2) : [cell.t.replace(/\n/g, ' ')];
      const blockH = lines.length * size;
      const top = s?.valign === 'top' ? rc.y0 + pad : s?.valign === 'center' || !s?.valign ? (rc.y0 + rc.y1) / 2 - blockH / 2 : rc.y1 - pad - blockH;
      ctx.textAlign = align === 'right' ? 'right' : align === 'center' ? 'center' : 'left';
      const tx = align === 'right' ? rc.x1 - pad : align === 'center' ? (rc.x0 + rc.x1) / 2 : rc.x0 + pad;
      lines.forEach((line, i) => {
        const ty = top + size * (i + 0.5);
        ctx.fillText(line, tx, ty);
        if (s?.underline || s?.strike) {
          const tw = ctx.measureText(line).width;
          const lx = align === 'right' ? tx - tw : align === 'center' ? tx - tw / 2 : tx;
          ctx.strokeStyle = ctx.fillStyle as string;
          ctx.lineWidth = 0.6;
          ctx.beginPath();
          const ly = s.underline ? ty + size * 0.38 : ty;
          ctx.moveTo(lx, ly);
          ctx.lineTo(lx + tw, ly);
          ctx.stroke();
        }
      });
      ctx.restore();
    }
  }
  // 边框
  for (const r of page.rows) {
    for (const c of page.cols) {
      const b = cellBorder(st(r, c), r, c, merges.get(cellKey(r, c)));
      if (!b) continue;
      const x0 = x.get(c)!;
      const y0 = y.get(r)!;
      const x1 = x0 + layout.colPt(c);
      const y1 = y0 + layout.rowPt(r);
      const line = (side: NonNullable<typeof b.l>, ax: number, ay: number, bx: number, by: number) => {
        ctx.strokeStyle = side.color ?? '#000';
        ctx.lineWidth = side.w * 0.75;
        ctx.setLineDash(side.style === 'dashed' ? [3, 2] : side.style === 'dotted' ? [1, 1.5] : []);
        ctx.beginPath();
        if (side.style === 'double') {
          const dx = ay === by ? 0 : 1;
          const dy = ay === by ? 1 : 0;
          ctx.lineWidth = 0.6;
          ctx.moveTo(ax - dx, ay - dy);
          ctx.lineTo(bx - dx, by - dy);
          ctx.moveTo(ax + dx, ay + dy);
          ctx.lineTo(bx + dx, by + dy);
        } else {
          ctx.moveTo(ax, ay);
          ctx.lineTo(bx, by);
        }
        ctx.stroke();
      };
      if (b.t) line(b.t, x0, y0, x1, y0);
      if (b.b) line(b.b, x0, y1, x1, y1);
      if (b.l) line(b.l, x0, y0, x0, y1);
      if (b.r) line(b.r, x1, y0, x1, y1);
    }
  }
  ctx.setLineDash([]);
  return canvas;
}

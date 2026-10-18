import { SECTION_BREAK_LABEL, isPictureWatermark, type PageSetup, type RunProps, type Story, type Watermark } from './model';
import { pictureWatermarkBoxPt } from '../../vendor/docx-engine/src/watermark';
import { imageElement } from './images';
import { drawPicture, pictureBounds } from './picture';
import type { DocLayout, EqLayout, FloatBox, Measurer, PageLayout, PlacedLine } from './layout';
import { isCJK } from './text';
import { glyphSize, rowEnds } from './layout';
import { fontMetrics, lineAscentDescent } from './fontfile';
import type { MathBox } from './math';
import { isLineGeom, shapePath } from './shapes';

/** 字体名 → CSS 字体栈（中文字体同时给出英文名，便于不同系统匹配）。 */
const FONT_ALIASES: Record<string, string[]> = {
  宋体: ['宋体', 'SimSun', 'Songti SC', 'Noto Serif CJK SC', 'Source Han Serif SC', 'serif'],
  新宋体: ['新宋体', 'NSimSun', 'SimSun', 'serif'],
  黑体: ['黑体', 'SimHei', 'Heiti SC', 'Noto Sans CJK SC', 'Source Han Sans SC', 'sans-serif'],
  楷体: ['楷体', 'KaiTi', 'STKaiti', 'Kaiti SC', 'serif'],
  仿宋: ['仿宋', 'FangSong', 'STFangsong', 'serif'],
  仿宋_GB2312: ['仿宋_GB2312', 'FangSong_GB2312', '仿宋', 'FangSong', 'serif'],
  楷体_GB2312: ['楷体_GB2312', 'KaiTi_GB2312', '楷体', 'KaiTi', 'serif'],
  方正小标宋简体: ['方正小标宋简体', 'FZXiaoBiaoSong-B05S', '宋体', 'SimSun', 'serif'],
  微软雅黑: ['微软雅黑', 'Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC', 'sans-serif'],
  等线: ['等线', 'DengXian', 'Microsoft YaHei', 'sans-serif'],
  'Times New Roman': ['Times New Roman', 'Liberation Serif', 'Tinos', 'serif'],
  Arial: ['Arial', 'Liberation Sans', 'Arimo', 'sans-serif'],
  Calibri: ['Calibri', 'Carlito', 'sans-serif'],
  Cambria: ['Cambria', 'Caladea', 'serif'],
};

export const EDITOR_FONTS_EA = ['宋体', '黑体', '楷体', '仿宋', '仿宋_GB2312', '楷体_GB2312', '方正小标宋简体', '微软雅黑', '等线'];
export const EDITOR_FONTS_LATIN = ['Times New Roman', 'Arial', 'Calibri', 'Cambria'];

export function fontStack(name: string): string {
  const list = FONT_ALIASES[name] ?? [name, 'serif'];
  return list.map((f) => (/^(serif|sans-serif|monospace)$/.test(f) ? f : `"${f}"`)).join(', ');
}

/** 符号字体：字形放在私用区 U+F0xx（按字符编号 0x20–0xFF 排列），不是普通文字 */
const SYMBOL_FONTS = new Set(['wingdings', 'wingdings 2', 'wingdings 3', 'webdings', 'symbol']);
/** 电脑上没装这个符号字体时，常用符号（方框、勾选、项目符号）改用通用字符显示 */
const SYMBOL_FALLBACK: Record<string, Record<number, string>> = {
  'wingdings 2': { 0xa3: '☐', 0x52: '☑', 0x53: '☒', 0x54: '☒', 0x50: '✓', 0x4f: '✗', 0x51: '☒' },
  wingdings: { 0xa8: '☐', 0x6f: '☐', 0xfe: '☑', 0xfd: '☒', 0xfc: '✓', 0xfb: '✗', 0x6c: '●', 0x6e: '■', 0xa7: '▪', 0x75: '◆', 0x76: '❖', 0xd8: '➢' },
  symbol: { 0xb7: '•' },
};
const fontInstalled = new Map<string, boolean>();

/** 字体是否可用：与两种兜底字体下的宽度比较，都一样说明没有这个字体 */
function isFontInstalled(family: string): boolean {
  let ok = fontInstalled.get(family);
  if (ok === undefined) {
    const ctx = document.createElement('canvas').getContext('2d')!;
    const sample = '\uf052\uf0a3\uf06c\uf020mmwwli';
    const width = (font: string) => {
      ctx.font = `72px ${font}`;
      return ctx.measureText(sample).width;
    };
    ok = ['monospace', 'serif'].some((base) => width(`"${family}", ${base}`) !== width(base));
    fontInstalled.set(family, ok);
  }
  return ok;
}

/**
 * 实际用来测量和绘制的字符与字体。符号字体里的字符（w:sym，或直接用符号字体写的普通字符）
 * 统一换成私用区字符；字体没装时换成通用符号。其它字体原样返回。
 */
export function displayGlyph(ch: string, family: string): { ch: string; family: string } {
  const key = family.toLowerCase();
  if (!SYMBOL_FONTS.has(key)) return { ch, family };
  const c = ch.charCodeAt(0);
  const code = c >= 0xf000 && c <= 0xf0ff ? c - 0xf000 : c < 0x100 ? c : -1;
  if (code < 0x21) return { ch, family };
  if (!isFontInstalled(family)) {
    const sub = SYMBOL_FALLBACK[key]?.[code];
    if (sub) return { ch: sub, family: 'Segoe UI Symbol' };
  }
  return { ch: String.fromCharCode(0xf000 + code), family };
}

const REF = 100;

/** 用 Canvas 测量字宽，结果按字号 100 测量后缓存为比例。 */
export class CanvasMeasurer implements Measurer {
  private ctx = document.createElement('canvas').getContext('2d')!;
  private widths = new Map<string, Map<string, number>>();
  private metricCache = new Map<string, { ascent: number; descent: number }>();

  width(ch0: string, family0: string, bold: boolean, italic: boolean): number {
    const { ch, family } = displayGlyph(ch0, family0);
    const fk = bold ? (italic ? `${family}|bi` : `${family}|b`) : italic ? `${family}|i` : family;
    let m = this.widths.get(fk);
    if (!m) {
      m = new Map();
      this.widths.set(fk, m);
    }
    let w = m.get(ch);
    if (w === undefined) {
      this.ctx.font = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${REF}px ${fontStack(family)}`;
      w = this.ctx.measureText(ch).width / REF;
      m.set(ch, w);
    }
    return w;
  }

  epoch = 0;

  metrics(family: string): { ascent: number; descent: number } {
    // 有字体文件的度量时按 Word 的规则（fontfile.ts），没有时用画布量的近似值
    const fm = fontMetrics(family);
    if (fm) return lineAscentDescent(fm);
    let m = this.metricCache.get(family);
    if (!m) {
      this.ctx.font = `${REF}px ${fontStack(family)}`;
      const t = this.ctx.measureText('国Ag');
      const ascent = (t.fontBoundingBoxAscent ?? t.actualBoundingBoxAscent) / REF;
      const descent = (t.fontBoundingBoxDescent ?? t.actualBoundingBoxDescent) / REF;
      m = { ascent: ascent || 0.88, descent: descent || 0.12 };
      this.metricCache.set(family, m);
    }
    return m;
  }

  /** 字体加载完成后清空缓存 */
  reset() {
    this.widths.clear();
    this.metricCache.clear();
    this.epoch++;
  }
}

export interface SelRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** EQ 域（带圈字符等）：重叠时各部分在字形宽度里按对齐方式叠画，否则依次画；每段用自己的字体、字号和颜色 */
function drawEq(ctx: CanvasRenderingContext2D, e: EqLayout, x: number, w: number, base: number) {
  let cx = x;
  e.parts.forEach((part, i) => {
    let px = e.overstrike ? x + (e.align === 'left' ? 0 : e.align === 'right' ? w - e.widths[i] : (w - e.widths[i]) / 2) : cx;
    for (const seg of part) {
      ctx.fillStyle = seg.props.color === 'auto' ? '#000' : `#${seg.props.color}`;
      for (const c of seg.text) {
        ctx.font = cssFont(seg.props, isCJK(c) ? seg.props.fontEA : seg.props.fontLatin);
        ctx.fillText(c, px, base);
        px += ctx.measureText(c).width;
      }
    }
    cx += e.widths[i];
  });
}

function cssFont(p: RunProps, family: string): string {
  return `${p.italic ? 'italic ' : ''}${p.bold ? 'bold ' : ''}${glyphSize(p)}px ${fontStack(family)}`;
}

/** 修订的显示颜色：按修订人区分（同 Word 的“按作者”） */
const REV_COLORS = ['C00000', '2E75B6', '7030A0', '00804F', 'C55A11', '806000'];
export function revColor(author: string): string {
  let h = 0;
  for (const ch of author) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  return REV_COLORS[h % REV_COLORS.length];
}

/** 公式：按 math.ts 排好的指令画文字和线（分数线、根号上横线等） */
function drawMath(ctx: CanvasRenderingContext2D, box: MathBox, x: number, base: number, color: string) {
  ctx.save();
  ctx.fillStyle = color && color !== 'auto' ? `#${color}` : '#000';
  ctx.strokeStyle = ctx.fillStyle;
  for (const o of box.ops) {
    if (o.t === 'text') {
      ctx.font = `${o.italic ? 'italic ' : ''}${o.bold ? 'bold ' : ''}${o.size}px ${fontStack(o.family)}, "Cambria Math", "STIX Two Math", serif`;
      ctx.fillText(o.s, x + o.x, base + o.y);
    } else {
      ctx.lineWidth = o.w;
      ctx.beginPath();
      ctx.moveTo(x + o.x1, base + o.y1);
      ctx.lineTo(x + o.x2, base + o.y2);
      ctx.stroke();
    }
  }
  ctx.restore();
}

function drawLines(ctx: CanvasRenderingContext2D, lines: PlacedLine[]) {
  ctx.textBaseline = 'alphabetic';
  for (const pl of lines) {
    const base = pl.y + pl.line.baseline;
    let font = '';
    const gs = pl.line.glyphs;
    for (let gi = 0; gi < gs.length; gi++) {
      const g0 = gs[gi];
      const g = g0.image ? g0 : { ...g0, ...displayGlyph(g0.ch, g0.family) };
      if (g.math) {
        drawMath(ctx, g.math, pl.x + g.x, base, g.props.color);
        font = '';
        continue;
      }
      if (g.image) {
        // 嵌入式图片：底边在基线上；还没解码（或浏览器不支持的格式）画一个占位框
        const bounds = pictureBounds(g.image);
        const x = pl.x + g.x + (bounds.w - g.image.w) / 2;
        const y = base - bounds.h + (bounds.h - g.image.h) / 2;
        const el = imageElement(g.image.id);
        if (el) drawPicture(ctx, el, g.image, x, y);
        else {
          ctx.save();
          ctx.strokeStyle = '#9aa4b2';
          ctx.lineWidth = 0.5;
          ctx.strokeRect(x + 0.25, y + 0.25, g.image.w - 0.5, g.image.h - 0.5);
          ctx.restore();
        }
        continue;
      }
      const f = cssFont(g.props, g.family);
      if (f !== font) {
        ctx.font = f;
        font = f;
      }
      const pr = g.props;
      // 突出显示：文字后面的底色
      if (pr.highlight && g.w > 0) {
        ctx.fillStyle = `#${pr.highlight}`;
        ctx.fillRect(pl.x + g.x, base - pr.size * 0.9, g.w + 0.2, pr.size * 1.15);
      }
      const color = pr.rev ? `#${revColor(pr.rev.author)}` : pr.color === 'auto' ? '#000' : `#${pr.color}`;
      ctx.fillStyle = color;
      if (g.leader) {
        // 前导符：按字宽重复，靠右对齐到制表位（与 Word 一样两端留一点空）
        const cw = ctx.measureText(g.leader).width || pr.size * 0.3;
        const n = Math.floor((g.w - cw * 0.5) / cw);
        for (let k = 0; k < n; k++) ctx.fillText(g.leader, pl.x + g.x + g.w - (n - k) * cw, base);
        continue;
      }
      // 上标 / 下标：小一号，上移 / 下移
      const dy = pr.vertAlign === 'superscript' ? -pr.size * 0.33 : pr.vertAlign === 'subscript' ? pr.size * 0.12 : 0;
      if (g.eq) {
        drawEq(ctx, g.eq, pl.x + g.x, g.w, base);
        ctx.fillStyle = color;
        font = '';
      } else ctx.fillText(g.ch, pl.x + g.x, base + dy);
      const thick = Math.max(0.5, pr.size * 0.06);
      // 修订：插入的加下划线，删除的加删除线。后一个字也有同样的线时，线连续画过两字之间的空隙
      // （中西文间距、两端对齐拉开的距离、字符间距；Word 截图：“价的 100%”的下划线不断开）
      const nx = gs[gi + 1];
      const reach = (has: (q: RunProps) => boolean) => (nx && has(nx.props) ? Math.max(g.w, nx.x - g.x) : g.w);
      const ul = (q: RunProps) => !!q.underline || q.rev?.kind === 'ins';
      const st = (q: RunProps) => !!q.strike || q.rev?.kind === 'del';
      if (ul(pr)) ctx.fillRect(pl.x + g.x, base + pr.size * 0.12, reach(ul), thick);
      if (st(pr)) ctx.fillRect(pl.x + g.x, base - pr.size * 0.3 + dy, reach(st), thick);
    }
  }
}

/**
 * 线宽（磅）→ 设备像素：至少 1 像素，零头舍去（Word 屏幕显示：125% 缩放时 1 磅 = 1.67 像素画 1 像素，与 0.5 磅一样细；
 * 0.75 磅在 2 倍缩放下画 1 像素而不是 2 像素），细线才不会比 Word 里显得粗。打印 / PDF 按真实线宽。
 */
export function linePixels(pt: number, scale: number): number {
  return Math.max(1, Math.floor(pt * scale + 0.01));
}

/**
 * 坐标取整到设备像素。先抹掉浮点误差再四舍五入：同一条线由不同算式得到的坐标（如 472.4999999 和 472.5000001）
 * 必须落在同一个像素上，否则公共边会画成两条相邻的线。
 */
export function snapPixel(v: number): number {
  return Math.round(Math.round(v * 1000) / 1000);
}

/** 文本框：填充、边框，里面的文字裁剪在框内 */
function drawFloat(ctx: CanvasRenderingContext2D, f: FloatBox, marks: 'para' | 'all' | null) {
  const b = f.box;
  if (b.image) {
    // 浮动图片：按框的大小画（还没解码时画浅色占位框）
    const el = imageElement(b.image.id);
    if (el) drawPicture(ctx, el, { ...b.image, w: b.w, h: b.h }, f.x, f.y);
    else {
      ctx.save();
      ctx.strokeStyle = '#c8c8c8';
      ctx.strokeRect(f.x, f.y, b.w, b.h);
      ctx.restore();
    }
    return;
  }
  ctx.save();
  if ((b.geom && b.geom !== 'rect' && b.geom !== 'flowChartProcess') || b.rot) {
    // 形状：按外形画填充和轮廓（旋转绕框的中心；翻转只翻外形，文字不翻）
    const cx = f.x + b.w / 2;
    const cy = f.y + b.h / 2;
    ctx.translate(cx, cy);
    if (b.rot) ctx.rotate((b.rot * Math.PI) / 180);
    ctx.save();
    ctx.scale(b.flipH ? -1 : 1, b.flipV ? -1 : 1);
    ctx.translate(-b.w / 2, -b.h / 2);
    ctx.beginPath();
    for (const c of shapePath(b.geom, b.w, b.h)) {
      if (c[0] === 'M') ctx.moveTo(c[1], c[2]);
      else if (c[0] === 'L') ctx.lineTo(c[1], c[2]);
      else if (c[0] === 'C') ctx.bezierCurveTo(c[1], c[2], c[3], c[4], c[5], c[6]);
      else ctx.closePath();
    }
    if (b.fill && !isLineGeom(b.geom)) {
      ctx.fillStyle = `#${b.fill}`;
      ctx.fill();
    }
    if (b.line) {
      ctx.strokeStyle = `#${b.line.color}`;
      ctx.lineWidth = b.line.w;
      ctx.lineJoin = 'miter';
      ctx.stroke();
    }
    ctx.restore();
    ctx.translate(-cx, -cy);
    if (b.paragraphs.length && !isLineGeom(b.geom)) {
      ctx.beginPath();
      ctx.rect(f.x, f.y, b.w, b.h);
      ctx.clip();
      drawLines(ctx, f.lines);
      if (marks) drawFormatMarks(ctx, f.lines, f.x + b.inset.l, f.x + b.w - b.inset.r, false);
    }
    ctx.restore();
    return;
  }
  if (b.fill) {
    ctx.fillStyle = `#${b.fill}`;
    ctx.fillRect(f.x, f.y, b.w, b.h);
  }
  if (b.line) {
    // 边框对齐到设备像素再画（和表格线一样），否则细线落在半个像素上会发虚、显得粗
    const t = ctx.getTransform();
    const lw = linePixels(b.line.w, t.a);
    const half = Math.floor(lw / 2);
    const x1 = snapPixel(f.x * t.a + t.e);
    const y1 = snapPixel(f.y * t.d + t.f);
    const x2 = snapPixel((f.x + b.w) * t.a + t.e);
    const y2 = snapPixel((f.y + b.h) * t.d + t.f);
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.fillStyle = `#${b.line.color}`;
    ctx.fillRect(x1 - half, y1 - half, x2 - x1 + lw, lw);
    ctx.fillRect(x1 - half, y2 - half, x2 - x1 + lw, lw);
    ctx.fillRect(x1 - half, y1 - half, lw, y2 - y1 + lw);
    ctx.fillRect(x2 - half, y1 - half, lw, y2 - y1 + lw);
    ctx.restore();
  }
  ctx.beginPath();
  ctx.rect(f.x, f.y, b.w, b.h);
  ctx.clip();
  drawLines(ctx, f.lines);
  if (marks) drawFormatMarks(ctx, f.lines, f.x + b.inset.l, f.x + b.w - b.inset.r, false);
  ctx.restore();
}

const MARK_COLOR = '#3f7fd8';
/** 段落标记、手动换行符、表格行尾标记：固定的灰色，不跟随文字颜色（Office 是 #808080、按像素对齐；这里线条有抗锯齿，取深一些看起来才一样） */
const RETURN_COLOR = '#606060';
/** 这些标记的高度（磅）：固定大小，不随字号、行高变化（空段落、新起的一行也一样大） */
const RETURN_H = 4.5;

/**
 * 段落标记 / 表格行尾标记：Word 里是贴着基线的小弯箭头（竖线经圆角转向左、实心箭头），
 * 不用字体里的“↵”（那个字形是直角，而且有一个字高，显得太大）。
 */
function drawParaMark(ctx: CanvasRenderingContext2D, x: number, base: number) {
  const h = RETURN_H;
  const head = h * 0.3;
  const w = h * 0.85;
  const left = x + 1;
  const sx = left + w;
  const cy = base - head * 1.1;
  const r = Math.min(h * 0.35, w - head * 1.2);
  ctx.save();
  ctx.fillStyle = RETURN_COLOR;
  ctx.strokeStyle = RETURN_COLOR;
  ctx.lineWidth = 0.75;
  ctx.lineCap = 'round';
  ctx.beginPath();
  ctx.moveTo(sx, base - h);
  ctx.lineTo(sx, cy - r);
  ctx.quadraticCurveTo(sx, cy, sx - r, cy);
  ctx.lineTo(left + head, cy);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(left, cy);
  ctx.lineTo(left + head * 1.4, cy - head);
  ctx.lineTo(left + head * 1.4, cy + head);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

/** 手动换行符：贴着基线的小向下箭头，大小、颜色同段落标记，不用字体里的“↓”（那个字形有一个字高，显得太大） */
function drawLineBreakMark(ctx: CanvasRenderingContext2D, x: number, base: number) {
  const h = RETURN_H;
  const head = h * 0.32;
  const cx = x + 1 + head;
  ctx.save();
  ctx.fillStyle = RETURN_COLOR;
  ctx.strokeStyle = RETURN_COLOR;
  ctx.lineWidth = 0.75;
  ctx.beginPath();
  ctx.moveTo(cx, base - h);
  ctx.lineTo(cx, base - head * 0.6);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(cx - head, base - head * 1.2);
  ctx.lineTo(cx + head, base - head * 1.2);
  ctx.lineTo(cx, base);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

/**
 * 编辑标记：段落标记 ↵、手动换行 ↓；breaks 为真时还画分页符与分节符的虚线
 * （与 Word 一致：默认只显示段落标记，“显示全部编辑标记”才显示分页符 / 分节符）。
 */
function drawFormatMarks(ctx: CanvasRenderingContext2D, lines: PlacedLine[], textLeft: number, textRight: number, breaks: boolean) {
  ctx.fillStyle = MARK_COLOR;
  ctx.strokeStyle = MARK_COLOR;
  ctx.lineWidth = 0.5;
  for (const pl of lines) {
    const { line } = pl;
    const base = pl.y + line.baseline;
    const x = pl.x + line.right + 0.5;
    if (breaks && (pl.sectionBreak || line.brk === 'page' || line.brk === 'column')) {
      const label = pl.sectionBreak ? SECTION_BREAK_LABEL[pl.sectionBreak] : line.brk === 'column' ? '分栏符' : '分页符';
      const y = pl.y + line.height / 2;
      const fs = 7.5;
      ctx.font = `${fs}px sans-serif`;
      const tw = ctx.measureText(label).width;
      const cx = (textLeft + textRight) / 2;
      ctx.setLineDash([1.5, 1.5]);
      ctx.beginPath();
      const segs: [number, number][] = [
        [x + 2, cx - tw / 2 - 3],
        [cx + tw / 2 + 3, textRight],
      ];
      for (const [a, b] of segs) {
        if (b <= a) continue;
        for (const dy of pl.sectionBreak ? [-1, 1] : [0]) {
          ctx.moveTo(a, y + dy);
          ctx.lineTo(b, y + dy);
        }
      }
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillText(label, cx - tw / 2, y + fs / 3);
      if (line.brk !== 'page' && line.brk !== 'column') drawParaMark(ctx, x, base);
    } else if (line.brk === 'end') drawParaMark(ctx, x, base);
    else if (line.brk === 'line') drawLineBreakMark(ctx, x, base);
  }
}

/** 绘制一页。scale 为每磅像素数（含设备像素比）。 */
/** 文字水印：页面正中，斜 45 度（Word 默认）或水平，半透明，大小按 Word 的版框（斜 412×247 pt，平 528×132 pt）撑满 */
export function drawWatermark(ctx: CanvasRenderingContext2D, w: Watermark, s: PageSetup) {
  if (isPictureWatermark(w)) {
    // 图片水印：版心正中，冲蚀时淡化
    const el = imageElement(w.image.id);
    if (!el) return;
    const margin = { widthPt: s.width - s.marginLeft - s.marginRight, heightPt: s.height - s.marginTop - s.marginBottom };
    const box = pictureWatermarkBoxPt({ image: { base64: '', mime: 'image/png', widthPx: w.image.widthPx, heightPx: w.image.heightPx }, ...(w.scale !== undefined ? { scale: w.scale } : {}) }, margin);
    ctx.save();
    ctx.globalAlpha = w.washout === false ? 1 : 0.25;
    ctx.drawImage(el, s.marginLeft + margin.widthPt / 2 - box.widthPt / 2, s.marginTop + margin.heightPt / 2 - box.heightPt / 2, box.widthPt, box.heightPt);
    ctx.restore();
    return;
  }
  if (!w.text) return;
  const diagonal = w.diagonal !== false;
  const boxW = diagonal ? 412 : 528;
  const boxH = diagonal ? 247 : 132;
  const family = `"${w.fontFamily ?? '等线'}", "Microsoft YaHei", "PingFang SC", sans-serif`;
  const font = (size: number) => `${w.italic ? 'italic ' : ''}${w.bold ? 'bold ' : ''}${size}px ${family}`;
  ctx.save();
  ctx.font = font(100);
  const natural = Math.max(1, ctx.measureText(w.text).width);
  const size = Math.max(8, Math.min((boxW / natural) * 100, boxH * 0.9));
  ctx.translate(s.width / 2, s.height / 2);
  if (diagonal) ctx.rotate(-Math.PI / 4);
  ctx.globalAlpha = Math.max(0, Math.min(1, w.opacity ?? 0.5));
  ctx.fillStyle = `#${w.colorHex ?? 'C0C0C0'}`;
  ctx.font = font(size);
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(w.text, 0, 0);
  ctx.restore();
}

export function drawPage(
  ctx: CanvasRenderingContext2D,
  page: PageLayout,
  scale: number,
  /** formatMarks：'para' 只画段落标记和手动换行符，'all' 还画分页符 / 分节符 */
  opts: { story: Story; selection: SelRect[]; highlights?: SelRect[]; showMarks: boolean; formatMarks?: 'para' | 'all' | null; print?: boolean },
) {
  const s = page.setup;
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, s.width, s.height);
  if (page.watermark) drawWatermark(ctx, page.watermark, s);

  // 版心角标（与 Word 相同的“L”形标记）
  if (opts.showMarks) {
    ctx.strokeStyle = '#b8b8b8';
    ctx.lineWidth = 0.6;
    const L = 14;
    const corners: [number, number, number, number][] = [
      [s.marginLeft, s.marginTop, -1, -1],
      [s.width - s.marginRight, s.marginTop, 1, -1],
      [s.marginLeft, s.height - s.marginBottom, -1, 1],
      [s.width - s.marginRight, s.height - s.marginBottom, 1, 1],
    ];
    ctx.beginPath();
    for (const [x, y, dx, dy] of corners) {
      ctx.moveTo(x + dx * L, y);
      ctx.lineTo(x, y);
      ctx.lineTo(x, y + dy * L);
    }
    ctx.stroke();
  }

  const bodyAlpha = opts.print || opts.story === 'body' ? 1 : 0.4;
  const hfAlpha = opts.print || opts.story !== 'body' ? 1 : 0.45;

  // 单元格 / 文本框底纹最先画；批注的范围（浅黄底）和选区（浅蓝底）画在底纹之后、所有文字之前，
  // 否则有底纹（哪怕是白色）的表格里选中的文字看不到颜色
  ctx.globalAlpha = bodyAlpha;
  for (const f of page.fills ?? []) {
    ctx.fillStyle = f.color;
    ctx.fillRect(f.x, f.y, f.w, f.h);
  }
  ctx.globalAlpha = 1;
  ctx.fillStyle = 'rgba(255, 200, 0, 0.28)';
  for (const r of opts.highlights ?? []) ctx.fillRect(r.x, r.y, r.w, r.h);
  ctx.fillStyle = 'rgba(64, 128, 255, 0.28)';
  for (const r of opts.selection) ctx.fillRect(r.x, r.y, r.w, r.h);

  ctx.globalAlpha = hfAlpha;
  drawLines(ctx, page.header);
  drawLines(ctx, page.footer);
  if (opts.formatMarks && !opts.print) {
    drawFormatMarks(ctx, page.header, s.marginLeft, s.width - s.marginRight, opts.formatMarks === 'all');
    drawFormatMarks(ctx, page.footer, s.marginLeft, s.width - s.marginRight, opts.formatMarks === 'all');
  }
  if (opts.story !== 'body') {
    ctx.strokeStyle = '#3f5ef0';
    ctx.setLineDash([3, 2]);
    ctx.lineWidth = 0.6;
    const hy = Math.max(s.marginTop - 4, ...page.header.map((l) => l.y + l.line.height));
    const fy = Math.min(s.height - s.marginBottom + 4, ...page.footer.map((l) => l.y));
    ctx.beginPath();
    ctx.moveTo(0, hy);
    ctx.lineTo(s.width, hy);
    ctx.moveTo(0, fy);
    ctx.lineTo(s.width, fy);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = '#3f5ef0';
    ctx.font = '8px sans-serif';
    ctx.fillText('页眉', 6, hy - 3);
    ctx.fillText('页脚', 6, fy + 10);
  }

  ctx.globalAlpha = bodyAlpha;
  const floatMarks = opts.print ? null : (opts.formatMarks ?? null);
  // 衬于文字下方的文本框先画，其余的画在正文和表格线之后
  for (const f of page.floats) if (f.box.behind || f.hf) drawFloat(ctx, f, floatMarks);
  drawLines(ctx, page.body);
  if (page.notes) drawLines(ctx, page.notes);
  if (opts.formatMarks && !opts.print) drawFormatMarks(ctx, page.body, s.marginLeft, s.width - s.marginRight, opts.formatMarks === 'all');
  if (opts.formatMarks && !opts.print) {
    // 表格行尾标记：画在每行右边框外，与该行第一行文字同一基线，样式同段落标记（Word 显示段落标记时一起显示）
    for (const e of rowEnds(page)) {
      drawParaMark(ctx, e.x + 1, e.baseline);
    }
  }
  // 表格边框：对齐到设备像素再画（至少 1 像素），相邻单元格的公共边只画一次，
  // 否则细线落在半个像素上会发虚变粗、重叠处颜色加深
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.fillStyle = '#000';
  const drawn = new Set<string>();
  for (const b of page.borders) {
    ctx.fillStyle = b.color ?? '#000';
    const lw = linePixels(b.w ?? 0.5, scale);
    const x1 = snapPixel(Math.min(b.x1, b.x2) * scale);
    const y1 = snapPixel(Math.min(b.y1, b.y2) * scale);
    const x2 = snapPixel(Math.max(b.x1, b.x2) * scale);
    const y2 = snapPixel(Math.max(b.y1, b.y2) * scale);
    const key = `${x1},${y1},${x2},${y2},${lw},${b.style ?? ''}`;
    if (drawn.has(key)) continue;
    drawn.add(key);
    const half = Math.floor(lw / 2);
    const horiz = y1 === y2;
    // 两端各多出半个线宽，拐角处接得上
    const span = (from: number, len: number, off: number, t: number) => (horiz ? ctx.fillRect(x1 - half + from, y1 - half + off, len, t) : ctx.fillRect(x1 - half + off, y1 - half + from, t, len));
    const total = (horiz ? x2 - x1 : y2 - y1) + lw;
    if (b.style === 'double') {
      // 双线：两条细线，中间空一条线的宽度
      const t = Math.max(1, Math.round(lw / 3));
      span(0, total, 0, t);
      span(0, total, Math.max(2 * t, lw - t), t);
    } else if (b.style === 'dashed' || b.style === 'dotted') {
      const on = b.style === 'dotted' ? lw : Math.max(3, 3 * lw);
      const off = b.style === 'dotted' ? Math.max(1, lw) : Math.max(2, 2 * lw);
      for (let at = 0; at < total; at += on + off) span(at, Math.min(on, total - at), 0, lw);
    } else span(0, total, 0, lw);
  }
  ctx.restore();
  for (const f of page.floats) if (!f.box.behind && !f.hf) drawFloat(ctx, f, floatMarks);
  ctx.globalAlpha = 1;
}

/** 渲染全部页面为图片（打印用）。 */
export function renderPagesToImages(layout: DocLayout, pxPerPt = 2): string[] {
  return layout.pages.map((page) => {
    const c = document.createElement('canvas');
    c.width = Math.round(page.setup.width * pxPerPt);
    c.height = Math.round(page.setup.height * pxPerPt);
    drawPage(c.getContext('2d')!, page, pxPerPt, { story: 'body', selection: [], showMarks: false, print: true });
    return c.toDataURL('image/png');
  });
}

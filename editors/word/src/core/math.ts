/**
 * 公式（OMML，m:oMath / m:oMathPara）的只读显示（需求 W14 / M9）：原 XML 保存时逐字节保留，这里只负责排版和画出来。
 *
 * OMML → 盒子树 → 用测量器排版成一组绘制指令（文字、线）。支持常用结构：文字 m:r、分数 m:f、上下标 m:sSup / m:sSub /
 * m:sSubSup / m:sPre、根式 m:rad、括号 m:d、大型运算符 m:nary、函数 m:func、重音 m:acc、上下划线 m:bar、
 * 上下限 m:limLow / m:limUpp、矩阵 m:m、方程组 m:eqArr、分组字符 m:groupChr、方框 m:borderBox / m:box；
 * 不认识的元素按顺序排它的内容。
 *
 * 字体按 Word：西文、数字、运算符用 Cambria Math，中文用段落的中文字体；字母和中文默认斜体（m:sty = p 为正体）；
 * 运算符间距按 Word 实测：关系符（= < > ≤ ≥ …）两边各 5/18 em，二元运算符（+ − × ÷ …）两边各 4/18 em。
 */
import type { Measurer } from './layout';

export const MATH_FONT = 'Cambria Math';

/** 绘制指令：文字的 y 是基线，相对公式基线（向下为正） */
export type MathOp =
  | { t: 'text'; x: number; y: number; s: string; size: number; family: string; italic: boolean; bold: boolean }
  | { t: 'line'; x1: number; y1: number; x2: number; y2: number; w: number };

export interface MathBox {
  w: number;
  /** 基线以上 / 以下（磅） */
  asc: number;
  desc: number;
  ops: MathOp[];
}

const RELATIONS = new Set([...'=<>≤≥≈≠≡∼≃≅∝→←↔⇒⇐⇔∈∉⊂⊃⊆⊇∋']);
const BINARY = new Set([...'+−-×÷·±∓∗∘⊕⊗∪∩∧∨']);
const NARY_DEFAULT = '∫';
const BIG = new Set([...'∑∏∐∫∬∭∮⋃⋂⋁⋀']);

interface Ctx {
  m: Measurer;
  /** 中文字符用的字体 */
  ea: string;
  size: number;
  /** 当前文字样式 */
  bold: boolean;
}

const isCJKChar = (c: string) => /[⺀-鿿豈-﫿＀-￯　-〿]/.test(c);
const local = (el: Element) => el.localName;
const kids = (el: Element | null | undefined, name?: string): Element[] =>
  el ? Array.from(el.children).filter((c) => !name || local(c) === name) : [];
const kid = (el: Element | null | undefined, name: string): Element | null => kids(el, name)[0] ?? null;
const val = (el: Element | null | undefined): string | null => el?.getAttributeNS('*', 'val') ?? el?.getAttribute('m:val') ?? el?.getAttribute('val') ?? null;
const prop = (el: Element, pr: string, name: string): string | null => val(kid(kid(el, pr), name));

function empty(): MathBox {
  return { w: 0, asc: 0, desc: 0, ops: [] };
}

/** 把盒子 b 平移 (dx, dy) 后并入 out */
function place(out: MathBox, b: MathBox, dx: number, dy: number) {
  for (const o of b.ops) {
    if (o.t === 'text') out.ops.push({ ...o, x: o.x + dx, y: o.y + dy });
    else out.ops.push({ ...o, x1: o.x1 + dx, x2: o.x2 + dx, y1: o.y1 + dy, y2: o.y2 + dy });
  }
  out.asc = Math.max(out.asc, b.asc - dy);
  out.desc = Math.max(out.desc, b.desc + dy);
}

/** 横向排列 */
function row(boxes: MathBox[]): MathBox {
  const out = empty();
  let x = 0;
  for (const b of boxes) {
    place(out, b, x, 0);
    x += b.w;
  }
  out.w = x;
  return out;
}

/** atStart：在一组元素的开头（开头的 + − 是正负号，不加二元运算符的间距） */
function text(s: string, c: Ctx, italic: boolean, upright = false, atStart = true): MathBox {
  const out = empty();
  let x = 0;
  for (const ch of s) {
    const cjk = isCJKChar(ch);
    const family = cjk ? c.ea : MATH_FONT;
    // 字母（含中文）斜体，数字和符号正体（Word 实测：公式里的“总价”也是斜体）
    const it = !upright && italic && /\p{L}/u.test(ch);
    const w = c.m.width(ch, family, c.bold, it) * c.size;
    // 运算符两边的间距
    const sp = RELATIONS.has(ch) ? (5 / 18) * c.size : BINARY.has(ch) && (x > 0 || !atStart) ? (4 / 18) * c.size : 0;
    x += sp;
    out.ops.push({ t: 'text', x, y: 0, s: ch, size: c.size, family, italic: it, bold: c.bold });
    x += w + sp;
  }
  const fm = c.m.metrics(MATH_FONT);
  out.w = x;
  out.asc = s ? Math.max(0.75, Math.min(fm.ascent, 1)) * c.size : 0;
  out.desc = s ? Math.min(Math.max(0.22, fm.descent), 0.3) * c.size : 0;
  return out;
}

function scaled(c: Ctx, k: number): Ctx {
  return { ...c, size: c.size * k };
}

/** 一组元素顺序排列（m:e、m:num、m:oMath 等容器） */
function seq(el: Element | null, c: Ctx): MathBox {
  return row(kids(el).map((k, i) => node(k, c, i === 0)));
}

function fraction(el: Element, c: Ctx): MathBox {
  const type = prop(el, 'fPr', 'type');
  const num = seq(kid(el, 'num'), c);
  const den = seq(kid(el, 'den'), c);
  if (type === 'lin') return row([num, text('/', c, false, true), den]);
  if (type === 'skw') return row([num, text('⁄', c, false, true), den]);
  const out = empty();
  const pad = 0.12 * c.size;
  const w = Math.max(num.w, den.w) + pad * 2;
  const axis = 0.25 * c.size;
  const gap = 0.15 * c.size;
  const thick = type === 'noBar' ? 0 : Math.max(0.4, 0.05 * c.size);
  place(out, num, (w - num.w) / 2, -(axis + gap + thick / 2 + num.desc));
  place(out, den, (w - den.w) / 2, -axis + gap + thick / 2 + den.asc);
  if (thick) out.ops.push({ t: 'line', x1: pad / 2, x2: w - pad / 2, y1: -axis, y2: -axis, w: thick });
  out.w = w;
  return out;
}

function scripts(base: MathBox, sub: MathBox | null, sup: MathBox | null, c: Ctx, pre = false): MathBox {
  const out = empty();
  const sw = Math.max(sub?.w ?? 0, sup?.w ?? 0);
  const bx = pre ? sw : 0;
  place(out, base, bx, 0);
  const sx = pre ? 0 : base.w + 0.05 * c.size;
  if (sup) place(out, sup, sx + (pre ? sw - sup.w : 0), -Math.max(base.asc - 0.35 * c.size, 0.45 * c.size));
  if (sub) place(out, sub, sx + (pre ? sw - sub.w : 0), Math.max(base.desc, 0.15 * c.size) + 0.1 * c.size);
  out.w = pre ? sw + base.w : sx + sw;
  return out;
}

function radical(el: Element, c: Ctx): MathBox {
  const body = seq(kid(el, 'e'), c);
  const degHide = prop(el, 'radPr', 'degHide');
  const deg = degHide === '1' || degHide === 'on' ? null : seq(kid(el, 'deg'), scaled(c, 0.5));
  const out = empty();
  const gap = 0.1 * c.size;
  const top = body.asc + gap;
  const sign = text('√', { ...c, size: (top + body.desc) * 1.05 }, false, true);
  const dx = deg && deg.w ? Math.max(0, deg.w - sign.w * 0.4) : 0;
  if (deg && deg.w) place(out, deg, 0, -top * 0.45);
  place(out, sign, dx, body.desc * 0.9);
  const bx = dx + sign.w;
  place(out, body, bx, 0);
  const thick = Math.max(0.4, 0.05 * c.size);
  out.ops.push({ t: 'line', x1: bx, x2: bx + body.w + gap, y1: -top, y2: -top, w: thick });
  out.asc = Math.max(out.asc, top + thick);
  out.w = bx + body.w + gap;
  return out;
}

/** 可伸缩的括号：按内容高度放大 */
function fence(ch: string, inner: MathBox, c: Ctx): MathBox {
  if (!ch) return empty();
  const h = Math.max(inner.asc + inner.desc, c.size);
  const k = Math.max(1, h / c.size);
  const b = text(ch, { ...c, size: c.size * k }, false, true);
  // 括号中线对齐内容中线
  const mid = (inner.asc - inner.desc) / 2;
  const bmid = (b.asc - b.desc) / 2;
  const out = empty();
  place(out, b, 0, bmid - mid);
  out.w = b.w;
  return out;
}

function delim(el: Element, c: Ctx): MathBox {
  const pr = kid(el, 'dPr');
  const beg = val(kid(pr, 'begChr')) ?? '(';
  const end = val(kid(pr, 'endChr')) ?? ')';
  const sep = val(kid(pr, 'sepChr')) ?? '|';
  const parts: MathBox[] = [];
  kids(el, 'e').forEach((e, i) => {
    if (i > 0) parts.push(text(sep, c, false, true));
    parts.push(seq(e, c));
  });
  const inner = row(parts);
  return row([fence(beg, inner, c), inner, fence(end, inner, c)]);
}

function nary(el: Element, c: Ctx): MathBox {
  const pr = kid(el, 'naryPr');
  const chr = val(kid(pr, 'chr')) ?? NARY_DEFAULT;
  const loc = val(kid(pr, 'limLoc')) ?? (chr === NARY_DEFAULT ? 'subSup' : 'undOvr');
  const hide = (n: string) => ['1', 'on', 'true'].includes(val(kid(pr, n)) ?? '');
  const op = text(chr, scaled(c, BIG.has(chr) ? 1.4 : 1), false, true);
  const sub = hide('subHide') ? null : seq(kid(el, 'sub'), scaled(c, 0.7));
  const sup = hide('supHide') ? null : seq(kid(el, 'sup'), scaled(c, 0.7));
  const body = seq(kid(el, 'e'), c);
  let opBox: MathBox;
  if (loc === 'undOvr') {
    opBox = empty();
    const w = Math.max(op.w, sub?.w ?? 0, sup?.w ?? 0);
    place(opBox, op, (w - op.w) / 2, 0);
    if (sup) place(opBox, sup, (w - sup.w) / 2, -(op.asc + sup.desc + 0.05 * c.size));
    if (sub) place(opBox, sub, (w - sub.w) / 2, op.desc + sub.asc + 0.05 * c.size);
    opBox.w = w;
  } else opBox = scripts(op, sub, sup, c);
  return row([opBox, text(' ', c, false, true), body]);
}

function accent(el: Element, c: Ctx): MathBox {
  const chr = val(kid(kid(el, 'accPr'), 'chr')) ?? '̂';
  const body = seq(kid(el, 'e'), c);
  const out = empty();
  place(out, body, 0, 0);
  // 组合附加符号单独画在内容上方
  const mark = text(chr.replace(/[̀-ͯ]/, (m) => ({ '̂': 'ˆ', '̃': '˜', '̄': '¯', '̇': '˙', '̈': '¨', '⃗': '→' })[m] ?? m), c, false, true);
  place(out, mark, (body.w - mark.w) / 2, -(body.asc - 0.35 * c.size));
  out.w = Math.max(body.w, mark.w);
  return out;
}

function bar(el: Element, c: Ctx): MathBox {
  const top = (prop(el, 'barPr', 'pos') ?? 'bot') === 'top';
  const body = seq(kid(el, 'e'), c);
  const out = empty();
  place(out, body, 0, 0);
  const y = top ? -(body.asc + 0.08 * c.size) : body.desc + 0.08 * c.size;
  const w = Math.max(0.4, 0.05 * c.size);
  out.ops.push({ t: 'line', x1: 0, x2: body.w, y1: y, y2: y, w });
  if (top) out.asc = Math.max(out.asc, -y + w);
  else out.desc = Math.max(out.desc, y + w);
  out.w = body.w;
  return out;
}

function limit(el: Element, c: Ctx, low: boolean): MathBox {
  const body = seq(kid(el, 'e'), c);
  const lim = seq(kid(el, 'lim'), scaled(c, 0.7));
  const out = empty();
  const w = Math.max(body.w, lim.w);
  place(out, body, (w - body.w) / 2, 0);
  place(out, lim, (w - lim.w) / 2, low ? body.desc + lim.asc + 0.05 * c.size : -(body.asc + lim.desc + 0.05 * c.size));
  out.w = w;
  return out;
}

/** 多行（方程组、矩阵的每一行）上下排列，整体中线对齐公式轴 */
function stack(rows: MathBox[][], c: Ctx, colGap: number): MathBox {
  const ncol = Math.max(0, ...rows.map((r) => r.length));
  const colW = Array.from({ length: ncol }, (_, i) => Math.max(0, ...rows.map((r) => r[i]?.w ?? 0)));
  const gap = 0.25 * c.size;
  const out = empty();
  let y = 0;
  const placed: { b: MathBox; x: number; y: number }[] = [];
  rows.forEach((r, ri) => {
    const asc = Math.max(0.7 * c.size, ...r.map((b) => b.asc));
    const desc = Math.max(0.25 * c.size, ...r.map((b) => b.desc));
    if (ri > 0) y += gap;
    y += asc;
    let x = 0;
    r.forEach((b, ci) => {
      placed.push({ b, x: x + (ncol > 1 ? (colW[ci] - b.w) / 2 : 0), y });
      x += colW[ci] + colGap;
    });
    y += desc;
  });
  const shift = y / 2 + 0.25 * c.size; // 整体中心放在公式轴上
  for (const p of placed) place(out, p.b, p.x, p.y - shift);
  out.w = colW.reduce((a, b) => a + b, 0) + colGap * Math.max(0, ncol - 1);
  out.asc = Math.max(out.asc, shift);
  out.desc = Math.max(out.desc, y - shift);
  return out;
}

function node(el: Element, c: Ctx, atStart = true): MathBox {
  switch (local(el)) {
    case 'r': {
      const sty = prop(el, 'rPr', 'sty');
      const nor = ['1', 'on', 'true'].includes(prop(el, 'rPr', 'nor') ?? '');
      const s = kids(el, 't').map((t) => t.textContent ?? '').join('');
      const bold = sty === 'b' || sty === 'bi';
      return text(s, { ...c, bold: c.bold || bold }, !nor && sty !== 'p' && sty !== 'b', false, atStart);
    }
    case 'f':
      return fraction(el, c);
    case 'sSup':
      return scripts(seq(kid(el, 'e'), c), null, seq(kid(el, 'sup'), scaled(c, 0.7)), c);
    case 'sSub':
      return scripts(seq(kid(el, 'e'), c), seq(kid(el, 'sub'), scaled(c, 0.7)), null, c);
    case 'sSubSup':
      return scripts(seq(kid(el, 'e'), c), seq(kid(el, 'sub'), scaled(c, 0.7)), seq(kid(el, 'sup'), scaled(c, 0.7)), c);
    case 'sPre':
      return scripts(seq(kid(el, 'e'), c), seq(kid(el, 'sub'), scaled(c, 0.7)), seq(kid(el, 'sup'), scaled(c, 0.7)), c, true);
    case 'rad':
      return radical(el, c);
    case 'd':
      return delim(el, c);
    case 'nary':
      return nary(el, c);
    case 'func':
      return row([seq(kid(el, 'fName'), c), text(' ', c, false, true), seq(kid(el, 'e'), c)]);
    case 'acc':
      return accent(el, c);
    case 'bar':
      return bar(el, c);
    case 'limLow':
      return limit(el, c, true);
    case 'limUpp':
      return limit(el, c, false);
    case 'groupChr': {
      const chr = val(kid(kid(el, 'groupChrPr'), 'chr')) ?? '⏟';
      const top = val(kid(kid(el, 'groupChrPr'), 'pos')) === 'top';
      const body = seq(kid(el, 'e'), c);
      const out = empty();
      place(out, body, 0, 0);
      const g = text(chr, c, false, true);
      place(out, g, (body.w - g.w) / 2, top ? -(body.asc + g.desc) : body.desc + g.asc);
      out.w = Math.max(body.w, g.w);
      return out;
    }
    case 'eqArr':
      return stack(kids(el, 'e').map((e) => [seq(e, c)]), c, 0);
    case 'm':
      return stack(kids(el, 'mr').map((r) => kids(r, 'e').map((e) => seq(e, c))), c, 0.6 * c.size);
    case 'borderBox': {
      const body = seq(kid(el, 'e'), c);
      const out = empty();
      const p = 0.1 * c.size;
      place(out, body, p, 0);
      const t = -(body.asc + p);
      const b = body.desc + p;
      const w = body.w + 2 * p;
      for (const [x1, y1, x2, y2] of [[0, t, w, t], [0, b, w, b], [0, t, 0, b], [w, t, w, b]]) out.ops.push({ t: 'line', x1, y1, x2, y2, w: 0.5 });
      out.asc = -t;
      out.desc = b;
      out.w = w;
      return out;
    }
    case 'oMathPara': {
      // 多个公式各占一行
      const parts = kids(el, 'oMath');
      return parts.length > 1 ? stack(parts.map((p) => [seq(p, c)]), c, 0) : seq(parts[0] ?? el, c);
    }
    case 'rPr':
    case 'ctrlPr':
    case 'oMathParaPr':
    case 'fPr':
    case 'sSupPr':
    case 'sSubPr':
    case 'sSubSupPr':
    case 'radPr':
    case 'dPr':
    case 'naryPr':
    case 'funcPr':
    case 'accPr':
    case 'barPr':
    case 'limLowPr':
    case 'limUppPr':
    case 'mPr':
    case 'eqArrPr':
    case 'groupChrPr':
    case 'boxPr':
    case 'borderBoxPr':
    case 'bookmarkStart':
    case 'bookmarkEnd':
    case 'proofErr':
      return empty();
    default:
      // m:oMath、m:e、m:box 以及不认识的元素：排里面的内容；w:r 里的普通文字也显示
      if (local(el) === 't') return text(el.textContent ?? '', c, false, true);
      return seq(el, c);
  }
}

const cache = new Map<string, MathBox>();

/** 公式的排版结果（按 XML、字号、测量器缓存）；解析失败时返回 null（按占位框显示） */
export function layoutMath(xml: string, size: number, ea: string, m: Measurer): MathBox | null {
  const key = `${size}|${ea}|${m.epoch ?? 0}|${xml}`;
  const hit = cache.get(key);
  if (hit) return hit;
  if (typeof DOMParser === 'undefined') return null;
  // 原样保存的 XML 片段不带命名空间声明，包一层再解析
  const doc = new DOMParser().parseFromString(
    `<x xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</x>`,
    'application/xml',
  );
  const root = doc.documentElement.firstElementChild;
  if (!root || doc.getElementsByTagName('parsererror').length) return null;
  const box = node(root, { m, ea, size, bold: false });
  if (cache.size > 500) cache.clear();
  cache.set(key, box);
  return box;
}

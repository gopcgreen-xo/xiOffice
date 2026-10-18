/**
 * 带格式的复制粘贴：文档片段（段落、表格）↔ HTML。
 * 复制时写 HTML（Word、WPS、邮件里粘贴保留字体、字号、粗斜体、颜色、对齐、表格）；
 * 粘贴时读 Word / WPS / 网页 / Excel 放进剪贴板的 HTML：只认行内样式和 b / i / u / font 等标签，
 * 段落（p、div、h1–h6、li）、换行（br）、表格（含 colspan / rowspan）；不认识的标签取其中的文字。
 */
import { FIELD_CHAR, isAtom, newParagraph, type Block, type Paragraph, type Run, type RunProps, type Table, type Align, type ParaProps } from '../core/model';
import { isCJK } from '../core/text';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// ---------- 片段 → HTML ----------

function runCss(p: RunProps): string {
  const css = [`font-family:'${p.fontLatin}','${p.fontEA}'`, `font-size:${p.size}pt`];
  if (p.bold) css.push('font-weight:bold');
  if (p.italic) css.push('font-style:italic');
  if (p.underline) css.push('text-decoration:underline');
  if (p.color && p.color !== 'auto') css.push(`color:#${p.color}`);
  return css.join(';');
}

const ALIGN_CSS: Record<Align, string> = { left: 'left', center: 'center', right: 'right', both: 'justify', distribute: 'justify' };

function paraHtml(p: Paragraph): string {
  const css = ['margin:0', `text-align:${ALIGN_CSS[p.props.align]}`];
  if (p.props.indentLeft) css.push(`margin-left:${p.props.indentLeft}pt`);
  const first = p.props.firstLineChars !== null ? p.props.firstLineChars * p.mark.size : p.props.firstLine;
  if (first) css.push(`text-indent:${first}pt`);
  const body = p.runs
    .filter((r) => !isAtom(r))
    .map((r) => {
      const text = esc(r.text.replace(/\f/g, '').replace(new RegExp(FIELD_CHAR, 'g'), '')).replace(/\t/g, '&emsp;').replace(/\n/g, '<br>');
      return text ? `<span style="${runCss(r.props)}">${text}</span>` : '';
    })
    .join('');
  return `<p style="${css.join(';')}">${body || '&nbsp;'}</p>`;
}

function tableHtml(t: Table): string {
  const rows = t.rows.map((row, ri) => {
    let col = 0;
    const cells = row.cells.map((c) => {
      const at = col;
      col += c.gridSpan;
      if (c.vMerge === 'continue') return '';
      let span = 1;
      if (c.vMerge === 'restart') {
        for (let r = ri + 1; r < t.rows.length; r++) {
          let k = 0;
          const below = t.rows[r].cells.find((x) => {
            const hit = k === at;
            k += x.gridSpan;
            return hit;
          });
          if (below?.vMerge !== 'continue') break;
          span++;
        }
      }
      const w = t.cols.slice(at, at + c.gridSpan).reduce((a, b) => a + b, 0);
      const attrs = `${c.gridSpan > 1 ? ` colspan="${c.gridSpan}"` : ''}${span > 1 ? ` rowspan="${span}"` : ''}`;
      return `<td${attrs} style="width:${w.toFixed(1)}pt;border:${t.borders ? '0.5pt solid #000' : 'none'};padding:0 5.4pt;vertical-align:${c.vAlign === 'center' ? 'middle' : c.vAlign}">${c.paragraphs.map(paraHtml).join('')}</td>`;
    });
    return `<tr>${cells.join('')}</tr>`;
  });
  return `<table style="border-collapse:collapse" cellspacing="0" cellpadding="0">${rows.join('')}</table>`;
}

export function fragmentToHtml(blocks: Block[]): string {
  return `<html><body><!--StartFragment-->${blocks.map((b) => (b.kind === 'p' ? paraHtml(b) : tableHtml(b))).join('')}<!--EndFragment--></body></html>`;
}

/** 片段的纯文本（段落之间换行，单元格之间制表符） */
export function fragmentToText(blocks: Block[]): string {
  const para = (p: Paragraph) => p.runs.filter((r) => !isAtom(r)).map((r) => r.text).join('').replace(/\f/g, '');
  return blocks
    .map((b) => (b.kind === 'p' ? para(b) : b.rows.map((r) => r.cells.filter((c) => c.vMerge !== 'continue').map((c) => c.paragraphs.map(para).join(' ')).join('\t')).join('\r\n')))
    .join('\r\n');
}

// ---------- HTML → 片段 ----------

const NAMED: Record<string, string> = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', navy: '000080', maroon: '800000', purple: '800080', orange: 'FFA500' };

/** CSS 颜色 → 'RRGGBB'；windowtext、黑色按“自动”处理 */
export function cssColor(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const s = v.trim().toLowerCase();
  if (s === 'windowtext' || s === 'auto' || s === 'inherit' || s === 'initial') return 'auto';
  let hex: string | undefined;
  const m6 = s.match(/^#([0-9a-f]{6})$/);
  const m3 = s.match(/^#([0-9a-f]{3})$/);
  const rgb = s.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (m6) hex = m6[1];
  else if (m3) hex = [...m3[1]].map((c) => c + c).join('');
  else if (rgb) hex = [rgb[1], rgb[2], rgb[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('');
  else hex = NAMED[s]?.toLowerCase();
  if (!hex) return undefined;
  return hex === '000000' ? 'auto' : hex.toUpperCase();
}

/** CSS 长度 → 磅（pt、px、cm、mm、in、em 相对 base） */
export function cssLength(v: string | undefined, base = 10.5): number | undefined {
  if (!v) return undefined;
  const m = v.trim().match(/^(-?[\d.]+)\s*(pt|px|cm|mm|in|em|rem|%)?$/i);
  if (!m) return undefined;
  const n = Number(m[1]);
  switch ((m[2] ?? 'px').toLowerCase()) {
    case 'pt':
      return n;
    case 'px':
      return n * 0.75;
    case 'cm':
      return (n * 72) / 2.54;
    case 'mm':
      return (n * 72) / 25.4;
    case 'in':
      return n * 72;
    case 'em':
    case 'rem':
      return n * base;
    case '%':
      return (n / 100) * base;
  }
  return undefined;
}

/** 正在解析的是 Excel 放进剪贴板的 HTML（表格默认画边框，同 Excel 的网格） */
let fromExcel = false;

/** <font size="1..7"> → 磅 */
const FONT_SIZES = [7.5, 10, 12, 13.5, 18, 24, 36];

function parseStyle(attr: string | null): Record<string, string> {
  const out: Record<string, string> = {};
  if (!attr) return out;
  for (const part of attr.split(';')) {
    const i = part.indexOf(':');
    if (i < 0) continue;
    out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

const firstFamily = (v: string) => v.split(',')[0].trim().replace(/^['"]|['"]$/g, '');
const hasCJK = (s: string) => [...s].some((ch) => isCJK(ch));

/** 元素的行内样式和标签 → 字符格式 */
function applyRunStyle(el: Element, p: RunProps): RunProps {
  const out = { ...p };
  const tag = el.tagName.toLowerCase();
  if (tag === 'b' || tag === 'strong') out.bold = true;
  if (tag === 'i' || tag === 'em') out.italic = true;
  if (tag === 'u' || tag === 'ins') out.underline = true;
  if (/^h[1-6]$/.test(tag)) out.bold = true;
  if (tag === 'font') {
    const face = el.getAttribute('face');
    if (face) {
      const f = firstFamily(face);
      if (hasCJK(f)) out.fontEA = f;
      else out.fontLatin = f;
    }
    const color = cssColor(el.getAttribute('color') ?? undefined);
    if (color) out.color = color;
    const size = Number(el.getAttribute('size'));
    if (size >= 1 && size <= 7) out.size = FONT_SIZES[size - 1];
  }
  const st = parseStyle(el.getAttribute('style'));
  // Word 的 mso-fareast-font-family 是中文字体
  const ea = st['mso-fareast-font-family'];
  if (ea) out.fontEA = firstFamily(ea);
  const ff = st['font-family'];
  if (ff) {
    const families = ff.split(',').map(firstFamily).filter(Boolean);
    const cjk = families.find(hasCJK);
    const latin = families.find((f) => !hasCJK(f) && !/^(serif|sans-serif|monospace|宋体|黑体)$/i.test(f));
    if (cjk && !ea) out.fontEA = cjk;
    if (latin) out.fontLatin = latin;
  }
  const size = cssLength(st['font-size'], p.size);
  if (size && size > 1 && size < 200) out.size = Math.round(size * 2) / 2;
  const w = st['font-weight'];
  if (w) out.bold = w === 'bold' || w === 'bolder' || Number(w) >= 600;
  const fs = st['font-style'];
  if (fs) out.italic = fs === 'italic' || fs === 'oblique';
  const td = st['text-decoration'] ?? st['text-decoration-line'] ?? st['text-underline'];
  if (td) out.underline = /underline|single/.test(td);
  const color = cssColor(st.color);
  if (color) out.color = color;
  return out;
}

function applyParaStyle(el: Element, p: Partial<ParaProps>, size: number): Partial<ParaProps> {
  const out = { ...p };
  const st = parseStyle(el.getAttribute('style'));
  const align = (st['text-align'] ?? el.getAttribute('align') ?? '').toLowerCase();
  if (align === 'center') out.align = 'center';
  else if (align === 'right') out.align = 'right';
  else if (align === 'justify') out.align = 'both';
  else if (align === 'left') out.align = 'left';
  const indent = cssLength(st['text-indent'], size);
  if (indent !== undefined) {
    out.firstLine = indent;
    out.firstLineChars = null;
  }
  const left = cssLength(st['margin-left'] ?? st['padding-left'], size);
  if (left !== undefined && left >= 0 && left < 300) out.indentLeft = left;
  return out;
}

const BLOCK = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'blockquote', 'pre', 'section', 'article', 'header', 'footer', 'dt', 'dd']);
const SKIP = new Set(['style', 'script', 'head', 'title', 'meta', 'link', 'xml', 'noscript', 'template']);

interface Builder {
  blocks: Block[];
  cur: Paragraph | null;
}

function flush(b: Builder) {
  if (b.cur) {
    // 去掉段首段尾由 HTML 换行带来的空白
    const runs = b.cur.runs;
    if (runs.length) {
      runs[0].text = runs[0].text.replace(/^ +/, '');
      runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/ +$/, '');
    }
    b.cur.runs = runs.filter((r) => r.text);
    b.blocks.push(b.cur);
  }
  b.cur = null;
}

function startPara(b: Builder, para: Partial<ParaProps>, run: RunProps) {
  flush(b);
  b.cur = newParagraph('', { firstLineChars: null, ...para }, run);
}

function addText(b: Builder, text: string, run: RunProps, para: Partial<ParaProps>) {
  if (!b.cur) {
    if (!text.trim()) return;
    startPara(b, para, run);
  }
  const p = b.cur!;
  const last = p.runs[p.runs.length - 1] as Run | undefined;
  if (last && !isAtom(last) && JSON.stringify(last.props) === JSON.stringify(run)) last.text += text;
  else p.runs.push({ text, props: { ...run } });
}

function walk(node: Node, b: Builder, run: RunProps, para: Partial<ParaProps>, ctx: { pre: boolean; width: number; list: { ordered: boolean; n: number } | null }) {
  if (node.nodeType === 3) {
    let t = node.textContent ?? '';
    // 普通空白合并成一个空格；&nbsp; 保留为空格（Word 用它做段首缩进和空段落）
    if (!ctx.pre) t = t.replace(/[ \t\r\n\f]+/g, ' ').replace(/\u00a0/g, ' ');
    if (t) addText(b, t, run, para);
    return;
  }
  if (node.nodeType !== 1) return;
  const el = node as Element;
  const tag = el.tagName.toLowerCase().replace(/^[a-z]+:/, (m) => (m === 'o:' || m === 'v:' || m === 'w:' ? '#' : m));
  if (SKIP.has(tag) || tag.startsWith('#')) return;
  const st = parseStyle(el.getAttribute('style'));
  if (st.display === 'none' || st['mso-hide'] === 'all') return;
  const r = applyRunStyle(el, run);
  if (tag === 'br') {
    addText(b, '\n', r, para);
    return;
  }
  if (tag === 'table') {
    flush(b);
    const t = parseTable(el, r, ctx.width);
    if (t) b.blocks.push(t);
    return;
  }
  if (tag === 'ul' || tag === 'ol') {
    flush(b);
    const list = { ordered: tag === 'ol', n: 0 };
    el.childNodes.forEach((c) => walk(c, b, r, para, { ...ctx, list }));
    flush(b);
    return;
  }
  if (BLOCK.has(tag)) {
    const pp = applyParaStyle(el, para, r.size);
    startPara(b, pp, r);
    if (tag === 'li' && ctx.list) {
      ctx.list.n++;
      addText(b, ctx.list.ordered ? `${ctx.list.n}. ` : '• ', r, pp);
    }
    el.childNodes.forEach((c) => walk(c, b, r, pp, { ...ctx, pre: ctx.pre || tag === 'pre' }));
    // 空段落（Word 的 <p>&nbsp;</p>）也要保留
    if (!b.cur) startPara(b, pp, r);
    flush(b);
    return;
  }
  el.childNodes.forEach((c) => walk(c, b, r, para, ctx));
}

/** <table> → 表格：colspan → gridSpan，rowspan → 纵向合并；列宽按单元格宽度，没写时平分 */
function parseTable(el: Element, run: RunProps, width: number): Table | null {
  const trs = [...el.querySelectorAll('tr')].filter((tr) => tr.closest('table') === el);
  if (!trs.length) return null;
  // 先排出网格：每格占的行列
  type Slot = { el: Element; r: number; c: number; cs: number; rs: number };
  const slots: Slot[] = [];
  const taken = new Set<string>();
  let ncols = 0;
  trs.forEach((tr, r) => {
    let c = 0;
    [...tr.children].filter((x) => /^(td|th)$/i.test(x.tagName)).forEach((td) => {
      while (taken.has(`${r}:${c}`)) c++;
      const cs = Math.max(1, Math.min(63, Number(td.getAttribute('colspan')) || 1));
      const rs = Math.max(1, Math.min(trs.length - r, Number(td.getAttribute('rowspan')) || 1));
      slots.push({ el: td, r, c, cs, rs });
      for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) taken.add(`${r + i}:${c + j}`);
      c += cs;
      ncols = Math.max(ncols, c);
    });
  });
  if (!ncols) return null;
  // 列宽：单格（不跨列）里写的宽度
  const widths: (number | undefined)[] = Array(ncols).fill(undefined);
  for (const s of slots) {
    if (s.cs !== 1 || widths[s.c] !== undefined) continue;
    const st = parseStyle(s.el.getAttribute('style'));
    const w = cssLength(st.width ?? s.el.getAttribute('width') ?? undefined, run.size);
    if (w && w > 4) widths[s.c] = w;
  }
  const known = widths.filter((w): w is number => w !== undefined);
  const rest = width - known.reduce((a, b) => a + b, 0);
  const fill = known.length < ncols ? Math.max(20, rest / (ncols - known.length)) : 0;
  let cols = widths.map((w) => w ?? fill);
  const total = cols.reduce((a, b) => a + b, 0);
  // 比版心宽时按比例缩小
  if (total > width) cols = cols.map((w) => (w * width) / total);
  const rows: Table['rows'] = trs.map(() => ({ height: 0, heightRule: 'auto', header: false, cells: [] }));
  const byPos = new Map(slots.map((s) => [`${s.r}:${s.c}`, s]));
  trs.forEach((_, r) => {
    for (let c = 0; c < ncols; ) {
      const own = byPos.get(`${r}:${c}`);
      if (own) {
        const b: Builder = { blocks: [], cur: null };
        const st = parseStyle(own.el.getAttribute('style'));
        const va = (st['vertical-align'] ?? own.el.getAttribute('valign') ?? '').toLowerCase();
        own.el.childNodes.forEach((n) => walk(n, b, applyRunStyle(own.el, run), applyParaStyle(own.el, { align: 'left' }, run.size), { pre: false, width: 200, list: null }));
        flush(b);
        const paragraphs = b.blocks.flatMap((x) => (x.kind === 'p' ? [x] : []));
        rows[r].cells.push({
          gridSpan: own.cs,
          vMerge: own.rs > 1 ? 'restart' : 'none',
          vAlign: va === 'middle' || va === 'center' ? 'center' : va === 'bottom' ? 'bottom' : 'top',
          paragraphs: paragraphs.length ? paragraphs : [newParagraph('', { align: 'left', firstLineChars: null }, run)],
        });
        c += own.cs;
        continue;
      }
      // 被上面的 rowspan 盖住的格子：纵向合并的延续
      const above = slots.find((s) => s.c === c && s.r < r && s.r + s.rs > r);
      if (above) {
        rows[r].cells.push({ gridSpan: above.cs, vMerge: 'continue', vAlign: 'top', paragraphs: [newParagraph('', { align: 'left', firstLineChars: null }, run)] });
        c += above.cs;
        continue;
      }
      // 这一行比别的行短：补空格子
      rows[r].cells.push({ gridSpan: 1, vMerge: 'none', vAlign: 'top', paragraphs: [newParagraph('', { align: 'left', firstLineChars: null }, run)] });
      c++;
    }
  });
  const border = el.getAttribute('border');
  const st = parseStyle(el.getAttribute('style'));
  const firstTd = parseStyle(slots[0]?.el.getAttribute('style'));
  const borders = (border !== null && border !== '0') || /solid|windowtext/.test(`${st.border ?? ''}${firstTd.border ?? ''}${firstTd['border-top'] ?? ''}${firstTd['border-left'] ?? ''}`);
  return {
    kind: 'table',
    cols,
    rows,
    align: 'center',
    indent: 0,
    borders: borders || fromExcel,
    cellMargin: { left: 5.4, right: 5.4, top: 0, bottom: 0 },
  };
}

/**
 * HTML → 片段。run = 光标处的格式（没写样式的文字用它），width = 版心宽（磅，表格列宽用）。
 * 解析不出内容时返回 null（调用方改用纯文本）。
 */
export function htmlToFragment(html: string, run: RunProps, width: number, doc?: Document): Block[] | null {
  // 只取 StartFragment 和 EndFragment 之间（Word、Excel、浏览器都会写这两个标记）
  const m = html.match(/<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/);
  const body = m ? m[1] : html;
  const d = doc ?? new DOMParser().parseFromString(`<body>${body}</body>`, 'text/html');
  if (doc) d.body.innerHTML = body;
  fromExcel = /urn:schemas-microsoft-com:office:excel|ProgId"?\s*content="?Excel/i.test(html);
  const b: Builder = { blocks: [], cur: null };
  d.body.childNodes.forEach((n) => walk(n, b, run, { align: 'left' }, { pre: false, width, list: null }));
  flush(b);
  // 去掉末尾的空段落（选到段落标记时浏览器常多带一个）
  while (b.blocks.length > 1) {
    const last = b.blocks[b.blocks.length - 1];
    if (last.kind === 'p' && !last.runs.length) b.blocks.pop();
    else break;
  }
  return b.blocks.length ? b.blocks : null;
}

/**
 * styles.xml 的可编辑表示：数字格式、字体、填充、边框、单元格格式（cellXfs）。
 * 读入时保留每个条目的原始 XML，编辑只**追加**新条目（原有条目不改不删），保存时把这几段写回原文件，其它部分原样保留。
 * 全是普通数据（可以在 Worker 和界面之间传递、可以 structuredClone 进撤销记录）。
 */
import { unescapeXml, xmlAttr } from '@wz/editor-shared/xlsx';
import { builtinFormat } from './numfmt';
import { readPalette, resolveBorder, resolveColor, resolveFill, type CellBorder, type Palette } from './color';

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** 字体名、字号（磅）；与默认字体相同时也会写 */
  fontName?: string;
  fontSize?: number;
  align?: 'left' | 'center' | 'right';
  valign?: 'top' | 'center' | 'bottom';
  wrap?: boolean;
  /** 底色（CSS #rrggbb），无填充时没有 */
  fill?: string;
  /** 字体颜色（CSS #rrggbb），自动色时没有 */
  color?: string;
  /** 边框 */
  border?: CellBorder;
}

export interface Xf {
  /** 开始标签上的属性（保持原顺序） */
  attrs: Record<string, string>;
  /** 子元素（alignment / protection / extLst）的原始 XML */
  inner: string;
}

export interface StyleBook {
  /** 自定义数字格式 id → 格式代码 */
  numFmts: Record<number, string>;
  fonts: string[];
  fills: string[];
  borders: string[];
  xfs: Xf[];
  palette: Palette;
  /** 条件格式、按颜色筛选用的差异格式（dxfs） */
  dxfs?: CellStyle[];
  /** 新追加的差异格式（原始 XML），保存时接在 dxfs 后面 */
  newDxfs?: string[];
  /** 有追加的条目，保存时要写回 styles.xml */
  dirty?: boolean;
}

/** 修改格式：undefined = 不变；null = 清除（颜色、底色） */
export interface StyleChange {
  numFmt?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontName?: string;
  fontSize?: number;
  color?: string | null;
  fill?: string | null;
  align?: 'left' | 'center' | 'right' | null;
  valign?: 'top' | 'center' | 'bottom' | null;
  wrap?: boolean;
  /** 单元格锁定（默认锁定；工作表保护后锁定的格子不能改） */
  locked?: boolean;
  /** 保护后在编辑栏里隐藏公式 */
  formulaHidden?: boolean;
  /** 边框：每条边 style（thin / medium / thick / double / dashed / dotted），null 去掉该边 */
  border?: Partial<Record<'left' | 'right' | 'top' | 'bottom', { style: string; color?: string } | null>>;
}

const DEFAULT_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="等线"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

/** 新建工作簿用的 styles.xml（等线 11 号、垂直居中，同 Excel 中文版默认） */
export const BLANK_STYLES_XML = DEFAULT_STYLES;

const section = (xml: string, tag: string) => xml.match(new RegExp(`<${tag}\\b[^>]*?(?:/>|>([\\s\\S]*?)</${tag}>)`))?.[1] ?? '';
const items = (body: string, tag: string) => [...body.matchAll(new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g'))].map((m) => m[0]);

function parseAttrs(s: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of s.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = m[2];
  return out;
}

export function readStyleBook(xml: string, themeXml = ''): StyleBook {
  const src = xml || DEFAULT_STYLES;
  const numFmts: Record<number, string> = {};
  for (const m of src.matchAll(/<numFmt\b[^>]*>/g)) {
    const code = xmlAttr(m[0], 'formatCode');
    if (code !== undefined) numFmts[Number(xmlAttr(m[0], 'numFmtId'))] = unescapeXml(code);
  }
  const xfs: Xf[] = [];
  for (const m of section(src, 'cellXfs').matchAll(/<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g)) xfs.push({ attrs: parseAttrs(m[1]), inner: m[2] ?? '' });
  if (!xfs.length) xfs.push({ attrs: { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0', xfId: '0' }, inner: '' });
  const fonts = items(section(src, 'fonts'), 'font');
  const pal = readPalette(src, themeXml);
  // 差异格式：只写了要改的项；填充用 bgColor（同 Excel 对 dxf 的约定）
  const dxfs = items(section(src, 'dxfs'), 'dxf').map((x): CellStyle => {
    const st: CellStyle = {};
    const font = x.match(/<font\b[\s\S]*?<\/font>/)?.[0];
    if (font) {
      const f = fontInfo(font, pal);
      if (/<b\s*\/>|<b\s+val="(1|true)"/.test(font)) st.bold = f.bold;
      if (/<i\s*\/>|<i\s+val="(1|true)"/.test(font)) st.italic = f.italic;
      if (/<u\b/.test(font)) st.underline = f.underline;
      if (/<strike\b/.test(font)) st.strike = f.strike;
      if (f.color) st.color = f.color;
    }
    const fill = x.match(/<fill\b[\s\S]*?<\/fill>/)?.[0];
    if (fill) {
      const bg = resolveColor(fill.match(/<bgColor\b[^>]*>/)?.[0], pal) ?? resolveColor(fill.match(/<fgColor\b[^>]*>/)?.[0], pal);
      if (bg) st.fill = bg;
    }
    const border = x.match(/<border\b[\s\S]*?<\/border>/)?.[0];
    if (border) {
      const b = resolveBorder(border.replace(/^<border\b[^>]*>|<\/border>$/g, ''), pal);
      if (b) st.border = b;
    }
    return st;
  });
  return {
    dxfs,
    numFmts,
    fonts: fonts.length ? fonts : ['<font><sz val="11"/><name val="等线"/></font>'],
    fills: items(section(src, 'fills'), 'fill'),
    borders: items(section(src, 'borders'), 'border'),
    xfs,
    palette: readPalette(src, themeXml),
  };
}

const num = (x: Xf, k: string) => Number(x.attrs[k] ?? 0);

export function formatCode(book: StyleBook, s: number | undefined): string {
  const xf = book.xfs[s ?? 0];
  if (!xf) return 'General';
  const id = num(xf, 'numFmtId');
  return book.numFmts[id] ?? builtinFormat(id);
}

interface FontInfo {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  name?: string;
  size?: number;
  color?: string;
}

const flag = (inner: string, tag: string) => new RegExp(`<${tag}(?:\\s*/>|\\s+val="(1|true|single|double|singleAccounting|doubleAccounting)")`).test(inner) && !new RegExp(`<${tag}\\s+val="(0|false|none)"`).test(inner);

export function fontInfo(fontXml: string, pal: Palette): FontInfo {
  const inner = fontXml.replace(/^<font\b[^>]*>|<\/font>$/g, '');
  const name = inner.match(/<name\b[^>]*>/)?.[0];
  const sz = inner.match(/<sz\b[^>]*>/)?.[0];
  return {
    bold: flag(inner, 'b'),
    italic: flag(inner, 'i'),
    underline: flag(inner, 'u'),
    strike: flag(inner, 'strike'),
    name: name ? unescapeXml(xmlAttr(name, 'val') ?? '') || undefined : undefined,
    size: sz ? Number(xmlAttr(sz, 'val')) || undefined : undefined,
    color: resolveColor(inner.match(/<color\b[^>]*>/)?.[0], pal),
  };
}

/** 单元格格式序号 → 显示用的样式 */
export function cellStyle(book: StyleBook, s: number): CellStyle {
  const xf = book.xfs[s];
  if (!xf) return {};
  const st: CellStyle = {};
  const font = book.fonts[num(xf, 'fontId')];
  if (font) {
    const f = fontInfo(font, book.palette);
    if (f.bold) st.bold = true;
    if (f.italic) st.italic = true;
    if (f.underline) st.underline = true;
    if (f.strike) st.strike = true;
    if (f.name) st.fontName = f.name;
    if (f.size) st.fontSize = f.size;
    if (f.color) st.color = f.color;
  }
  const fillXml = book.fills[num(xf, 'fillId')];
  const fill = fillXml ? resolveFill(fillXml, book.palette) : undefined;
  if (fill) st.fill = fill;
  const borderXml = book.borders[num(xf, 'borderId')];
  const border = borderXml ? resolveBorder(borderXml, book.palette) : undefined;
  if (border) st.border = border;
  const al = xf.inner.match(/<alignment\b[^>]*>/)?.[0] ?? '';
  const h = xmlAttr(al, 'horizontal');
  if (h === 'left' || h === 'right') st.align = h;
  else if (h === 'center' || h === 'centerContinuous') st.align = 'center';
  const v = xmlAttr(al, 'vertical');
  if (v === 'top' || v === 'bottom') st.valign = v;
  else if (v === 'center') st.valign = 'center';
  if (['1', 'true'].includes(xmlAttr(al, 'wrapText') ?? '')) st.wrap = true;
  return st;
}

export function allCellStyles(book: StyleBook): CellStyle[] {
  return book.xfs.map((_, i) => cellStyle(book, i));
}

/** 默认字体（fonts[0]）的名称 */
export function defaultFontName(book: StyleBook): string | undefined {
  return fontInfo(book.fonts[0] ?? '', book.palette).name;
}

/** 默认字体（fonts[0]）的字号，没有写时 11 */
export function defaultFontSize(book: StyleBook): number {
  return fontInfo(book.fonts[0] ?? '', book.palette).size ?? 11;
}

// ---------- 修改 ----------

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const argb = (css: string) => `FF${css.replace('#', '').toUpperCase()}`;

/** 数字格式代码 → numFmtId：内置的直接用，已有自定义的复用，否则追加一个 */
export function numFmtId(book: StyleBook, code: string): number {
  if (code === 'General') return 0;
  for (let id = 1; id < 164; id++) if (builtinFormat(id) === code) return id;
  for (const [id, c] of Object.entries(book.numFmts)) if (c === code) return Number(id);
  const id = Math.max(163, ...Object.keys(book.numFmts).map(Number)) + 1;
  book.numFmts[id] = code;
  book.dirty = true;
  return id;
}

function setFlag(inner: string, tag: string, on: boolean): string {
  const re = new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`, 'g');
  const out = inner.replace(re, '');
  return on ? `<${tag}/>${out}` : out;
}

function setVal(inner: string, tag: string, val: string | null): string {
  const re = new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`, 'g');
  const out = inner.replace(re, '');
  return val === null ? out : `${out}<${tag} ${val}/>`;
}

function changeFont(fontXml: string, ch: StyleChange): string {
  let inner = fontXml.replace(/^<font\b[^>]*?(?:\/>|>)|<\/font>$/g, '');
  if (ch.bold !== undefined) inner = setFlag(inner, 'b', ch.bold);
  if (ch.italic !== undefined) inner = setFlag(inner, 'i', ch.italic);
  if (ch.underline !== undefined) inner = setFlag(inner, 'u', ch.underline);
  if (ch.strike !== undefined) inner = setFlag(inner, 'strike', ch.strike);
  if (ch.fontSize !== undefined) inner = setVal(inner, 'sz', `val="${ch.fontSize}"`);
  if (ch.color !== undefined) inner = setVal(inner, 'color', ch.color === null ? null : `rgb="${argb(ch.color)}"`);
  if (ch.fontName !== undefined) {
    inner = setVal(inner, 'name', `val="${esc(ch.fontName)}"`);
    // 指定了字体就不再跟随主题字体（scheme 会覆盖 name）
    inner = setVal(inner, 'scheme', null);
  }
  return `<font>${inner}</font>`;
}

function fillXml(color: string): string {
  return `<fill><patternFill patternType="solid"><fgColor rgb="${argb(color)}"/><bgColor indexed="64"/></patternFill></fill>`;
}

function changeBorder(borderXml: string | undefined, sides: NonNullable<StyleChange['border']>): string {
  const inner = (borderXml ?? '<border/>').replace(/^<border\b[^>]*?(?:\/>|>)|<\/border>$/g, '');
  const attrs = borderXml?.match(/^<border\b([^>]*?)\/?>/)?.[1] ?? '';
  const get = (name: string) => inner.match(new RegExp(`<${name}\\b[^>]*/>|<${name}\\b[^>]*>[\\s\\S]*?</${name}>`))?.[0];
  const side = (name: 'left' | 'right' | 'top' | 'bottom', alias?: string) => {
    const v = sides[name];
    if (v === undefined) return get(name) ?? (alias ? get(alias)?.replace(alias, name).replace(`</${alias}>`, `</${name}>`) : undefined) ?? `<${name}/>`;
    if (v === null) return `<${name}/>`;
    return `<${name} style="${v.style}"><color ${v.color ? `rgb="${argb(v.color)}"` : 'indexed="64"'}/></${name}>`;
  };
  // CT_Border 的子元素有固定顺序：left right top bottom diagonal vertical horizontal
  const rest = ['diagonal', 'vertical', 'horizontal'].map(get).filter(Boolean).join('');
  return `<border${attrs}>${side('left', 'start')}${side('right', 'end')}${side('top')}${side('bottom')}${rest || '<diagonal/>'}</border>`;
}

function changeAlignment(inner: string, ch: StyleChange): string {
  const old = inner.match(/<alignment\b[^>]*\/>|<alignment\b[^>]*>[\s\S]*?<\/alignment>/)?.[0];
  const attrs = parseAttrs(old?.match(/^<alignment\b([^>]*?)\/?>/)?.[1] ?? '');
  if (ch.align !== undefined) {
    if (ch.align === null) delete attrs.horizontal;
    else attrs.horizontal = ch.align;
  }
  if (ch.valign !== undefined) {
    if (ch.valign === null) delete attrs.vertical;
    else attrs.vertical = ch.valign;
  }
  if (ch.wrap !== undefined) {
    if (ch.wrap) attrs.wrapText = '1';
    else delete attrs.wrapText;
  }
  const keys = Object.keys(attrs);
  const el = keys.length ? `<alignment ${keys.map((k) => `${k}="${attrs[k]}"`).join(' ')}/>` : '';
  const rest = old ? inner.replace(old, '') : inner;
  // CT_Xf 子元素顺序：alignment protection extLst
  return el + rest;
}

/** <protection locked hidden/> 放在 alignment 之后；两项都是默认值（锁定、不隐藏）时整个元素去掉 */
function changeProtection(inner: string, ch: StyleChange): string {
  const old = inner.match(/<protection\b[^>]*\/>/)?.[0];
  const attrs = parseAttrs(old?.match(/^<protection\b([^>]*?)\/?>/)?.[1] ?? '');
  if (ch.locked !== undefined) {
    if (ch.locked) delete attrs.locked;
    else attrs.locked = '0';
  }
  if (ch.formulaHidden !== undefined) {
    if (ch.formulaHidden) attrs.hidden = '1';
    else delete attrs.hidden;
  }
  const keys = Object.keys(attrs);
  const el = keys.length ? `<protection ${keys.map((k) => `${k}="${attrs[k]}"`).join(' ')}/>` : '';
  const rest = old ? inner.replace(old, '') : inner;
  const tail = rest.match(/<extLst\b[\s\S]*$/)?.[0] ?? '';
  return rest.slice(0, rest.length - tail.length) + el + tail;
}

const xfKey = (x: Xf) => `${Object.entries(x.attrs).map(([k, v]) => `${k}=${v}`).join(' ')}|${x.inner}`;

/**
 * 在格式 base 的基础上改动 → 新的格式序号（与已有的完全一样时复用已有的）。
 * 字体 / 填充 / 边框同样先找完全一样的已有条目，没有才追加。
 */
export function deriveStyle(book: StyleBook, base: number | undefined, ch: StyleChange): number {
  const src = book.xfs[base ?? 0] ?? book.xfs[0];
  const xf: Xf = { attrs: { ...src.attrs }, inner: src.inner };
  const findOrAdd = (list: string[], xml: string) => {
    const i = list.indexOf(xml);
    if (i >= 0) return i;
    list.push(xml);
    book.dirty = true;
    return list.length - 1;
  };
  if (ch.numFmt !== undefined) {
    xf.attrs.numFmtId = String(numFmtId(book, ch.numFmt));
    xf.attrs.applyNumberFormat = '1';
  }
  if (ch.bold !== undefined || ch.italic !== undefined || ch.underline !== undefined || ch.strike !== undefined || ch.fontSize !== undefined || ch.color !== undefined || ch.fontName !== undefined) {
    const font = book.fonts[num(xf, 'fontId')] ?? book.fonts[0];
    xf.attrs.fontId = String(findOrAdd(book.fonts, changeFont(font, ch)));
    xf.attrs.applyFont = '1';
  }
  if (ch.fill !== undefined) {
    xf.attrs.fillId = ch.fill === null ? '0' : String(findOrAdd(book.fills, fillXml(ch.fill)));
    xf.attrs.applyFill = '1';
  }
  if (ch.border) {
    xf.attrs.borderId = String(findOrAdd(book.borders, changeBorder(book.borders[num(xf, 'borderId')], ch.border)));
    xf.attrs.applyBorder = '1';
  }
  if (ch.align !== undefined || ch.valign !== undefined || ch.wrap !== undefined) {
    xf.inner = changeAlignment(xf.inner, ch);
    xf.attrs.applyAlignment = '1';
  }
  if (ch.locked !== undefined || ch.formulaHidden !== undefined) {
    xf.inner = changeProtection(xf.inner, ch);
    xf.attrs.applyProtection = '1';
  }
  const key = xfKey(xf);
  const same = book.xfs.findIndex((x) => xfKey(x) === key);
  if (same >= 0) return same;
  book.xfs.push(xf);
  book.dirty = true;
  return book.xfs.length - 1;
}

// ---------- 写回 ----------

/** styleSheet 子元素的顺序（CT_Stylesheet） */
const ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];

function putSection(xml: string, tag: string, content: string): string {
  const re = new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`);
  if (re.test(xml)) return xml.replace(re, () => content);
  // 插到后面第一个已有的段之前，没有就放到 </styleSheet> 前
  for (const next of ORDER.slice(ORDER.indexOf(tag) + 1)) {
    const m = xml.match(new RegExp(`<${next}\\b`));
    if (m?.index !== undefined) return xml.slice(0, m.index) + content + xml.slice(m.index);
  }
  return xml.replace(/<\/styleSheet>\s*$/, `${content}</styleSheet>`);
}

/**
 * 按颜色筛选要引用一个差异格式（colorFilter dxfId）：cell = 按单元格底色（填充写在 bgColor），否则按字体颜色。
 * 已有同色的就用它，没有就追加一个。返回 dxf 序号。
 */
export function dxfForColor(book: StyleBook, color: string, cell: boolean): number {
  const list = (book.dxfs ??= []);
  const key = color.toLowerCase();
  const hit = list.findIndex((d) => {
    const keys = Object.keys(d);
    return keys.length === 1 && (cell ? d.fill?.toLowerCase() === key : d.color?.toLowerCase() === key);
  });
  if (hit >= 0) return hit;
  list.push(cell ? { fill: color } : { color });
  (book.newDxfs ??= []).push(cell ? `<dxf><fill><patternFill patternType="solid"><fgColor rgb="${argb(color)}"/><bgColor rgb="${argb(color)}"/></patternFill></fill></dxf>` : `<dxf><font><color rgb="${argb(color)}"/></font></dxf>`);
  book.dirty = true;
  return list.length - 1;
}

/** 条件格式用的差异格式：底色 / 字体颜色 / 加粗的任意组合；已有完全相同的就用它，没有就追加。返回 dxf 序号。 */
export function dxfForStyle(book: StyleBook, st: { fill?: string; color?: string; bold?: boolean }): number {
  const list = (book.dxfs ??= []);
  const same = (d: CellStyle) => (d.fill?.toLowerCase() ?? '') === (st.fill?.toLowerCase() ?? '') && (d.color?.toLowerCase() ?? '') === (st.color?.toLowerCase() ?? '') && !!d.bold === !!st.bold && Object.keys(d).every((k) => ['fill', 'color', 'bold'].includes(k));
  const hit = list.findIndex(same);
  if (hit >= 0) return hit;
  list.push({ ...(st.fill ? { fill: st.fill } : {}), ...(st.color ? { color: st.color } : {}), ...(st.bold ? { bold: true } : {}) });
  const font = st.color || st.bold ? `<font>${st.bold ? '<b/>' : ''}${st.color ? `<color rgb="${argb(st.color)}"/>` : ''}</font>` : '';
  const fill = st.fill ? `<fill><patternFill patternType="solid"><fgColor rgb="${argb(st.fill)}"/><bgColor rgb="${argb(st.fill)}"/></patternFill></fill>` : '';
  (book.newDxfs ??= []).push(`<dxf>${font}${fill}</dxf>`);
  book.dirty = true;
  return list.length - 1;
}

/** 把编辑后的 StyleBook 写回 styles.xml（原文件为空时用默认样式表为底） */
export function writeStyles(xml: string, book: StyleBook): string {
  let out = xml || DEFAULT_STYLES;
  if (book.newDxfs?.length) {
    const old = items(section(out, 'dxfs'), 'dxf');
    const all = [...old, ...book.newDxfs];
    out = putSection(out, 'dxfs', `<dxfs count="${all.length}">${all.join('')}</dxfs>`);
  }
  const fmts = Object.entries(book.numFmts);
  if (fmts.length) out = putSection(out, 'numFmts', `<numFmts count="${fmts.length}">${fmts.map(([id, c]) => `<numFmt numFmtId="${id}" formatCode="${esc(c)}"/>`).join('')}</numFmts>`);
  out = putSection(out, 'fonts', `<fonts count="${book.fonts.length}">${book.fonts.join('')}</fonts>`);
  if (book.fills.length) out = putSection(out, 'fills', `<fills count="${book.fills.length}">${book.fills.join('')}</fills>`);
  if (book.borders.length) out = putSection(out, 'borders', `<borders count="${book.borders.length}">${book.borders.join('')}</borders>`);
  const xfs = book.xfs.map((x) => {
    const a = Object.entries(x.attrs).map(([k, v]) => ` ${k}="${v}"`).join('');
    return x.inner ? `<xf${a}>${x.inner}</xf>` : `<xf${a}/>`;
  });
  return putSection(out, 'cellXfs', `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>`);
}

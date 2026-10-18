/**
 * .docx 读写（Office Open XML 的常用子集）。
 * 读取时把不支持的内容记入 doc.lossy，保存时这些内容会丢失，界面据此提示用户另存。
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  type Columns,
  type CellEdge,
  type BorderStyle,
  type CellEdges,
  type TableFloat,
  FIELD_CHAR,
  HF_STORIES,
  defaultSetup,
  isAtom,
  isHeaderStory,
  pageGeom,
  sameGeom,
  sectionList,
  type Align,
  type Block,
  type Cell,
  type Doc,
  type Watermark,
  isPictureWatermark,
  type GridType,
  type HFStory,
  type NumFmt,
  type NumLevel,
  type NumList,
  type Opaque,
  type PageGeom,
  type Paragraph,
  type ParaProps,
  type Row,
  type Run,
  type RunImage,
  type RunProps,
  type Revision,
  type Comment,
  type SectionBreak,
  type SectionHF,
  type Table,
  type TextBox,
} from '../core/model';
import { imageData, imageExt, imageMime, registerImage } from '../core/images';
import { eqText, parseEq, type EqChar } from '../core/eq';
import { instrPieces, newFieldId, pairedMarkers } from '../core/fieldView';
import { mergeFieldView, parseMergeInstr, parseRule, ruleFieldView } from '../core/merge';
import { upsertStyleXml } from '../../vendor/docx-engine/src/style-upsert';
import { editToUpsert } from '../core/styles';
import { SHAPE_GEOMS, applyColorMods, isLineGeom } from '../core/shapes';
import { pictureWatermarkParagraphXml, readPictureWatermark, readWatermarkShape, watermarkParagraphXml } from '../../vendor/docx-engine/src/watermark';

/** XML 里引用了关系（r:id、r:embed…）：这样的原样对象只能在字节级保留保存时写回 */
const REL_REF = /\br:(?:id|embed|link|pict|dm|lo|qs|cs)="/;

/**
 * 元素 → XML 文字（原样对象、原图片要写回原文档用）。不用 XMLSerializer：它单独序列化一个元素时会补命名空间声明，
 * 还会把子树里没声明的属性前缀改成 ns1: 之类。这里元素名、属性名都照原样（带原来的前缀）输出，
 * 前缀由原文档根元素上的声明负责。
 */
function serialize(node: Node): string {
  switch (node.nodeType) {
    case 1: {
      const el = node as Element;
      let s = `<${el.tagName}`;
      for (const a of Array.from(el.attributes)) s += ` ${a.name}="${a.value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;')}"`;
      if (!el.childNodes.length) return `${s}/>`;
      return `${s}>${Array.from(el.childNodes).map(serialize).join('')}</${el.tagName}>`;
    }
    case 3:
      return (node.nodeValue ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    case 4:
      return `<![CDATA[${node.nodeValue ?? ''}]]>`;
    case 8:
      return `<!--${node.nodeValue ?? ''}-->`;
    default:
      return '';
  }
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
/** 图片尺寸单位：1 磅 = 12700 EMU */
const EMU_PER_PT = 12700;
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WRAPS: Record<string, TextBox['wrap']> = { wrapNone: 'none', wrapSquare: 'square', wrapTight: 'tight', wrapThrough: 'through', wrapTopAndBottom: 'topAndBottom' };
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ---------------- 读取 ----------------

function kids(el: Element | null | undefined, name?: string): Element[] {
  if (!el) return [];
  return Array.from(el.children).filter((c) => !name || c.localName === name);
}
function kid(el: Element | null | undefined, name: string): Element | null {
  return kids(el, name)[0] ?? null;
}
function pictureTransform(pic: Element | null | undefined): Partial<RunImage> {
  const x = kid(kid(pic, 'spPr'), 'xfrm');
  const crop = kid(kid(pic, 'blipFill'), 'srcRect');
  const rot = Number(x?.getAttribute('rot') ?? 0) / 60000;
  return {
    ...(rot ? { rot } : {}),
    ...(x?.getAttribute('flipH') === '1' || x?.getAttribute('flipH') === 'true' ? { flipH: true } : {}),
    ...(x?.getAttribute('flipV') === '1' || x?.getAttribute('flipV') === 'true' ? { flipV: true } : {}),
    ...(crop ? { crop: { l: Number(crop.getAttribute('l') ?? 0) / 100000, t: Number(crop.getAttribute('t') ?? 0) / 100000, r: Number(crop.getAttribute('r') ?? 0) / 100000, b: Number(crop.getAttribute('b') ?? 0) / 100000 } } : {}),
  };
}
function attr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(W_NS, name) ?? el.getAttribute(`w:${name}`);
}
function num(el: Element | null | undefined, name: string): number | null {
  const v = attr(el, name);
  if (v === null || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
/** 分栏 w:cols（一栏时只写默认间距） */
function colsXml(c: Columns | undefined): string {
  if (!c || c.num <= 1) return '<w:cols w:space="425"/>';
  const head = `<w:cols w:num="${c.num}" w:space="${twip(c.space)}"${c.sep ? ' w:sep="1"' : ''}`;
  if (!c.widths) return `${head}/>`;
  return `${head} w:equalWidth="0">${c.widths.map((x) => `<w:col w:w="${twip(x.w)}" w:space="${twip(x.space)}"/>`).join('')}</w:cols>`;
}

/** wp:anchor 的环绕距离（EMU → 磅） */
function anchorDist(anchor: Element): { t: number; b: number; l: number; r: number } {
  const d = (n: string) => (Number(anchor.getAttribute(n)) || 0) / 12700;
  return { t: d('distT'), b: d('distB'), l: d('distL'), r: d('distR') };
}

/** 浮动表格 w:tblpPr（缇 → 磅）；没有时 undefined */
function tableFloat(tblPr: Element | null): TableFloat | undefined {
  const f = kid(tblPr, 'tblpPr');
  if (!f) return undefined;
  const anchorOf = (v: string | null): TableFloat['hAnchor'] => (v === 'margin' || v === 'page' ? v : 'text');
  const tw = (n: string) => (num(f, n) ?? 0) / 20;
  const xAlign = attr(f, 'tblpXSpec');
  const yAlign = attr(f, 'tblpYSpec');
  return {
    hAnchor: anchorOf(attr(f, 'horzAnchor')),
    vAnchor: anchorOf(attr(f, 'vertAnchor')),
    x: tw('tblpX'),
    y: tw('tblpY'),
    ...(xAlign ? { xAlign } : {}),
    ...(yAlign ? { yAlign } : {}),
    dist: { l: tw('leftFromText'), r: tw('rightFromText'), t: tw('topFromText'), b: tw('bottomFromText') },
  };
}

/** w:tblW w:type="pct"：值是五十分之一百分点（5000 = 100%），也可以写成“75%” */
function tablePct(tblPr: Element | null): number | undefined {
  const w = kid(tblPr, 'tblW');
  if (attr(w, 'type') !== 'pct') return undefined;
  const v = attr(w, 'w') ?? '';
  const n = v.endsWith('%') ? parseFloat(v) / 100 : Number(v) / 5000;
  return Number.isFinite(n) && n > 0 ? n : undefined;
}

function tableFloatXml(f: TableFloat | undefined): string {
  if (!f) return '';
  const t = (v: number) => Math.round(v * 20);
  return (
    `<w:tblpPr w:leftFromText="${t(f.dist.l)}" w:rightFromText="${t(f.dist.r)}" w:topFromText="${t(f.dist.t)}" w:bottomFromText="${t(f.dist.b)}"` +
    ` w:vertAnchor="${f.vAnchor}" w:horzAnchor="${f.hAnchor}"` +
    (f.xAlign ? ` w:tblpXSpec="${f.xAlign}"` : ` w:tblpX="${t(f.x)}"`) +
    (f.yAlign ? ` w:tblpYSpec="${f.yAlign}"` : ` w:tblpY="${t(f.y)}"`) +
    '/>'
  );
}

/** 一组边框（上下左右、内部横竖线）：undefined = 没写，null = 不画（nil / none） */
type BorderSet = Partial<Record<'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV', CellEdge | null>>;

/** 依次合并多个 tblBorders / tcBorders（后面的覆盖前面的；start / end 当作 left / right） */
function mergeBorders(els: (Element | null)[]): BorderSet {
  const out: BorderSet = {};
  for (const el of els) {
    for (const b of kids(el)) {
      const name = b.localName === 'start' ? 'left' : b.localName === 'end' ? 'right' : b.localName;
      if (!['top', 'bottom', 'left', 'right', 'insideH', 'insideV'].includes(name)) continue;
      const v = attr(b, 'val');
      const style = edgeStyle(v);
      (out as Record<string, CellEdge | null>)[name] =
        !v || v === 'nil' || v === 'none'
          ? null
          : { w: Math.max(0.25, (num(b, 'sz') ?? 4) / 8) * (v === 'double' ? 2 : 1), color: (attr(b, 'color') ?? 'auto').toUpperCase() === 'AUTO' ? 'auto' : (attr(b, 'color') ?? '000000').toUpperCase(), ...(style && { style }) };
    }
  }
  return out;
}

/** 边框线型（w:val）→ 本编辑器能画的线型；单实线和不认识的线型返回 undefined（按单实线画） */
function edgeStyle(v: string | null): BorderStyle | undefined {
  if (!v) return undefined;
  if (v === 'double') return 'double';
  if (v === 'dotted') return 'dotted';
  if (/^dash/.test(v)) return 'dashed';
  return undefined;
}

/** 单元格自己设的框线（Cell.bdr）→ w:tcBorders；线宽按 1/8 磅，双线的 w:sz 是单条线的宽度 */
export function tcBordersXml(bdr: Partial<CellEdges>): string {
  const side = (n: 'top' | 'left' | 'bottom' | 'right') => {
    const e = bdr[n];
    if (e === undefined) return '';
    if (e === null) return `<w:${n} w:val="nil"/>`;
    const style = e.style ?? 'single';
    const sz = Math.max(2, Math.min(96, Math.round((style === 'double' ? e.w / 2 : e.w) * 8)));
    return `<w:${n} w:val="${style}" w:sz="${sz}" w:space="0" w:color="${e.color === 'auto' ? 'auto' : e.color}"/>`;
  };
  return `<w:tcBorders>${side('top')}${side('left')}${side('bottom')}${side('right')}</w:tcBorders>`;
}

/** w:tblLook：新写法是各个属性，旧写法是 val 里的位（0x20 首行、0x40 末行、0x80 首列、0x100 末列、0x200 无镶边行、0x400 无镶边列） */
function tableLook(el: Element | null): { firstRow: boolean; lastRow: boolean; firstColumn: boolean; lastColumn: boolean; noHBand: boolean; noVBand: boolean } {
  const bits = parseInt(attr(el, 'val') ?? '0', 16) || 0;
  const flag = (name: string, bit: number) => {
    const v = attr(el, name);
    return v !== null ? v === '1' || v === 'true' || v === 'on' : (bits & bit) !== 0;
  };
  return {
    firstRow: flag('firstRow', 0x20),
    lastRow: flag('lastRow', 0x40),
    firstColumn: flag('firstColumn', 0x80),
    lastColumn: flag('lastColumn', 0x100),
    noHBand: flag('noHBand', 0x200),
    noVBand: flag('noVBand', 0x400),
  };
}

/** 布尔属性（w:fldLock 等）：1 / true / on 为真 */
function onoffAttr(el: Element, name: string): boolean {
  const v = attr(el, name);
  return v === '1' || v === 'true' || v === 'on';
}

/** 域的开始部分（begin … separate）/ 结束部分（end）的 w:r 只能有这些子元素，才原样记下 */
const HEAD_KIDS = new Set(['rPr', 'fldChar', 'instrText']);
const TAIL_KIDS = new Set(['rPr', 'fldChar']);

/** 开关属性：元素存在且 val 不是 0/false/off。 */
function onoff(el: Element | null): boolean | null {
  if (!el) return null;
  const v = attr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

interface Theme {
  minorLatin: string;
  minorEA: string;
  majorLatin: string;
  majorEA: string;
  /** 主题颜色（dk1、lt1、accent1…，'RRGGBB'） */
  colors?: Record<string, string>;
  /** 主题线条样式的线宽（fmtScheme/lnStyleLst，磅；形状样式的 lnRef idx 从 1 数） */
  lineWidths?: number[];
}

/** 形状里的颜色名 → 主题颜色（settings 的 clrSchemeMapping 缺省映射） */
const SCHEME_ALIAS: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };

interface StyleDef {
  /** 样式元素本身（表格样式的 tblStylePr 条件格式从这里取） */
  el: Element;
  basedOn: string | null;
  /** w:name，小写（heading 1、toc 1 …） */
  name: string;
  pPr: Element | null;
  rPr: Element | null;
  tblBorders: boolean | null;
  tblBorderSize: number | null;
}

/** 边框元素（tblBorders / tcBorders）里有没有可见的边，以及最粗的一条有多粗（磅；w:sz 的单位是 1/8 磅） */
function borderInfo(el: Element | null): { on: boolean; size: number | null } | null {
  if (!el) return null;
  const visible = kids(el).filter((b) => !['none', 'nil'].includes(attr(b, 'val') ?? 'none'));
  const sizes = visible.map((b) => num(b, 'sz')).filter((v): v is number => v !== null && v > 0);
  return { on: visible.length > 0, size: sizes.length ? Math.max(...sizes) / 8 : null };
}

const NUM_FMTS: Record<string, NumFmt> = {
  decimal: 'decimal',
  decimalZero: 'decimal',
  chineseCounting: 'chineseCounting',
  chineseCountingThousand: 'chineseCounting',
  chineseLegalSimplified: 'chineseCounting',
  japaneseCounting: 'chineseCounting',
  taiwaneseCounting: 'chineseCounting',
  taiwaneseCountingThousand: 'chineseCounting',
  ideographTraditional: 'ideographTraditional',
  upperLetter: 'upperLetter',
  lowerLetter: 'lowerLetter',
  upperRoman: 'upperRoman',
  lowerRoman: 'lowerRoman',
  decimalEnclosedCircle: 'decimalEnclosedCircle',
  decimalEnclosedCircleChinese: 'decimalEnclosedCircle',
  bullet: 'bullet',
  none: 'none',
};

/** Symbol / Wingdings 字体的私用区项目符号换成通用字符 */
export function bulletChar(t: string): string {
  const c = t.codePointAt(0) ?? 0;
  if (c >= 0xf000 && c <= 0xf0ff) return { 0xf0a7: '■', 0xf0d8: '▶', 0xf076: '◆', 0xf06f: '○', 0xf0fc: '✓' }[c] ?? '●';
  return t || '●';
}

class Reader {
  lossy = new Set<string>();
  lists: NumList[] = [];
  /** 正文中间的分节（段落属性里的 sectPr） */
  sections: Element[] = [];
  /** 带分节符的段落对应的 sectPr */
  sectOf = new WeakMap<Paragraph, Element>();
  styles = new Map<string, StyleDef>();
  defaultParaStyle: string | null = null;
  baseRun!: RunProps;
  basePara!: ParaProps;
  theme: Theme = { minorLatin: 'Times New Roman', minorEA: '宋体', majorLatin: 'Times New Roman', majorEA: '黑体' };

  /** numbering.xml 里自带缩进的编号级别：段落引用它时先套用这份缩进 */
  levelInd = new WeakSet<NumLevel>();
  /** 当前部件（正文）的关系表，用来找嵌入图片的文件；读页眉页脚时为 null（那里的图片暂不支持） */
  imageRels: Map<string, string> | null = null;

  constructor(private files: Record<string, Uint8Array>) {}

  /**
   * 嵌入式图片（wp:inline 里的 pic:pic，浏览器能显示的格式）→ 登记到图片仓库。
   * 浮动图片（wp:anchor）、图表、形状、EMF / WMF 等返回 null，仍按“不支持的内容”处理。
   */
  inlineImage(drawing: Element): RunImage | null {
    const inline = kid(drawing, 'inline');
    const extent = kid(inline, 'extent');
    const blip = inline?.getElementsByTagNameNS('*', 'blip')[0];
    if (!inline || !extent || !blip || !this.imageRels) return null;
    const rid = blip.getAttributeNS(R_NS, 'embed') ?? blip.getAttribute('r:embed');
    const target = rid ? this.imageRels.get(rid) : undefined;
    if (!target) return null;
    // 关系里的路径相对 word/ 目录；也可能写成 /word/media/...
    const path = target.startsWith('/') ? target.slice(1) : `word/${target}`;
    const data = this.files[path];
    const mime = imageMime(path);
    const w = Number(extent.getAttribute('cx')) / EMU_PER_PT;
    const h = Number(extent.getAttribute('cy')) / EMU_PER_PT;
    if (!data || !mime || !(w > 0) || !(h > 0)) return null;
    const id = registerImage(data, mime);
    return { id, rawId: id, w, h, ...pictureTransform(inline.getElementsByTagNameNS('*', 'pic')[0]), raw: serialize(drawing) };
  }

  /**
   * DrawingML 颜色元素（srgbClr / schemeClr / sysClr / prstClr，带 shade / tint / lumMod / lumOff 变换）→ 'RRGGBB'；
   * phClr（样式里的占位颜色）取 ph。
   */
  drawingColor(parent: Element | null, ph?: string | null): string | null {
    const c = kids(parent).find((x) => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr'].includes(x.localName));
    if (!c) return null;
    const val = c.getAttribute('val') ?? '';
    let base: string | null | undefined =
      c.localName === 'srgbClr'
        ? val
        : c.localName === 'sysClr'
          ? c.getAttribute('lastClr')
          : c.localName === 'prstClr'
            ? ({ black: '000000', white: 'FFFFFF', red: 'FF0000', green: '00FF00', blue: '0000FF', yellow: 'FFFF00' } as Record<string, string>)[val]
            : val === 'phClr'
              ? ph
              : this.theme.colors?.[SCHEME_ALIAS[val] ?? val];
    if (!base || !/^[0-9a-fA-F]{6}$/.test(base)) return null;
    base = base.toUpperCase();
    const mods = kids(c).map((m) => ({ name: m.localName, val: Number(m.getAttribute('val')) })).filter((m) => Number.isFinite(m.val));
    return mods.length ? applyColorMods(base, mods) : base;
  }

  /**
   * 文本框和形状：浮动绘图（wp:anchor）里的一个 wps:wsp（DrawingML 形状）。
   * 带文字（wps:txbx）的读出文字；外形是 SHAPE_GEOMS 里的预设形状。
   * 填充 / 线条：spPr 里写了就用 spPr 的（noFill = 无），没写的用形状样式（wps:style 的 fillRef / lnRef，idx 0 = 无）。
   * 组合、画布、自由曲线、不认识的预设形状返回 null（原样保留，不显示）。
   */
  textBox(drawing: Element): TextBox | null {
    const anchor = kid(drawing, 'anchor');
    const gd = kid(kid(anchor, 'graphic'), 'graphicData');
    // 单个形状的 wps:wsp 直接在 graphicData 下；组合（wpg:wgp）、画布（wpc:wpc）里的不按单个形状读
    const wsp = kid(gd, 'wsp');
    const extent = kid(anchor, 'extent');
    if (!anchor || !wsp || !extent) return null;
    const content = kid(kid(wsp, 'txbx'), 'txbxContent');
    const spPr = kid(wsp, 'spPr');
    const named = kid(spPr, 'prstGeom')?.getAttribute('prst') ?? null;
    // 没有外形（没写 prstGeom）、自由曲线、不认识的预设形状：带文字的按矩形文本框显示（文字才是主要内容），不带文字的不读
    const prst = named && SHAPE_GEOMS.has(named) ? named : null;
    if (!prst && !content) return null;
    const txBox = kid(wsp, 'cNvSpPr')?.getAttribute('txBox') === '1';
    const emu = (el: Element | null, name: string, d: number) => {
      const v = el?.getAttribute(name);
      return (v === null || v === undefined || v === '' ? d : Number(v)) / EMU_PER_PT;
    };
    const pos = (name: string, rel: string): TextBox['posH'] => {
      const el = kid(anchor, name);
      const align = kid(el, 'align')?.textContent?.trim();
      return { rel: el?.getAttribute('relativeFrom') ?? rel, offset: Number(kid(el, 'posOffset')?.textContent ?? 0) / EMU_PER_PT, ...(align ? { align } : {}) };
    };
    const style = kid(wsp, 'style');
    const refIdx = (el: Element | null) => Number(el?.getAttribute('idx') ?? 0);
    const fillRef = kid(style, 'fillRef');
    const lnRef = kid(style, 'lnRef');
    /** spPr 里的填充：undefined = 没写（用样式） */
    const fillOf = (el: Element | null): string | null | undefined => {
      if (kid(el, 'noFill')) return null;
      const solid = kid(el, 'solidFill');
      if (solid) return this.drawingColor(solid);
      // 渐变：按第一个色标近似
      const gs = kid(kid(kid(el, 'gradFill'), 'gsLst'), 'gs');
      if (gs) return this.drawingColor(gs);
      return undefined;
    };
    const ownFill = fillOf(spPr);
    const fill = ownFill !== undefined ? ownFill : refIdx(fillRef) > 0 ? this.drawingColor(fillRef) : null;
    const ln = kid(spPr, 'ln');
    const ownLine = fillOf(ln);
    const lineColor = ownLine !== undefined ? ownLine : refIdx(lnRef) > 0 ? this.drawingColor(lnRef) : null;
    const lineW = ln?.getAttribute('w') ? emu(ln, 'w', 9525) : (this.theme.lineWidths?.[refIdx(lnRef) - 1] ?? 0.75);
    const xfrm = kid(spPr, 'xfrm');
    const rot = Number(xfrm?.getAttribute('rot') ?? 0) / 60000;
    const bodyPr = kid(wsp, 'bodyPr');
    const wrapEl = kids(anchor).find((c) => c.localName in WRAPS);
    const va = bodyPr?.getAttribute('anchor');
    if (content && kids(content, 'tbl').length) this.lossy.add('文本框中的表格');
    const paragraphs = content ? kids(content, 'p').map((p) => this.paragraph(p)) : [];
    return {
      w: emu(extent, 'cx', 0),
      h: emu(extent, 'cy', 0),
      posH: pos('positionH', 'column'),
      posV: pos('positionV', 'paragraph'),
      wrap: wrapEl ? WRAPS[wrapEl.localName] : 'none',
      dist: anchorDist(anchor),
      behind: anchor.getAttribute('behindDoc') === '1',
      fill,
      line: lineColor ? { color: lineColor, w: lineW } : null,
      inset: { l: emu(bodyPr, 'lIns', 91440), t: emu(bodyPr, 'tIns', 45720), r: emu(bodyPr, 'rIns', 91440), b: emu(bodyPr, 'bIns', 45720) },
      vAlign: va === 'ctr' ? 'center' : va === 'b' ? 'bottom' : 'top',
      paragraphs: content ? (paragraphs.length ? paragraphs : [{ kind: 'p', props: this.basePara, runs: [], mark: this.baseRun }]) : [],
      ...(!prst || (txBox && prst === 'rect') ? {} : { geom: prst }),
      ...(xfrm?.getAttribute('flipH') === '1' ? { flipH: true } : {}),
      ...(xfrm?.getAttribute('flipV') === '1' ? { flipV: true } : {}),
      ...(rot ? { rot } : {}),
    };
  }

  /**
   * 浮动图片：wp:anchor 里的 pic:pic（公章、签名、Logo 等），按文本框的方式定位，带上图片。
   */
  floatImage(drawing: Element): TextBox | null {
    const anchor = kid(drawing, 'anchor');
    const extent = kid(anchor, 'extent');
    const pic = anchor?.getElementsByTagNameNS('*', 'pic')[0];
    const blip = pic?.getElementsByTagNameNS('*', 'blip')[0];
    if (!anchor || !extent || !blip || !this.imageRels) return null;
    const rid = blip.getAttributeNS(R_NS, 'embed') ?? blip.getAttribute('r:embed');
    const target = rid ? this.imageRels.get(rid) : undefined;
    if (!target) return null;
    const path = target.startsWith('/') ? target.slice(1) : `word/${target}`;
    const data = this.files[path];
    const mime = imageMime(path);
    const w = Number(extent.getAttribute('cx')) / EMU_PER_PT;
    const h = Number(extent.getAttribute('cy')) / EMU_PER_PT;
    if (!data || !mime || !(w > 0) || !(h > 0)) return null;
    const pos = (name: string, rel: string): TextBox['posH'] => {
      const el = kid(anchor, name);
      const align = kid(el, 'align')?.textContent?.trim();
      return { rel: el?.getAttribute('relativeFrom') ?? rel, offset: Number(kid(el, 'posOffset')?.textContent ?? 0) / EMU_PER_PT, ...(align ? { align } : {}) };
    };
    const wrapEl = kids(anchor).find((c) => c.localName in WRAPS);
    return {
      w,
      h,
      posH: pos('positionH', 'column'),
      posV: pos('positionV', 'paragraph'),
      wrap: wrapEl ? WRAPS[wrapEl.localName] : 'none',
      dist: anchorDist(anchor),
      behind: anchor.getAttribute('behindDoc') === '1',
      fill: null,
      line: null,
      inset: { l: 0, t: 0, r: 0, b: 0 },
      vAlign: 'top',
      paragraphs: [],
      image: { id: registerImage(data, mime), rawId: registerImage(data, mime), w, h, ...pictureTransform(pic) },
    };
  }

  /** word/footnotes.xml → 普通脚注的段落（分隔线等特殊脚注不读；脚注里的表格不读） */
  readFootnotes(rels: Map<string, string>): Record<string, Paragraph[]> {
    const target = [...rels.values()].find((t) => /(^|\/)footnotes\.xml$/i.test(t));
    const x = target ? this.xml(target.startsWith('/') ? target.slice(1) : `word/${target}`) : null;
    const out: Record<string, Paragraph[]> = {};
    for (const fn of kids(x?.documentElement, 'footnote')) {
      const type = attr(fn, 'type');
      if (type && type !== 'normal') continue;
      out[attr(fn, 'id') ?? ''] = kids(fn, 'p').map((p) => this.paragraph(p));
    }
    return out;
  }

  /** word/comments.xml → 批注；commentsExtended.xml 里的 paraIdParent → 回复关系 */
  readComments(rels: Map<string, string>): Comment[] {
    const target = [...rels.values()].find((t) => /(^|\/)comments\.xml$/i.test(t)) ?? 'comments.xml';
    const x = this.xml(target.startsWith('/') ? target.slice(1) : `word/${target}`);
    if (!x) return [];
    const byPara = new Map<string, string>();
    const out: Comment[] = kids(x.documentElement, 'comment').map((c) => {
      const id = attr(c, 'id') ?? '';
      const ps = kids(c, 'p');
      const last = ps[ps.length - 1];
      const paraId = last?.getAttribute('w14:paraId') ?? last?.getAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'paraId');
      if (paraId) byPara.set(paraId, id);
      const text = ps.map((p) => Array.from(p.getElementsByTagNameNS(W_NS, 't')).map((t) => t.textContent ?? '').join('')).join('\n');
      const initials = attr(c, 'initials');
      return { id, author: attr(c, 'author') ?? '', date: attr(c, 'date') ?? '', text, ...(initials ? { initials } : {}) };
    });
    const ext = this.xml('word/commentsExtended.xml');
    for (const e of Array.from(ext?.documentElement.children ?? [])) {
      const self = byPara.get(e.getAttribute('w15:paraId') ?? '');
      const parent = byPara.get(e.getAttribute('w15:paraIdParent') ?? '');
      const c = out.find((x) => x.id === self);
      if (c && parent && parent !== self) c.parent = parent;
    }
    return out;
  }

  xml(path: string): Document | null {
    const f = this.files[path];
    if (!f) return null;
    return new DOMParser().parseFromString(strFromU8(f), 'application/xml');
  }

  readNumbering() {
    const n = this.xml('word/numbering.xml');
    if (!n) return;
    const root = n.documentElement;
    const abstract = new Map<string, NumLevel[]>();
    // 哪些级别在 numbering.xml 里写了缩进（没写的不影响段落缩进）
    const withInd = new WeakSet<NumLevel>();
    for (const a of kids(root, 'abstractNum')) {
      const levels: NumLevel[] = [];
      for (const lvl of kids(a, 'lvl')) {
        const i = num(lvl, 'ilvl') ?? levels.length;
        const fmt = NUM_FMTS[attr(kid(lvl, 'numFmt'), 'val') ?? 'decimal'] ?? 'decimal';
        const ind = kid(kid(lvl, 'pPr'), 'ind');
        const hg = num(ind, 'hanging');
        const fl = num(ind, 'firstLine');
        const flc = num(ind, 'firstLineChars');
        const suff = attr(kid(lvl, 'suff'), 'val');
        const text = attr(kid(lvl, 'lvlText'), 'val') ?? '';
        // 编号自己的字符格式（只记写了的项）
        const lr = kid(lvl, 'rPr');
        const label: Partial<RunProps> = {};
        const lb = onoff(kid(lr, 'b'));
        const li = onoff(kid(lr, 'i'));
        const lsz = num(kid(lr, 'sz'), 'val');
        const lc = attr(kid(lr, 'color'), 'val');
        const lf = kid(lr, 'rFonts');
        if (lb !== null) label.bold = lb;
        if (li !== null) label.italic = li;
        if (lsz) label.size = lsz / 2;
        if (lc && /^[0-9a-fA-F]{6}$/.test(lc)) label.color = lc.toUpperCase();
        if (attr(lf, 'ascii')) label.fontLatin = attr(lf, 'ascii')!;
        if (attr(lf, 'eastAsia')) label.fontEA = attr(lf, 'eastAsia')!;
        levels[i] = {
          fmt,
          text: fmt === 'bullet' ? bulletChar(text) : text,
          start: num(kid(lvl, 'start'), 'val') ?? 1,
          indentLeft: (num(ind, 'left') ?? num(ind, 'start') ?? 0) / 20,
          firstLine: hg !== null ? -hg / 20 : (fl ?? 0) / 20,
          firstLineChars: flc !== null ? flc / 100 : null,
          suffix: suff === 'space' ? 'space' : suff === 'nothing' ? 'nothing' : 'tab',
          ...(Object.keys(label).length ? { label } : {}),
        };
        if (ind) withInd.add(levels[i]);
      }
      for (let i = 0; i < 9; i++) levels[i] ??= { fmt: 'decimal', text: `%${i + 1}.`, start: 1, indentLeft: 21 * (i + 1), firstLine: -21, firstLineChars: null, suffix: 'tab' };
      abstract.set(attr(a, 'abstractNumId') ?? '', levels);
    }
    for (const nm of kids(root, 'num')) {
      const id = num(nm, 'numId');
      const base = abstract.get(attr(kid(nm, 'abstractNumId'), 'val') ?? '');
      if (id === null || !base) continue;
      const levels = base.map((l) => {
        const copy = { ...l };
        if (withInd.has(l)) this.levelInd.add(copy);
        return copy;
      });
      for (const o of kids(nm, 'lvlOverride')) {
        const i = num(o, 'ilvl') ?? 0;
        const st = num(kid(o, 'startOverride'), 'val');
        if (st !== null && levels[i]) levels[i].start = st;
      }
      this.lists.push({ id, preset: `docx-${id}`, levels });
    }
  }

  readTheme() {
    const t = this.xml('word/theme/theme1.xml');
    if (!t) return;
    const font = (major: boolean) => t.getElementsByTagNameNS('*', major ? 'majorFont' : 'minorFont')[0];
    const pick = (major: boolean) => {
      const f = font(major);
      if (!f) return null;
      const latin = kid(f, 'latin')?.getAttribute('typeface') || '';
      const ea = kid(f, 'ea')?.getAttribute('typeface') || '';
      const hans = kids(f, 'font').find((x) => x.getAttribute('script') === 'Hans')?.getAttribute('typeface') || '';
      return { latin, ea: ea || hans };
    };
    const scheme = t.getElementsByTagNameNS('*', 'clrScheme')[0];
    const colors: Record<string, string> = {};
    for (const c of kids(scheme)) {
      const v = kid(c, 'srgbClr')?.getAttribute('val') ?? kid(c, 'sysClr')?.getAttribute('lastClr');
      if (v && /^[0-9a-fA-F]{6}$/.test(v)) colors[c.localName] = v.toUpperCase();
    }
    const lns = kids(t.getElementsByTagNameNS('*', 'lnStyleLst')[0], 'ln');
    this.theme = { ...this.theme, colors, lineWidths: lns.map((l) => Number(l.getAttribute('w') ?? 9525) / EMU_PER_PT) };
    const mi = pick(false);
    const ma = pick(true);
    if (mi) this.theme = { ...this.theme, minorLatin: mi.latin || this.theme.minorLatin, minorEA: mi.ea || this.theme.minorEA };
    if (ma) this.theme = { ...this.theme, majorLatin: ma.latin || this.theme.majorLatin, majorEA: ma.ea || this.theme.majorEA };
  }

  themeFont(v: string | null, ea: boolean): string | null {
    if (!v) return null;
    const major = v.startsWith('major');
    if (ea || v.endsWith('EastAsia')) return major ? this.theme.majorEA : this.theme.minorEA;
    return major ? this.theme.majorLatin : this.theme.minorLatin;
  }

  applyRPr(base: RunProps, rPr: Element | null): RunProps {
    if (!rPr) return base;
    const r = { ...base };
    const f = kid(rPr, 'rFonts');
    if (f) {
      const ea = attr(f, 'eastAsia') || this.themeFont(attr(f, 'eastAsiaTheme'), true);
      const latin = attr(f, 'ascii') || attr(f, 'hAnsi') || this.themeFont(attr(f, 'asciiTheme') ?? attr(f, 'hAnsiTheme'), false);
      if (ea) r.fontEA = ea;
      if (latin) r.fontLatin = latin;
    }
    const sz = num(kid(rPr, 'sz'), 'val');
    if (sz !== null) r.size = sz / 2;
    const b = onoff(kid(rPr, 'b'));
    if (b !== null) r.bold = b;
    const i = onoff(kid(rPr, 'i'));
    if (i !== null) r.italic = i;
    const u = kid(rPr, 'u');
    if (u) r.underline = attr(u, 'val') !== 'none';
    const c = attr(kid(rPr, 'color'), 'val');
    if (c) r.color = /^[0-9a-fA-F]{6}$/.test(c) ? c.toUpperCase() : 'auto';
    const st = onoff(kid(rPr, 'strike')) ?? onoff(kid(rPr, 'dstrike'));
    if (st !== null) r.strike = st || undefined;
    const va = attr(kid(rPr, 'vertAlign'), 'val');
    if (va) r.vertAlign = va === 'superscript' || va === 'subscript' ? va : undefined;
    const spc = num(kid(rPr, 'spacing'), 'val');
    if (spc !== null) r.spacing = spc ? spc / 20 : undefined;
    const hl = attr(kid(rPr, 'highlight'), 'val');
    const shd = attr(kid(rPr, 'shd'), 'fill');
    if (hl) r.highlight = hl === 'none' ? undefined : (HIGHLIGHTS[hl] ?? r.highlight);
    else if (shd && /^[0-9a-fA-F]{6}$/.test(shd) && shd.toUpperCase() !== 'FFFFFF') r.highlight = shd.toUpperCase();
    return r;
  }

  applyPPr(base: ParaProps, pPr: Element | null, runSize: number): ParaProps {
    if (!pPr) return base;
    const p = { ...base, line: { ...base.line } };
    const jc = attr(kid(pPr, 'jc'), 'val');
    if (jc) {
      const map: Record<string, Align> = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'both', justify: 'both', distribute: 'distribute' };
      p.align = map[jc] ?? p.align;
    }
    // 编号：Word 的缩进按“样式 → 编号级别自带的缩进 → 本层写的缩进”逐层覆盖，所以先处理编号
    const numPr = kid(pPr, 'numPr');
    if (numPr) {
      const id = num(kid(numPr, 'numId'), 'val');
      const level = num(kid(numPr, 'ilvl'), 'val') ?? p.num?.level ?? 0;
      if (id === 0) p.num = null;
      else if (id !== null && this.lists.some((l) => l.id === id)) p.num = { id, level: Math.min(8, Math.max(0, level)) };
      else if (id === null && p.num) p.num = { ...p.num, level };
      const lv = p.num ? this.lists.find((l) => l.id === p.num!.id)?.levels[p.num.level] : undefined;
      if (lv && this.levelInd.has(lv)) {
        p.indentLeft = lv.indentLeft;
        p.firstLine = lv.firstLine;
        if (lv.firstLineChars !== null) p.firstLineChars = lv.firstLineChars;
      }
    }
    const ind = kid(pPr, 'ind');
    if (ind) {
      const left = num(ind, 'left') ?? num(ind, 'start');
      const right = num(ind, 'right') ?? num(ind, 'end');
      const leftChars = num(ind, 'leftChars');
      // 同时写了磅值和字符数时 Word 按文件里的磅值排版（它保存时已按当时的字符宽度算好）；只有字符数时才换算
      if (left !== null) p.indentLeft = left / 20;
      else if (leftChars !== null) p.indentLeft = (leftChars / 100) * runSize;
      if (right !== null) p.indentRight = right / 20;
      const fl = num(ind, 'firstLine');
      const hg = num(ind, 'hanging');
      const flc = num(ind, 'firstLineChars');
      const hgc = num(ind, 'hangingChars');
      if (fl !== null) {
        p.firstLine = fl / 20;
        p.firstLineChars = null;
      }
      if (hg !== null) {
        p.firstLine = -hg / 20;
        p.firstLineChars = null;
      }
      // 字符数为 0 表示“不按字符数缩进”，改用磅值（firstLine / hanging）
      if (flc !== null) p.firstLineChars = flc === 0 ? null : flc / 100;
      if (hgc !== null) p.firstLineChars = hgc === 0 ? null : -hgc / 100;
    }
    const sp = kid(pPr, 'spacing');
    if (sp) {
      const b = num(sp, 'before');
      const a = num(sp, 'after');
      const bl = num(sp, 'beforeLines');
      const al = num(sp, 'afterLines');
      if (b !== null) {
        p.spaceBefore = b / 20;
        p.spaceBeforeLines = null;
      }
      if (a !== null) {
        p.spaceAfter = a / 20;
        p.spaceAfterLines = null;
      }
      if (bl !== null) p.spaceBeforeLines = bl / 100;
      if (al !== null) p.spaceAfterLines = al / 100;
      const on = (v: string | null) => (v === null ? null : !['0', 'false', 'off'].includes(v));
      const bAuto = on(attr(sp, 'beforeAutospacing'));
      const aAuto = on(attr(sp, 'afterAutospacing'));
      if (bAuto !== null) {
        if (bAuto) p.spaceBeforeAuto = true;
        else delete p.spaceBeforeAuto;
      }
      if (aAuto !== null) {
        if (aAuto) p.spaceAfterAuto = true;
        else delete p.spaceAfterAuto;
      }
      const line = num(sp, 'line');
      const rule = attr(sp, 'lineRule') ?? 'auto';
      if (line !== null) p.line = rule === 'exact' ? { rule: 'exact', value: line / 20 } : rule === 'atLeast' ? { rule: 'atLeast', value: line / 20 } : { rule: 'auto', value: line / 240 };
    }
    const flags: [keyof ParaProps, string][] = [
      ['snapToGrid', 'snapToGrid'],
      ['keepNext', 'keepNext'],
      ['keepLines', 'keepLines'],
      ['widowControl', 'widowControl'],
      ['pageBreakBefore', 'pageBreakBefore'],
    ];
    for (const [key, name] of flags) {
      const v = onoff(kid(pPr, name));
      if (v !== null) (p[key] as boolean) = v;
    }
    // 制表位：样式里的和段落里的按位置合并，val="clear" 清掉同一位置的
    const tabsEl = kid(pPr, 'tabs');
    if (tabsEl) {
      let list = [...(p.tabs ?? [])];
      for (const t of kids(tabsEl, 'tab')) {
        const pos = (num(t, 'pos') ?? 0) / 20;
        const val = attr(t, 'val') ?? 'left';
        list = list.filter((x) => Math.abs(x.pos - pos) > 0.05);
        if (val === 'clear' || val === 'bar' || val === 'num') continue;
        const lead = attr(t, 'leader') ?? 'none';
        list.push({
          pos,
          // 小数点对齐按右对齐近似
          align: val === 'right' || val === 'end' || val === 'decimal' ? 'right' : val === 'center' ? 'center' : 'left',
          leader: lead === 'dot' || lead === 'middleDot' ? 'dot' : lead === 'hyphen' ? 'hyphen' : lead === 'underscore' || lead === 'heavy' ? 'underscore' : 'none',
        });
      }
      list.sort((a, b) => a.pos - b.pos);
      p.tabs = list.length ? list : undefined;
      if (!p.tabs) delete p.tabs;
    }
    // 段落边框：样式和段落里的按边合并，val = nil / none 去掉
    const bdr = kid(pPr, 'pBdr');
    if (bdr) {
      const b = { ...(p.borders ?? {}) };
      for (const side of ['top', 'bottom', 'left', 'right'] as const) {
        const el = kid(bdr, side);
        if (!el) continue;
        const v = attr(el, 'val');
        if (!v || v === 'nil' || v === 'none') delete b[side];
        else b[side] = { w: Math.max(0.25, (num(el, 'sz') ?? 4) / 8), space: num(el, 'space') ?? 0, color: attr(el, 'color') ?? 'auto' };
      }
      p.borders = b;
      if (!Object.keys(b).length) delete p.borders;
    }
    const fr = kid(pPr, 'framePr');
    if (fr) {
      const xAlign = attr(fr, 'xAlign');
      p.frame = { hAnchor: attr(fr, 'hAnchor') ?? 'text', vAnchor: attr(fr, 'vAnchor') ?? 'margin', x: (num(fr, 'x') ?? 0) / 20, y: (num(fr, 'y') ?? 0) / 20, ...(xAlign ? { xAlign } : {}) };
    }
    const ol = num(kid(pPr, 'outlineLvl'), 'val');
    if (ol !== null) {
      if (ol >= 0 && ol <= 8) p.outline = ol;
      else delete p.outline;
    }
    const sect = kid(pPr, 'sectPr');
    if (sect) {
      // 先标记“段后有分节符”；分节符的类型写在下一节的 sectPr 上，readDocx 里再确定
      p.sectionBreak = 'nextPage';
      this.sections.push(sect);
    }
    return p;
  }

  readStyles() {
    const st = this.xml('word/styles.xml');
    this.baseRun = { fontEA: this.theme.minorEA, fontLatin: this.theme.minorLatin, size: 10, bold: false, italic: false, underline: false, color: 'auto' };
    this.basePara = {
      align: 'left',
      indentLeft: 0,
      indentRight: 0,
      firstLine: 0,
      firstLineChars: null,
      spaceBefore: 0,
      spaceAfter: 0,
      spaceBeforeLines: null,
      spaceAfterLines: null,
      line: { rule: 'auto', value: 1 },
      snapToGrid: true,
      keepNext: false,
      keepLines: false,
      widowControl: true,
      pageBreakBefore: false,
      num: null,
      sectionBreak: null,
    };
    if (!st) return;
    const root = st.documentElement;
    const dd = kid(root, 'docDefaults');
    this.baseRun = this.applyRPr(this.baseRun, kid(kid(dd, 'rPrDefault'), 'rPr'));
    this.basePara = this.applyPPr(this.basePara, kid(kid(dd, 'pPrDefault'), 'pPr'), this.baseRun.size);
    for (const s of kids(root, 'style')) {
      const id = attr(s, 'styleId');
      if (!id) continue;
      const borders = borderInfo(kid(kid(s, 'tblPr'), 'tblBorders'));
      this.styles.set(id, {
        el: s,
        basedOn: attr(kid(s, 'basedOn'), 'val'),
        name: (attr(kid(s, 'name'), 'val') ?? '').toLowerCase(),
        pPr: kid(s, 'pPr'),
        rPr: kid(s, 'rPr'),
        tblBorders: borders ? borders.on : null,
        tblBorderSize: borders?.size ?? null,
      });
      if (attr(s, 'type') === 'paragraph' && ['1', 'true', 'on'].includes(attr(s, 'default') ?? '')) this.defaultParaStyle = id;
    }
  }

  chain(id: string | null): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    while (id && !seen.has(id)) {
      seen.add(id);
      const s = this.styles.get(id);
      if (!s) break;
      out.unshift(s);
      id = s.basedOn;
    }
    return out;
  }

  /** 内置的标题 / 目录样式（按样式名识别，中文 Word 的样式 ID 是数字）→ Heading1… / TOC1… */
  knownStyle(id: string | null): string | null {
    if (!id) return null;
    const name = this.styles.get(id)?.name ?? '';
    const h = name.match(/^heading ([1-9])$/) ?? id.match(/^heading([1-9])$/i);
    if (h) return `Heading${h[1]}`;
    const t = name.match(/^toc ([1-9])$/) ?? id.match(/^toc([1-9])$/i);
    return t ? `TOC${t[1]}` : null;
  }

  /** 正在读的表格单元格的表格样式层（整表 + 条件格式的 pPr / rPr，按顺序）：在文档缺省之后、段落样式之前 */
  tableLayer: { pPr: Element | null; rPr: Element | null }[] | null = null;

  styleProps(id: string | null): { run: RunProps; para: ParaProps } {
    let run = this.baseRun;
    let para = this.basePara;
    for (const l of this.tableLayer ?? []) {
      run = this.applyRPr(run, l.rPr);
      para = this.applyPPr(para, l.pPr, run.size);
    }
    for (const s of this.chain(id ?? this.defaultParaStyle)) {
      run = this.applyRPr(run, s.rPr);
      para = this.applyPPr(para, s.pPr, run.size);
    }
    return { run, para };
  }

  paragraph(pEl: Element): Paragraph {
    const pPr = kid(pEl, 'pPr');
    const styled = this.styleProps(attr(kid(pPr, 'pStyle'), 'val'));
    const props = this.applyPPr(styled.para, pPr, styled.run.size);
    const mark = this.applyRPr(styled.run, kid(pPr, 'rPr'));
    const styleId = attr(kid(pPr, 'pStyle'), 'val');
    if (styleId) props.styleId = styleId;
    const style = this.knownStyle(styleId ?? this.defaultParaStyle);
    if (style) {
      props.style = style;
      // 内置标题样式本身就带大纲级别（标题 n = 级别 n-1），样式里没写时也是
      const h = style.match(/^Heading(\d)$/);
      if (h && props.outline === undefined) props.outline = Number(h[1]) - 1;
    }
    const runs: Run[] = [];
    /** 当前所在的修订（w:ins / w:del 里面） */
    let curRev: Revision | undefined;
    /** 当前所在的超链接 */
    let curLink: RunProps['link'];
    // 复杂域状态：instr 收集指令；result 为显示结果（页码域时跳过结果）
    // sep：separate 之前经过的 w:r 个数（开始部分的原始 XML 用）；nested：结果里又有域；props：begin 所在文字的格式
    let fld: { depth: number; instr: string; phase: 'instr' | 'result'; skip: boolean; code: EqChar[]; at: number; full: string; sep?: number; nested?: boolean; locked?: boolean; props?: RunProps } | null = null;
    /** 正在读的 w:r（域的结束部分的原始 XML 用） */
    let curRunEl: Element | null = null;
    /** 当前复杂域经过的 w:r（没有显示结果、又不是页码域时整个原样保留） */
    let fieldEls: Element[] = [];
    /** 刚结束的、没有显示结果的复杂域的代码 */
    let endedField: EqChar[] | null = null;
    /** 刚结束的邮件合并域：显示结果从 runs 的第 at 项开始 */
    let mergeEnded: { merge?: NonNullable<Opaque['merge']>; rule?: string; at: number } | null = null;

    const pushText = (text: string, rp: RunProps) => {
      if (!text) return;
      if (fld && (fld.phase === 'instr' || fld.skip)) return;
      runs.push({ text, props: rp });
    };
    const fieldKind = (instr: string) => {
      const k = instr.trim().split(/\s+/)[0]?.toUpperCase();
      return k === 'PAGE' ? 'PAGE' : k === 'NUMPAGES' ? 'NUMPAGES' : null;
    };
    /** 原样对象：保存时按原始 XML 写回（不在复杂域的指令部分里时才记） */
    const pushOpaque = (o: Opaque, rp: RunProps) => {
      if (fld && (fld.phase === 'instr' || fld.skip)) return;
      // 引用了别的部件（图表、图片、OLE 数据）的：字节级保留保存时原样保留，整篇重写时会丢（那时引用的部件不在了）
      if (REL_REF.test(o.xml)) this.lossy.add('图片、文本框或图形');
      runs.push({ text: FIELD_CHAR, props: rp, opaque: o });
    };

    const walk = (el: Element) => {
      for (const c of kids(el)) {
        switch (c.localName) {
          case 'r':
            readRun(c);
            break;
          case 'hyperlink': {
            const outer = curLink;
            const rid = c.getAttributeNS(R_NS, 'id') ?? c.getAttribute('r:id');
            const anchor = attr(c, 'anchor');
            const url = rid ? this.imageRels?.get(rid) : undefined;
            curLink = rid || anchor ? { ...(rid ? { rid } : {}), ...(anchor ? { anchor } : {}), ...(url ? { url } : {}) } : outer;
            walk(c);
            curLink = outer;
            break;
          }
          case 'smartTag':
          case 'customXml':
            walk(c);
            break;
          case 'ins':
          case 'moveTo':
          case 'del':
          case 'moveFrom': {
            // 修订：里面的文字带上插入 / 删除的标记（修订人、时间）
            const outer = curRev;
            curRev = { kind: c.localName === 'ins' || c.localName === 'moveTo' ? 'ins' : 'del', author: attr(c, 'author') ?? '', date: attr(c, 'date') ?? '' };
            walk(c);
            curRev = outer;
            break;
          }
          case 'sdt':
            walk(kid(c, 'sdtContent') ?? c);
            break;
          case 'fldSimple': {
            const kind = fieldKind(attr(c, 'instr') ?? '');
            const rp = this.applyRPr(mark, kid(kid(c, 'r'), 'rPr'));
            const merge = parseMergeInstr(attr(c, 'instr') ?? '');
            const ruleView = merge ? null : ruleFieldView(attr(c, 'instr') ?? '', rp);
            if (merge) pushOpaque({ xml: serialize(c), level: 'para', kind: 'other', w: 0, h: 0, ...mergeFieldView(merge, rp) }, rp);
            else if (ruleView) pushOpaque({ xml: serialize(c), level: 'para', kind: 'other', w: 0, h: 0, ...ruleView }, rp);
            else if (kind) runs.push({ text: FIELD_CHAR, props: rp, field: kind });
            else {
              // 一般的简单域：开始 / 结束标记 + 结果（保存时写成复杂域）
              const id = newFieldId();
              const locked = onoffAttr(c, 'fldLock');
              runs.push({ text: FIELD_CHAR, props: rp, fc: { kind: 'begin', id, instr: (attr(c, 'instr') ?? '').trim(), ...(locked ? { locked } : {}) } });
              walk(c);
              runs.push({ text: FIELD_CHAR, props: rp, fc: { kind: 'end', id } });
            }
            break;
          }
          case 'commentRangeStart':
          case 'commentRangeEnd': {
            const id = attr(c, 'id');
            if (id !== null) runs.push({ text: FIELD_CHAR, props: mark, mark: { kind: c.localName === 'commentRangeStart' ? 'cs' : 'ce', id } });
            break;
          }
          case 'pPr':
          case 'proofErr':
            break;
          case 'oMath':
          case 'oMathPara': {
            // 公式（OMML）：原样保留；显示大小先按公式文字估算
            const text = Array.from(c.getElementsByTagNameNS('*', 't'))
              .map((t) => t.textContent ?? '')
              .join('');
            const size = mark.size;
            pushOpaque({ xml: serialize(c), level: 'para', kind: 'math', text, w: Math.max(size, [...text].length * size * 0.6), h: size * 1.3 }, mark);
            break;
          }
          default:
            // 书签、权限范围等其它段落级标记：不显示，原样保留
            pushOpaque({ xml: serialize(c), level: 'para', kind: c.localName.startsWith('bookmark') ? 'bookmark' : 'other', w: 0, h: 0 }, mark);
            break;
        }
      }
    };

    const readRun = (r: Element) => {
      const rPr = kid(r, 'rPr');
      const rStyle = attr(kid(rPr, 'rStyle'), 'val');
      let rp = styled.run;
      for (const s of this.chain(rStyle)) rp = this.applyRPr(rp, s.rPr);
      rp = this.applyRPr(rp, rPr);
      if (rStyle) rp = { ...rp, rStyle };
      if (curRev) rp = { ...rp, rev: curRev };
      if (curLink) rp = { ...rp, link: curLink };
      const inField = !!fld;
      curRunEl = r;
      readRunKids(r, rp);
      if (inField || fld) fieldEls.push(r);
      if (!fld) {
        if (endedField) {
          // 没有显示结果的域（EQ 带圈字符等）：组成它的各个 w:r 原样保留，EQ 域按域代码显示
          const code = endedField;
          const eq = parseEq(code);
          const first = fieldEls[0];
          const fp = first ? this.applyRPr(styled.run, kid(first, 'rPr')) : rp;
          pushOpaque({ xml: fieldEls.map(serialize).join(''), level: 'para', kind: 'other', w: 0, h: 0, ...(eq ? { eq, text: eqText(eq) } : {}) }, curRev || curLink ? { ...fp, ...(curRev ? { rev: curRev } : {}), ...(curLink ? { link: curLink } : {}) } : fp);
        }
        endedField = null;
        if (mergeEnded) {
          // 邮件合并域：显示结果（«字段名»）那几段文字换成一个整体，组成它的 w:r 原样保留
          const me: { merge?: NonNullable<Opaque['merge']>; rule?: string; at: number } = mergeEnded;
          const first = fieldEls[0];
          const fp = first ? this.applyRPr(styled.run, kid(first, 'rPr')) : rp;
          runs.length = Math.min(runs.length, me.at);
          const view = me.merge ? mergeFieldView(me.merge, fp) : ruleFieldView(me.rule ?? '', fp);
          pushOpaque({ xml: fieldEls.map(serialize).join(''), level: 'para', kind: 'other', w: 0, h: 0, ...view }, curRev || curLink ? { ...fp, ...(curRev ? { rev: curRev } : {}), ...(curLink ? { link: curLink } : {}) } : fp);
          mergeEnded = null;
        }
        fieldEls = [];
      }
    };

    const readRunKids = (parent: Element, rp: RunProps) => {
      for (const c of kids(parent)) {
        switch (c.localName) {
          case 'AlternateContent': {
            // 兼容性标记：优先读新格式（Choice），没有再读旧格式（Fallback）。
            // Choice 里的绘图显示不了（墨迹 w14:contentPart 等）而 Fallback 里有图片时，用 Fallback 的图片显示，
            // 保存时原样写回整个 AlternateContent（墨迹本身逐字节保留）
            const choice = kid(c, 'Choice');
            const fd = kid(kid(c, 'Fallback'), 'drawing');
            const cd = kid(choice, 'drawing');
            if (cd && fd && !(kid(cd, 'inline') ? kid(cd, 'inline')?.getElementsByTagNameNS('*', 'blip').length : this.textBox(cd) || this.floatImage(cd))) {
              const raw = serialize(c);
              const image = this.inlineImage(fd);
              if (image) {
                runs.push({ text: FIELD_CHAR, props: rp, image: { ...image, raw } });
                break;
              }
              const box = this.floatImage(fd);
              if (box) {
                runs.push({ text: FIELD_CHAR, props: rp, box: { ...box, raw } });
                break;
              }
            }
            const pick = choice ?? kid(c, 'Fallback');
            if (pick) readRunKids(pick, rp);
            break;
          }
          case 't':
          case 'delText':
            pushText(c.textContent ?? '', rp);
            break;
          case 'tab':
            pushText('\t', rp);
            break;
          case 'br':
            pushText(attr(c, 'type') === 'page' ? '\f' : attr(c, 'type') === 'column' ? '\v' : '\n', rp);
            break;
          case 'cr':
            pushText('\n', rp);
            break;
          case 'noBreakHyphen':
            pushText('-', rp);
            break;
          case 'sym': {
            // w:sym：符号字体（Wingdings 等）里的一个字符。保留字体，字符放在私用区 U+F0xx（符号字体的字形就在那里）
            const code = parseInt(attr(c, 'char') ?? '', 16);
            const font = attr(c, 'font');
            if (!Number.isFinite(code)) break;
            if (font) pushText(String.fromCharCode(code < 0x100 ? 0xf000 + code : code), { ...rp, fontLatin: font });
            else pushText(String.fromCharCode(code > 0xf000 ? code - 0xf000 : code), rp);
            break;
          }
          case 'drawing': {
            const skip = !!(fld && (fld.phase === 'instr' || fld.skip));
            const image = this.inlineImage(c);
            const box = image ? null : (this.textBox(c) ?? this.floatImage(c));
            if (image) {
              if (!skip) runs.push({ text: FIELD_CHAR, props: rp, image });
            } else if (box) {
              if (!skip) runs.push({ text: FIELD_CHAR, props: rp, box: { ...box, raw: serialize(c) } });
            } else {
              // 形状、图表、SmartArt、浏览器不能显示的图片格式等：原样保留；嵌入式的按大小占位，浮动的不占位置
              const inline = kid(c, 'inline');
              const ext = kid(inline, 'extent');
              const w = ext ? Number(ext.getAttribute('cx')) / EMU_PER_PT : 0;
              const h = ext ? Number(ext.getAttribute('cy')) / EMU_PER_PT : 0;
              pushOpaque({ xml: serialize(c), level: 'run', kind: 'drawing', w: w > 0 && h > 0 ? w : 0, h: w > 0 && h > 0 ? h : 0 }, rp);
            }
            break;
          }
          case 'pict':
          case 'object': {
            // VML 图形、OLE 对象（如 Visio、公式 3.0）：原样保留，按 v:shape 的大小占位
            const shape = c.getElementsByTagNameNS('*', 'shape')[0];
            const style = shape?.getAttribute('style') ?? '';
            const dim = (n: string) => {
              const m = style.match(new RegExp(`(?:^|;)\\s*${n}\\s*:\\s*([\\d.]+)(pt|in|cm|mm|px)?`));
              if (!m) return 0;
              const v = Number(m[1]);
              return m[2] === 'in' ? v * 72 : m[2] === 'cm' ? (v * 72) / 2.54 : m[2] === 'mm' ? (v * 72) / 25.4 : m[2] === 'px' ? v * 0.75 : v;
            };
            const anchored = /position\s*:\s*absolute/.test(style);
            pushOpaque({ xml: serialize(c), level: 'run', kind: 'object', w: anchored ? 0 : dim('width'), h: anchored ? 0 : dim('height') }, rp);
            break;
          }
          case 'footnoteReference':
          case 'endnoteReference':
            // 脚注 / 尾注的引用：原样保留，显示一个小占位
            pushOpaque({ xml: serialize(c), level: 'run', kind: 'note', w: rp.size * 0.5, h: rp.size * 0.6, note: { kind: c.localName === 'footnoteReference' ? 'foot' : 'end', id: attr(c, 'id') ?? '' } }, rp);
            break;
          case 'footnoteRef':
          case 'endnoteRef':
            // 脚注 / 尾注正文开头的本条编号
            pushOpaque({ xml: serialize(c), level: 'run', kind: 'note', w: rp.size * 0.5, h: rp.size * 0.6, note: { kind: c.localName === 'footnoteRef' ? 'foot' : 'end', id: '@self' } }, rp);
            break;
          case 'commentReference':
            // 批注的引用标记：保存时按批注范围重新生成
            break;
          case 'fldChar': {
            const t = attr(c, 'fldCharType');
            if (t === 'begin') {
              if (fld) {
                fld.depth++;
                if (fld.phase === 'result') fld.nested = true;
                else fld.full += '{';
              } else fld = { depth: 1, instr: '', phase: 'instr', skip: false, code: [], at: runs.length, full: '', props: rp, ...(onoffAttr(c, 'fldLock') ? { locked: true } : {}) };
            } else if (t === 'separate' && fld && fld.depth === 1) {
              const kind = fieldKind(fld.instr);
              fld.phase = 'result';
              fld.sep = fieldEls.length;
              if (kind) {
                runs.push({ text: FIELD_CHAR, props: rp, field: kind });
                fld.skip = true;
              }
            } else if (t === 'end' && fld) {
              fld.depth--;
              if (fld.depth > 0 && fld.phase === 'instr') fld.full += '}';
              if (fld.depth === 0) {
                const merge = parseMergeInstr(fld.instr);
                if (merge) mergeEnded = { merge, at: fld.at };
                else if (parseRule(fld.full)) mergeEnded = { rule: fld.full, at: fld.at };
                if (fld.phase === 'instr' && !mergeEnded) {
                  const kind = fieldKind(fld.instr);
                  if (kind) runs.push({ text: FIELD_CHAR, props: rp, field: kind });
                  else endedField = fld.code;
                } else if (fld.phase === 'result' && !fld.skip && !mergeEnded) {
                  if (fld.nested) this.lossy.add('域（保留显示文字）');
                  else {
                    // 一般的域：结果前后放开始 / 结束标记。只由域字符、域代码组成的 w:r 才记原始 XML（域代码没改时原样写回）
                    const id = newFieldId();
                    // separate 和 end 在同一个 w:r 里时不记（开始、结束部分分不开）
                    const head = fld.sep !== undefined && fld.sep < fieldEls.length ? fieldEls.slice(0, fld.sep + 1) : [];
                    const pure = (els: Element[], ok: Set<string>) => els.length > 0 && els.every((el) => kids(el).every((k) => ok.has(k.localName)));
                    const raw = pure(head, HEAD_KIDS) ? head.map(serialize).join('') : undefined;
                    const tail = curRunEl && head.length && pure([curRunEl], TAIL_KIDS) ? serialize(curRunEl) : undefined;
                    runs.splice(fld.at, 0, { text: FIELD_CHAR, props: fld.props ?? rp, fc: { kind: 'begin', id, instr: fld.full.trim(), ...(raw && tail ? { raw } : {}), ...(fld.locked ? { locked: true } : {}) } });
                    runs.push({ text: FIELD_CHAR, props: rp, fc: { kind: 'end', id, ...(raw && tail ? { raw: tail } : {}) } });
                  }
                }
                fld = null;
              }
            }
            break;
          }
          case 'instrText':
            if (fld) fld.full += c.textContent ?? '';
            if (fld && fld.depth === 1) {
              fld.instr += c.textContent ?? '';
              for (const ch of c.textContent ?? '') fld.code.push({ ch, props: rp });
            }
            break;
          default:
            break;
        }
      }
    };

    walk(pEl);
    // 跨段的域（目录等）：只保留显示文字
    if (fld) this.lossy.add('域（保留显示文字）');
    // 合并相邻同格式文本
    const merged: Run[] = [];
    for (const r of runs) {
      const last = merged[merged.length - 1];
      if (last && !isAtom(last) && !isAtom(r) && JSON.stringify(last.props) === JSON.stringify(r.props)) last.text += r.text;
      else merged.push({ ...r });
    }
    const para: Paragraph = { kind: 'p', props, runs: merged, mark };
    const sect = kid(pPr, 'sectPr');
    if (sect) this.sectOf.set(para, sect);
    return para;
  }

  table(tbl: Element): Table {
    const tblPr = kid(tbl, 'tblPr');
    const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => (num(g, 'w') ?? 0) / 20);
    const own = borderInfo(kid(tblPr, 'tblBorders'));
    let borders: boolean | null = own ? own.on : null;
    let borderSize: number | null = own?.size ?? null;
    if (borders === null) {
      for (const s of this.chain(attr(kid(tblPr, 'tblStyle'), 'val')).reverse()) {
        if (s.tblBorders !== null) {
          borders = s.tblBorders;
          borderSize = s.tblBorderSize;
          break;
        }
      }
    }
    const mar = kid(tblPr, 'tblCellMar');
    const m = (n: string, d: number) => {
      const v = num(kid(mar, n), 'w') ?? num(kid(mar, n === 'left' ? 'start' : n === 'right' ? 'end' : n), 'w');
      return v === null ? d : v / 20;
    };
    const jc = attr(kid(tblPr, 'jc'), 'val');
    // 表格样式（OOXML 17.7.6）：整表的格式 + 按 tblLook 开关的条件格式（镶边列 → 镶边行 → 首末列 → 首末行 → 四角）
    const chain = this.chain(attr(kid(tblPr, 'tblStyle'), 'val'));
    const look = tableLook(kid(tblPr, 'tblLook'));
    const styleTblPr = (name: string) => chain.map((s) => kid(kid(s.el, 'tblPr'), name)).filter((e): e is Element => !!e).pop() ?? null;
    const rowBand = Math.max(1, num(styleTblPr('tblStyleRowBandSize'), 'val') ?? 1);
    const colBand = Math.max(1, num(styleTblPr('tblStyleColBandSize'), 'val') ?? 1);
    const cond = (type: string) => chain.flatMap((s) => kids(s.el, 'tblStylePr').filter((t) => attr(t, 'type') === type));
    const tblEdges = mergeBorders([...chain.map((s) => kid(kid(s.el, 'tblPr'), 'tblBorders')), kid(tblPr, 'tblBorders')]);
    const trs = kids(tbl, 'tr');
    const nRows = trs.length;
    const rows: Row[] = trs.map((tr, ri) => {
      const trPr = kid(tr, 'trPr');
      const h = kid(trPr, 'trHeight');
      const tcs = kids(tr, 'tc');
      const cells: Cell[] = tcs.map((tc, ci) => {
        const tcPr = kid(tc, 'tcPr');
        const vm = kid(tcPr, 'vMerge');
        if (kids(tc, 'tbl').length) this.lossy.add('嵌套表格');
        // 这个单元格适用的条件格式
        const firstR = look.firstRow && ri === 0;
        const lastR = look.lastRow && ri === nRows - 1 && nRows > 1;
        const firstC = look.firstColumn && ci === 0;
        const lastC = look.lastColumn && ci === tcs.length - 1 && tcs.length > 1;
        const types: { type: string; kind: 'row' | 'col' | 'cell' }[] = [];
        if (!look.noVBand && !firstC && !lastC) types.push({ type: Math.floor((ci - (look.firstColumn ? 1 : 0)) / colBand) % 2 ? 'band2Vert' : 'band1Vert', kind: 'col' });
        if (!look.noHBand && !firstR && !lastR) types.push({ type: Math.floor((ri - (look.firstRow ? 1 : 0)) / rowBand) % 2 ? 'band2Horz' : 'band1Horz', kind: 'row' });
        if (firstC) types.push({ type: 'firstCol', kind: 'col' });
        if (lastC) types.push({ type: 'lastCol', kind: 'col' });
        if (firstR) types.push({ type: 'firstRow', kind: 'row' });
        if (lastR) types.push({ type: 'lastRow', kind: 'row' });
        if (firstR && firstC) types.push({ type: 'nwCell', kind: 'cell' });
        if (firstR && lastC) types.push({ type: 'neCell', kind: 'cell' });
        if (lastR && firstC) types.push({ type: 'swCell', kind: 'cell' });
        if (lastR && lastC) types.push({ type: 'seCell', kind: 'cell' });
        const parts = types.flatMap((t) => cond(t.type).map((el) => ({ el, kind: t.kind })));
        this.tableLayer = [...chain.map((s) => ({ pPr: s.pPr, rPr: s.rPr })), ...parts.map((p) => ({ pPr: kid(p.el, 'pPr'), rPr: kid(p.el, 'rPr') }))];
        const paragraphs = kids(tc, 'p').map((p) => this.paragraph(p));
        this.tableLayer = null;
        const tcb = kid(tcPr, 'tcBorders');
        // 底纹：单元格自己的 > 条件格式（后面的覆盖前面的）> 整表
        const shdFill = (el: Element | null) => {
          const f = attr(kid(el, 'shd'), 'fill');
          return f && /^[0-9a-f]{6}$/i.test(f) ? f.toUpperCase() : null;
        };
        let fill = shdFill(tcPr);
        if (!fill) for (const el of [...chain.map((s) => kid(s.el, 'tcPr')), ...parts.map((p) => kid(p.el, 'tcPr'))]) fill = shdFill(el) ?? fill;
        // 四条边：表格边框按位置取外框或内框 → 条件格式的 tcBorders → 单元格自己的
        const isTop = ri === 0;
        const isBottom = ri === nRows - 1;
        const isLeft = ci === 0;
        const isRight = ci === tcs.length - 1;
        const edges: CellEdges = {
          top: (isTop ? tblEdges.top : tblEdges.insideH) ?? null,
          bottom: (isBottom ? tblEdges.bottom : tblEdges.insideH) ?? null,
          left: (isLeft ? tblEdges.left : tblEdges.insideV) ?? null,
          right: (isRight ? tblEdges.right : tblEdges.insideV) ?? null,
        };
        const apply = (b: BorderSet, kind: 'row' | 'col' | 'cell') => {
          const pick = (side: keyof CellEdges, outer: boolean, inside: 'insideH' | 'insideV') => {
            const v = kind === 'cell' || outer ? b[side] : b[inside];
            if (v !== undefined) edges[side] = v;
          };
          pick('top', kind !== 'col' || isTop, 'insideH');
          pick('bottom', kind !== 'col' || isBottom, 'insideH');
          pick('left', kind !== 'row' || isLeft, 'insideV');
          pick('right', kind !== 'row' || isRight, 'insideV');
        };
        for (const p of parts) apply(mergeBorders([kid(kid(p.el, 'tcPr'), 'tcBorders')]), p.kind);
        if (tcb) apply(mergeBorders([tcb]), 'cell');
        return {
          ...(fill && { shading: fill }),
          ...(chain.length || tcb ? { edges } : {}),
          borders: tcb ? kids(tcb).some((b) => !['none', 'nil'].includes(attr(b, 'val') ?? 'none')) : undefined,
          gridSpan: num(kid(tcPr, 'gridSpan'), 'val') ?? 1,
          vMerge: vm ? (attr(vm, 'val') === 'restart' ? 'restart' : 'continue') : 'none',
          vAlign: (attr(kid(tcPr, 'vAlign'), 'val') as Cell['vAlign']) ?? 'top',
          paragraphs: paragraphs.length ? paragraphs : [{ kind: 'p', props: this.basePara, runs: [], mark: this.baseRun }],
        };
      });
      const rule = attr(h, 'hRule');
      return {
        height: (num(h, 'val') ?? 0) / 20,
        heightRule: rule === 'exact' ? 'exact' : rule === 'auto' || !h ? 'auto' : 'atLeast',
        header: onoff(kid(trPr, 'tblHeader')) ?? false,
        cells,
      };
    });
    // 缺少网格定义时按单元格数平均分配
    const nCols = Math.max(cols.length, ...rows.map((r) => r.cells.reduce((n, c) => n + c.gridSpan, 0)));
    while (cols.length < nCols) cols.push(cols[cols.length - 1] ?? 100);
    return {
      kind: 'table',
      cols,
      rows,
      align: jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : 'left',
      indent: (num(kid(tblPr, 'tblInd'), 'w') ?? 0) / 20,
      borders: borders ?? false,
      ...(borderSize !== null && borderSize !== 0.5 ? { borderSize } : {}),
      cellMargin: { left: m('left', 5.4), right: m('right', 5.4), top: m('top', 0), bottom: m('bottom', 0) },
      ...(tableFloat(tblPr) ? { float: tableFloat(tblPr) } : {}),
      ...(tablePct(tblPr) ? { widthPct: tablePct(tblPr) } : {}),
    };
  }

  blocks(container: Element): Block[] {
    const out: Block[] = [];
    for (const c of kids(container)) {
      if (c.localName === 'p') out.push(this.paragraph(c));
      else if (c.localName === 'tbl') out.push(this.table(c));
      else if (c.localName === 'sdt') out.push(...this.blocks(kid(c, 'sdtContent') ?? c));
    }
    return out;
  }

  /** 节属性中某类页眉页脚引用的文件路径；没有引用（与上一节相同）时返回 null。 */
  hfTarget(rels: Map<string, string>, sectPr: Element | null, kind: 'header' | 'footer', type: 'default' | 'first' | 'even'): string | null {
    const ref = kids(sectPr, `${kind}Reference`).find((r) => (attr(r, 'type') ?? 'default') === type);
    if (!ref) return null;
    const rid = ref.getAttributeNS(R_NS, 'id') ?? ref.getAttribute('r:id');
    const target = rid ? rels.get(rid) : undefined;
    return target ? `word/${target.replace(/^\/?word\//, '')}` : '';
  }

  headerFooter(path: string): Paragraph[] {
    const d = path ? this.xml(path) : null;
    if (!d) return [];
    // 图片按这个部件自己的关系文件找（word/_rels/header1.xml.rels）
    const rels = this.imageRels;
    const dir = path.slice(0, path.lastIndexOf('/') + 1);
    const own = this.xml(`${dir}_rels/${path.slice(dir.length)}.rels`);
    const map = new Map<string, string>();
    for (const r of Array.from(own?.documentElement.children ?? [])) map.set(r.getAttribute('Id') ?? '', r.getAttribute('Target') ?? '');
    this.imageRels = map;
    const blocks = this.blocks(d.documentElement);
    this.imageRels = rels;
    // 表格：占位段落带着表格（只读显示）和原 XML
    const tbls = kids(d.documentElement, 'tbl');
    let ti = 0;
    const ps = blocks.map((b): Paragraph => {
      if (b.kind === 'p') return b;
      const el = tbls[ti++];
      return { kind: 'p', props: { ...this.basePara }, runs: [], mark: { ...this.baseRun }, hfTable: { table: b, xml: el ? serialize(el) : tableXml(b) } };
    });
    if (ps.some((p) => p.runs.some((r) => r.image))) this.hfImages = true;
    for (const p of ps) p.hfPart = path;
    return ps;
  }

  /** 页眉页脚里有图片（整篇重写时图片的关系写不进页眉页脚部件，算会丢失） */
  hfImages = false;
}

/** sectPr 的 w:type：这一节如何开始，也就是它前面那个分节符的类型（缺省为下一页） */
function breakType(sect: Element | null): SectionBreak {
  const t = attr(kid(sect, 'type'), 'val');
  return t === 'continuous' ? 'continuous' : t === 'evenPage' ? 'evenPage' : t === 'oddPage' ? 'oddPage' : 'nextPage';
}

/** 页眉页脚分区在 OOXML 中的类型 */
function hfType(story: HFStory): 'default' | 'first' | 'even' {
  return story.endsWith('First') ? 'first' : story.endsWith('Even') ? 'even' : 'default';
}

/** 原文件的样式表：算样式的格式、找本编辑器的标题 / 目录样式对应的样式 id（字节级保留保存用） */
export interface StyleSource {
  base(pStyle: string | undefined, rStyle?: string): { para: ParaProps; run: RunProps };
  /** Heading1… / TOC1… → 原文件里的样式 id；没有返回 null */
  idOf(canonical: string): string | null;
  /** 样式 id → Heading1… / TOC1…（不是这两类返回 null） */
  canonical(id: string): string | null;
  /** 默认段落样式的 id（“正文”） */
  defaultId(): string | null;
  /** 段落样式和字符样式（样式窗格用）：name 是 styles.xml 里的原名 */
  list(): { id: string; name: string; type: 'paragraph' | 'character'; basedOn: string | null; quick: boolean; custom: boolean }[];
}

export function styleSource(bytes: Uint8Array): StyleSource {
  const rd = new Reader(unzipSync(bytes));
  rd.readTheme();
  rd.readNumbering();
  rd.readStyles();
  const cache = new Map<string, { para: ParaProps; run: RunProps }>();
  return {
    base(pStyle, rStyle) {
      const key = `${pStyle ?? ''}|${rStyle ?? ''}`;
      let hit = cache.get(key);
      if (!hit) {
        const st = rd.styleProps(pStyle ?? null);
        let run = st.run;
        if (rStyle) for (const s of rd.chain(rStyle)) run = rd.applyRPr(run, s.rPr);
        cache.set(key, (hit = { para: st.para, run }));
      }
      return hit;
    },
    idOf(canonical) {
      for (const id of rd.styles.keys()) if (rd.knownStyle(id) === canonical) return id;
      return null;
    },
    canonical: (id) => rd.knownStyle(id),
    defaultId: () => rd.defaultParaStyle,
    list() {
      const out: ReturnType<StyleSource['list']> = [];
      for (const [id, s] of rd.styles) {
        const type = attr(s.el, 'type');
        if (type !== 'paragraph' && type !== 'character') continue;
        if (kid(s.el, 'semiHidden') && !kid(s.el, 'unhideWhenUsed')) continue;
        out.push({
          id,
          name: attr(kid(s.el, 'name'), 'val') ?? id,
          type,
          basedOn: s.basedOn,
          quick: !!kid(s.el, 'qFormat'),
          custom: ['1', 'true', 'on'].includes(attr(s.el, 'customStyle') ?? ''),
        });
      }
      return out;
    },
  };
}

export interface ReadOptions {
  /**
   * 按 docx-engine 扫描出的正文顶层元素读（字节级保留保存用）：第 i 项是第 i 个元素去掉内容控件外壳后的 XML，
   * 不是段落 / 表格的为 null。读出的块记下 src = i。不传时按 DOM 读整个正文。
   */
  elements?: (string | null)[];
}

export function readDocx(bytes: Uint8Array, opts: ReadOptions = {}): Doc {
  const files = unzipSync(bytes);
  const docOpen = files['word/document.xml'] ? (strFromU8(files['word/document.xml']).match(/<w:document\b[^>]*>/)?.[0] ?? '') : '';
  const rd = new Reader(files);
  const main = rd.xml('word/document.xml');
  if (!main) throw new Error('不是有效的 Word 文档（缺少 word/document.xml）');
  rd.readTheme();
  rd.readNumbering();
  rd.readStyles();

  const relsDoc = rd.xml('word/_rels/document.xml.rels');
  const rels = new Map<string, string>();
  for (const r of Array.from(relsDoc?.documentElement.children ?? [])) rels.set(r.getAttribute('Id') ?? '', r.getAttribute('Target') ?? '');

  const body = kid(main.documentElement, 'body')!;
  rd.imageRels = rels;
  let blocks: Block[];
  if (opts.elements) {
    // 每个元素单独解析：沿用 w:document 开始标签上的命名空间声明
    const open = docOpen || `<w:document xmlns:w="${W_NS}">`;
    blocks = [];
    opts.elements.forEach((xml, i) => {
      if (!xml) return;
      const d = new DOMParser().parseFromString(`${open}<w:body>${xml}</w:body></w:document>`, 'application/xml');
      const el = kid(d.documentElement, 'body');
      if (!el || d.getElementsByTagName('parsererror').length) return;
      for (const b of rd.blocks(el)) {
        b.src = i;
        blocks.push(b);
      }
    });
  } else blocks = rd.blocks(body);
  const sectPr = kid(body, 'sectPr');
  const setup = defaultSetup();
  const grid = kid(sectPr, 'docGrid');
  const tw = (el: Element | null, n: string, d: number) => {
    const v = num(el, n);
    return v === null ? d : v / 20;
  };
  // 纸张与页边距（各节分别读取）
  const geomOf = (sp: Element | null): PageGeom => {
    const d = defaultSetup();
    const pgSz = kid(sp, 'pgSz');
    const pgMar = kid(sp, 'pgMar');
    return {
      width: tw(pgSz, 'w', d.width),
      height: tw(pgSz, 'h', d.height),
      marginTop: Math.abs(tw(pgMar, 'top', d.marginTop)),
      marginBottom: Math.abs(tw(pgMar, 'bottom', d.marginBottom)),
      marginLeft: tw(pgMar, 'left', d.marginLeft) + tw(pgMar, 'gutter', 0),
      marginRight: tw(pgMar, 'right', d.marginRight),
      headerDistance: tw(pgMar, 'header', d.headerDistance),
      footerDistance: tw(pgMar, 'footer', d.footerDistance),
    };
  };
  Object.assign(setup, geomOf(sectPr));
  const gt = attr(grid, 'type');
  setup.grid = {
    type: (gt === 'lines' || gt === 'linesAndChars' || gt === 'snapToChars' ? gt : 'none') as GridType,
    linePitch: tw(grid, 'linePitch', 15.6),
    charSpace: num(grid, 'charSpace') ?? 0,
  };
  // 分栏（OOXML：w:space 缺省 720 缇；w:equalWidth 为 0 时按 w:col 的宽度和间距）
  const colsOf = (sp: Element | null): Columns | undefined => {
    const c = kid(sp, 'cols');
    const cs = kids(c, 'col');
    const n = Math.max(1, num(c, 'num') ?? (cs.length || 1));
    if (n <= 1) return undefined;
    const out: Columns = { num: n, space: tw(c, 'space', 36), sep: ['1', 'true', 'on'].includes(attr(c, 'sep') ?? '') };
    if (attr(c, 'equalWidth') === '0' || attr(c, 'equalWidth') === 'false') {
      if (cs.length === n) out.widths = cs.map((x) => ({ w: tw(x, 'w', 0), space: tw(x, 'space', 0) }));
    }
    return out;
  };
  const lastCols = colsOf(sectPr);
  if (lastCols) setup.cols = lastCols;
  const settings = rd.xml('word/settings.xml');
  const dts = num(kid(settings?.documentElement, 'defaultTabStop'), 'val');
  if (dts) setup.defaultTab = dts / 20;
  setup.evenAndOdd = onoff(kid(settings?.documentElement, 'evenAndOddHeaders')) ?? false;
  if (onoff(kid(kid(settings?.documentElement, 'compat'), 'adjustLineHeightInTable'))) setup.tableGrid = true;
  if (/^compressPunctuation/.test(attr(kid(settings?.documentElement, 'characterSpacingControl'), 'val') ?? '')) setup.compressPunct = true;
  if (onoff(kid(settings?.documentElement, 'noPunctuationKerning'))) setup.noPunctKern = true;
  // 兼容模式与“另一套东亚断行规则”：决定左对齐等的行能否压缩标点、行尾标点能否悬挂（layout.ts 的 lineRules）
  const compat = kid(settings?.documentElement, 'compat');
  const mode = kids(compat, 'compatSetting').find((c) => attr(c, 'name') === 'compatibilityMode');
  const modeVal = mode ? Number(attr(mode, 'val')) : NaN;
  if (Number.isFinite(modeVal) && modeVal > 0) setup.compatMode = modeVal;
  if (onoff(kid(compat, 'useAltKinsokuLineBreakRules'))) setup.altKinsoku = true;
  if (onoff(kid(compat, 'doNotExpandShiftReturn'))) setup.noExpandShiftReturn = true;

  const defaults = rd.styleProps(null).run;
  const emptyPara = (): Paragraph => ({ kind: 'p', props: { ...rd.styleProps(null).para }, runs: [], mark: { ...defaults } });
  // 各节的页眉页脚：没有引用的分区与上一节相同（共用数组），同一文件被多节引用时也共用
  const partCache = new Map<string, Paragraph[]>();
  const sectionHF = (sp: Element | null, prev: SectionHF | null): SectionHF => {
    const out = { titlePg: onoff(kid(sp, 'titlePg')) ?? false } as SectionHF;
    for (const story of HF_STORIES) {
      const target = rd.hfTarget(rels, sp, isHeaderStory(story) ? 'header' : 'footer', hfType(story));
      if (target === null) {
        out[story] = prev?.[story] ?? [{ ...emptyPara(), hfNone: true }];
        continue;
      }
      let ps = partCache.get(target);
      if (!ps) {
        ps = rd.headerFooter(target);
        if (!ps.length) ps = [emptyPara()];
        partCache.set(target, ps);
      }
      out[story] = ps;
    }
    return out;
  };
  let prev: SectionHF | null = null;
  // 各节的页眉页脚与页面设置；页面设置与最后一节相同的不单独保存
  const lastGeom = pageGeom(setup);
  const breaks = blocks.filter((b): b is Paragraph => b.kind === 'p' && !!b.props.sectionBreak);
  breaks.forEach((b, i) => {
    // 段后分节符的类型 = 下一节 sectPr 的 w:type（最后一个分节符看正文末尾的 sectPr）
    b.props.sectionBreak = breakType(i + 1 < breaks.length ? (rd.sectOf.get(breaks[i + 1]) ?? null) : sectPr);
  });
  for (const b of blocks) {
    if (b.kind !== 'p' || !b.props.sectionBreak) continue;
    const sp = rd.sectOf.get(b) ?? null;
    b.sectionHF = prev = sectionHF(sp, prev);
    const g = geomOf(sp);
    if (!sameGeom(g, lastGeom)) b.sectionPage = g;
    const cols = colsOf(sp);
    if (cols) b.sectionCols = cols;
  }
  const last = sectionHF(sectPr, prev);
  setup.titlePg = last.titlePg;
  const comments = rd.readComments(rels);
  const footnotes = rd.readFootnotes(rels);
  const watermark = readWatermark(files);
  return {
    setup,
    defaults,
    body: blocks.length ? blocks : [emptyPara()],
    header: last.header,
    footer: last.footer,
    headerFirst: last.headerFirst,
    footerFirst: last.footerFirst,
    headerEven: last.headerEven,
    footerEven: last.footerEven,
    lists: rd.lists,
    lossy: Array.from(rd.lossy),
    ...(comments.length && { comments }),
    ...(Object.keys(footnotes).length ? { footnotes } : {}),
    ...(watermark ? { watermark } : {}),
    ...(onoff(kid(settings?.documentElement, 'trackRevisions')) ? { trackRevisions: true } : {}),
  };
}

/** 页眉里的水印：文字水印（Word 的艺术字 v:textpath）或图片水印（v:imagedata，图片取自页眉自己的关系） */
function readWatermark(files: Record<string, Uint8Array>): Watermark | undefined {
  for (const path of Object.keys(files).sort()) {
    if (!/^word\/[^/]*header[^/]*\.xml$/i.test(path)) continue;
    const xml = strFromU8(files[path]);
    if (xml.includes('<v:textpath')) {
      for (const m of xml.matchAll(/<w:pict[\s>][\s\S]*?<\/w:pict>/g)) {
        const shape = readWatermarkShape(m[0]);
        const art = shape?.wordArt;
        if (!shape || !art) continue;
        return {
          text: art.text,
          colorHex: art.colorHex,
          opacity: art.opacity,
          diagonal: !!shape.rotationDeg && Math.abs(shape.rotationDeg % 360) > 1,
          ...(art.fontFamily ? { fontFamily: art.fontFamily } : {}),
          ...(art.bold ? { bold: true } : {}),
          ...(art.italic ? { italic: true } : {}),
        };
      }
    }
    const pic = xml.includes('<v:imagedata') ? readPictureWatermark(xml) : null;
    if (!pic) continue;
    const relsPath = path.replace(/^word\//, 'word/_rels/') + '.rels';
    const rels = files[relsPath] ? strFromU8(files[relsPath]) : '';
    const tag = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((x) => x[0]).find((t) => t.includes(`Id="${pic.rId}"`));
    const target = tag?.match(/Target="([^"]+)"/)?.[1];
    const media = target ? (target.startsWith('/') ? target.slice(1) : `word/${target}`) : '';
    const mime = imageMime(media);
    if (!media || !mime || !files[media]) continue;
    // 显示的大小就是原始大小的 100%（原始像素按 96 dpi 折回去）
    const widthPx = Math.max(1, Math.round((pic.widthPt * 96) / 72));
    const heightPx = Math.max(1, Math.round((pic.heightPt * 96) / 72));
    return { image: { id: registerImage(files[media], mime), widthPx, heightPx }, scale: 100, washout: pic.washout };
  }
  return undefined;
}

// ---------------- 写入 ----------------

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
const twip = (pt: number) => Math.round(pt * 20);

/** Word 的突出显示颜色名 ↔ RRGGBB */
const HIGHLIGHTS: Record<string, string> = {
  yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', darkBlue: '000080', darkCyan: '008080',
  darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF',
};
const HIGHLIGHT_NAME: Record<string, string> = Object.fromEntries(Object.entries(HIGHLIGHTS).map(([k, v]) => [v, k]));

/**
 * 字节级保留保存时的写出上下文（patch.ts 设置）：生成的段落引用原文件的样式、只写与样式不同的格式，
 * 原样对象、原图片按原始 XML 写回，新图片登记新的关系。为 null 时是整篇重写（新建文档）。
 */
export interface PatchCtx {
  /** 样式算出的格式：段落样式 + 字符样式（都按原文件的样式表） */
  base(pStyle: string | undefined, rStyle?: string): { para: ParaProps; run: RunProps };
  /** 本编辑器的标题 / 目录样式名（Heading1…）→ 原文件里的样式 id；没有这个样式时 patch.ts 会补上 */
  styleIdOf(canonical: string): string;
  /** 新插入的图片：返回关系 id（patch.ts 负责写媒体文件和关系） */
  newImageRel(imageId: string): string | null;
  /** 新加的外部超链接：返回关系 id（patch.ts 负责写关系） */
  newLinkRel(url: string): string;
  /** 改过大小或位置的原文本框 / 浮动图片：原始 XML 与读入时的模型 key 一致就原样写回 */
  boxUnchanged(box: TextBox): boolean;
  /** 分节段落重新生成时：按原 sectPr 改页面设置后的 XML（没有原 sectPr 时返回 null，用整篇写出的节属性） */
  sectPr(p: Paragraph): string | null;
  /** 修订 id 的起点（接在原文件最大的 id 后面，避免重复） */
  revBase: number;
}

let patchCtx: PatchCtx | null = null;
/** 正在写出的段落的样式格式（只写与它不同的） */
let curBase: { para: ParaProps; run: RunProps } | null = null;

export function setPatchCtx(ctx: PatchCtx | null): void {
  patchCtx = ctx;
  curBase = null;
  revSeq = ctx ? ctx.revBase : 0;
}

/** 与样式算出的格式比较，只写不同的项（Word：直接格式覆盖样式，没写的继承样式） */
function rPrDiffXml(p: RunProps, b: RunProps): string {
  let x = '';
  if (p.rStyle) x += `<w:rStyle w:val="${esc(p.rStyle)}"/>`;
  let fonts = '';
  if (p.fontLatin !== b.fontLatin) fonts += ` w:ascii="${esc(p.fontLatin)}" w:hAnsi="${esc(p.fontLatin)}" w:cs="${esc(p.fontLatin)}"`;
  if (p.fontEA !== b.fontEA) fonts += ` w:eastAsia="${esc(p.fontEA)}"`;
  if (fonts) x += `<w:rFonts${fonts}/>`;
  if (p.bold !== b.bold) x += p.bold ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/><w:bCs w:val="0"/>';
  if (p.italic !== b.italic) x += p.italic ? '<w:i/><w:iCs/>' : '<w:i w:val="0"/><w:iCs w:val="0"/>';
  if (!!p.strike !== !!b.strike) x += p.strike ? '<w:strike/>' : '<w:strike w:val="0"/>';
  if (p.color !== b.color) x += `<w:color w:val="${p.color === 'auto' ? 'auto' : p.color}"/>`;
  if ((p.spacing ?? 0) !== (b.spacing ?? 0)) x += `<w:spacing w:val="${Math.round((p.spacing ?? 0) * 20)}"/>`;
  if (p.size !== b.size) x += `<w:sz w:val="${Math.round(p.size * 2)}"/><w:szCs w:val="${Math.round(p.size * 2)}"/>`;
  const hl = (v: string | undefined) => (v && HIGHLIGHT_NAME[v] ? HIGHLIGHT_NAME[v] : null);
  if (p.highlight !== b.highlight) x += hl(p.highlight) ? `<w:highlight w:val="${hl(p.highlight)}"/>` : hl(b.highlight) ? '<w:highlight w:val="none"/>' : '';
  if (p.underline !== b.underline) x += `<w:u w:val="${p.underline ? 'single' : 'none'}"/>`;
  if (p.highlight !== b.highlight && !hl(p.highlight)) x += `<w:shd w:val="clear" w:color="auto" w:fill="${p.highlight ?? 'auto'}"/>`;
  if (p.vertAlign !== b.vertAlign) x += `<w:vertAlign w:val="${p.vertAlign ?? 'baseline'}"/>`;
  return x ? `<w:rPr>${x}</w:rPr>` : '';
}

/** 段前 / 段后“自动”的属性：与样式不同时才写（关掉写 "0"） */
function autoAttr(side: 'before' | 'after', v: boolean | undefined, base: boolean | undefined): string {
  return !!v === !!base ? '' : ` w:${side}Autospacing="${v ? 1 : 0}"`;
}

/** 运行时选用：补丁模式按样式求差，整篇写出时写全 */
function runPr(p: RunProps): string {
  if (!patchCtx || !curBase) return rPrXml(p);
  const base = p.rStyle ? patchCtx.base(curStyle, p.rStyle).run : curBase.run;
  return rPrDiffXml(p, base);
}
let curStyle: string | undefined;

/** CT_PPr 子元素的规定顺序（重新组合原 pPr 时用） */
const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
  'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
  'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
];

/** 原 pPr 的子元素（按出现顺序，名字不带前缀） */
export function pPrChildren(pPrXml: string): [string, string][] {
  const inner = pPrXml.replace(/^<w:pPr\b[^>]*>/, '').replace(/<\/w:pPr>$/, '');
  const out: [string, string][] = [];
  let i = 0;
  while (i < inner.length) {
    const m = inner.slice(i).match(/^<w:([A-Za-z]+)\b/);
    if (!m) break;
    const name = m[1];
    const start = i;
    const head = inner.indexOf('>', i);
    if (inner[head - 1] === '/') {
      i = head + 1;
    } else {
      // 找配对的结束标签（同名元素可能嵌套，如 pPrChange 里的 pPr）
      let depth = 1;
      let j = head + 1;
      const re = new RegExp(`<(/?)w:${name}\\b[^>]*?(/?)>`, 'g');
      re.lastIndex = j;
      let t: RegExpExecArray | null;
      while (depth && (t = re.exec(inner))) {
        if (t[1]) depth--;
        else if (!t[2]) depth++;
        j = t.index + t[0].length;
      }
      i = j;
    }
    out.push([name, inner.slice(start, i)]);
  }
  return out;
}

/**
 * 补丁模式的段落属性：在原 pPr 的基础上替换本编辑器认识的子元素（与样式相同的去掉），
 * 不认识的子元素（边框、底纹、文字方向等）保留，再按规定顺序排好。
 */
function pPrDiffXml(p: Paragraph, b: ParaProps, bRun: RunProps, origPPr: string): string {
  const pr = p.props;
  const known: Record<string, string> = {};
  const styleId = pr.styleId ?? (pr.style ? patchCtx!.styleIdOf(pr.style) : undefined);
  if (styleId) known.pStyle = `<w:pStyle w:val="${esc(styleId)}"/>`;
  const flag = (name: string, v: boolean, base: boolean) => {
    if (v !== base) known[name] = v ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
  };
  flag('keepNext', pr.keepNext, b.keepNext);
  flag('keepLines', pr.keepLines, b.keepLines);
  flag('pageBreakBefore', pr.pageBreakBefore, b.pageBreakBefore);
  flag('widowControl', pr.widowControl, b.widowControl);
  flag('snapToGrid', pr.snapToGrid, b.snapToGrid);
  if (JSON.stringify(pr.num) !== JSON.stringify(b.num))
    known.numPr = pr.num ? `<w:numPr><w:ilvl w:val="${pr.num.level}"/><w:numId w:val="${pr.num.id}"/></w:numPr>` : '<w:numPr><w:numId w:val="0"/></w:numPr>';
  if (JSON.stringify(pr.tabs ?? []) !== JSON.stringify(b.tabs ?? [])) {
    const mine = pr.tabs ?? [];
    const clears = (b.tabs ?? []).filter((t) => !mine.some((x) => Math.abs(x.pos - t.pos) < 0.05)).map((t) => `<w:tab w:val="clear" w:pos="${twip(t.pos)}"/>`);
    known.tabs = `<w:tabs>${clears.join('')}${mine.map((t) => `<w:tab w:val="${t.align}" w:leader="${t.leader}" w:pos="${twip(t.pos)}"/>`).join('')}</w:tabs>`;
  }
  if (pr.spaceBefore !== b.spaceBefore || pr.spaceAfter !== b.spaceAfter || pr.spaceBeforeLines !== b.spaceBeforeLines || pr.spaceAfterLines !== b.spaceAfterLines || !!pr.spaceBeforeAuto !== !!b.spaceBeforeAuto || !!pr.spaceAfterAuto !== !!b.spaceAfterAuto || JSON.stringify(pr.line) !== JSON.stringify(b.line)) {
    let sp = `<w:spacing w:before="${twip(pr.spaceBefore)}"${autoAttr('before', pr.spaceBeforeAuto, b.spaceBeforeAuto)} w:after="${twip(pr.spaceAfter)}"${autoAttr('after', pr.spaceAfterAuto, b.spaceAfterAuto)}`;
    if (pr.spaceBeforeLines !== null) sp += ` w:beforeLines="${Math.round(pr.spaceBeforeLines * 100)}"`;
    if (pr.spaceAfterLines !== null) sp += ` w:afterLines="${Math.round(pr.spaceAfterLines * 100)}"`;
    const l = pr.line;
    sp += l.rule === 'auto' ? ` w:line="${Math.round(l.value * 240)}" w:lineRule="auto"` : ` w:line="${twip(l.value)}" w:lineRule="${l.rule}"`;
    known.spacing = sp + '/>';
  }
  if (pr.indentLeft !== b.indentLeft || pr.indentRight !== b.indentRight || pr.firstLine !== b.firstLine || pr.firstLineChars !== b.firstLineChars) {
    let ind = `<w:ind w:left="${twip(pr.indentLeft)}" w:right="${twip(pr.indentRight)}"`;
    if (pr.firstLineChars !== null) ind += pr.firstLineChars >= 0 ? ` w:firstLineChars="${Math.round(pr.firstLineChars * 100)}"` : ` w:hangingChars="${Math.round(-pr.firstLineChars * 100)}"`;
    else if (b.firstLineChars !== null) ind += ' w:firstLineChars="0"';
    ind += pr.firstLine >= 0 ? ` w:firstLine="${twip(pr.firstLine)}"` : ` w:hanging="${twip(-pr.firstLine)}"`;
    known.ind = ind + '/>';
  }
  if (pr.align !== b.align) known.jc = `<w:jc w:val="${{ left: 'left', center: 'center', right: 'right', both: 'both', distribute: 'distribute' }[pr.align]}"/>`;
  if (pr.outline !== b.outline) known.outlineLvl = `<w:outlineLvl w:val="${pr.outline ?? 9}"/>`;
  const mark = rPrDiffXml(p.mark, bRun);
  if (mark) known.rPr = mark;
  if (pr.sectionBreak) known.sectPr = patchCtx!.sectPr(p) ?? sectionXml(p);
  // 本编辑器管的子元素：以上面算出的为准（没有 = 与样式相同，去掉）；其它保留原样
  const MANAGED = new Set(['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'snapToGrid', 'numPr', 'tabs', 'spacing', 'ind', 'jc', 'outlineLvl', 'rPr', 'sectPr']);
  const parts: [string, string][] = pPrChildren(origPPr).filter(([n]) => !MANAGED.has(n));
  for (const [n, x] of Object.entries(known)) parts.push([n, x]);
  const rank = (n: string) => {
    const i = PPR_ORDER.indexOf(n);
    return i < 0 ? PPR_ORDER.indexOf('pPrChange') : i;
  };
  parts.sort((a, b2) => rank(a[0]) - rank(b2[0]));
  return parts.length ? `<w:pPr>${parts.map(([, x]) => x).join('')}</w:pPr>` : '';
}

/**
 * 补丁模式写一个段落：orig 是原文件里对应的段落（读入时的模型 + 原始 XML）。
 * 内容和格式都没变 → 原样；段落格式没变 → 原 pPr 照抄、只重写文字；否则按样式求差重写 pPr。
 */
export function paraPatchXml(p: Paragraph, orig?: { model: Paragraph; xml: string }): string {
  if (orig && JSON.stringify(p) === JSON.stringify(orig.model)) return orig.xml;
  const ctx = patchCtx!;
  const outer = { base: curBase, style: curStyle };
  const styleId = p.props.styleId ?? (p.props.style ? ctx.styleIdOf(p.props.style) : undefined);
  curStyle = styleId;
  curBase = ctx.base(styleId);
  const origPPr = orig?.xml.match(/^<w:p\b[^>]*>\s*(<w:pPr\b[\s\S]*?<\/w:pPr>|<w:pPr\s*\/>)/)?.[1] ?? '';
  const sameProps = orig && JSON.stringify(p.props) === JSON.stringify(orig.model.props) && JSON.stringify(p.mark) === JSON.stringify(orig.model.mark) && !p.props.sectionBreak;
  const pPr = sameProps ? origPPr : pPrDiffXml(p, curBase.para, curBase.run, origPPr);
  const open = orig?.xml.match(/^<w:p\b[^>]*?(?=\/?>)/)?.[0] ?? '<w:p';
  const out = `${open}>${pPr}${runsXml(p.runs)}</w:p>`;
  curBase = outer.base;
  curStyle = outer.style;
  return out;
}

/** 整篇重写时 doc.styleEdits 里的样式 id（段落 / 文字引用它们时写 pStyle / rStyle） */
let customStyleIds = new Set<string>();

function rPrXml(p: RunProps): string {
  return (
    '<w:rPr>' +
    (p.rStyle && customStyleIds.has(p.rStyle) ? `<w:rStyle w:val="${esc(p.rStyle)}"/>` : '') +
    `<w:rFonts w:ascii="${esc(p.fontLatin)}" w:hAnsi="${esc(p.fontLatin)}" w:eastAsia="${esc(p.fontEA)}" w:cs="${esc(p.fontLatin)}"/>` +
    (p.bold ? '<w:b/><w:bCs/>' : '') +
    (p.italic ? '<w:i/><w:iCs/>' : '') +
    (p.strike ? '<w:strike/>' : '') +
    (p.color !== 'auto' ? `<w:color w:val="${p.color}"/>` : '') +
    (p.spacing ? `<w:spacing w:val="${Math.round(p.spacing * 20)}"/>` : '') +
    `<w:sz w:val="${Math.round(p.size * 2)}"/><w:szCs w:val="${Math.round(p.size * 2)}"/>` +
    (p.highlight && HIGHLIGHT_NAME[p.highlight] ? `<w:highlight w:val="${HIGHLIGHT_NAME[p.highlight]}"/>` : '') +
    (p.underline ? '<w:u w:val="single"/>' : '') +
    (p.highlight && !HIGHLIGHT_NAME[p.highlight] ? `<w:shd w:val="clear" w:color="auto" w:fill="${p.highlight}"/>` : '') +
    (p.vertAlign ? `<w:vertAlign w:val="${p.vertAlign}"/>` : '') +
    '</w:rPr>'
  );
}

/** 当前正在写出的部件里图片的关系 id（writeDocx 设置）；返回 null 表示这个部件不写图片（页眉页脚） */
let imageRel: (id: string) => string | null = () => null;
/** 整篇重写时外部链接地址 → 关系 id（writeDocx 设置） */
let linkRel: (url: string) => string = () => '';
let drawingSeq = 0;

const pictureTransformAttrs = (i: RunImage) => `${i.rot ? ` rot="${Math.round(i.rot * 60000)}"` : ''}${i.flipH ? ' flipH="1"' : ''}${i.flipV ? ' flipV="1"' : ''}`;
const pictureCropXml = (i: RunImage) => i.crop ? `<a:srcRect${(['l', 't', 'r', 'b'] as const).map((k) => ` ${k}="${Math.round(i.crop![k] * 100000)}"`).join('')}/>` : '';

function drawingXml(image: RunImage, rid: string): string {
  const n = ++drawingSeq;
  const cx = Math.round(image.w * EMU_PER_PT);
  const cy = Math.round(image.h * EMU_PER_PT);
  return (
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${n}" name="图片 ${n}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}">` +
    `<pic:nvPicPr><pic:cNvPr id="${n}" name="图片 ${n}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/>${pictureCropXml(image)}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm${pictureTransformAttrs(image)}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
  );
}

/** 文本框 / 形状 → mc:AlternateContent 里的 wps:wsp（外形、翻转、旋转、填充、线条、文字按模型写；没有文字的形状不写 wps:txbx） */
function textBoxXml(b: TextBox): string {
  const n = ++drawingSeq;
  const emu = (pt: number) => Math.round(pt * EMU_PER_PT);
  const pos = (tag: string, p: TextBox['posH']) =>
    `<wp:${tag} relativeFrom="${esc(p.rel)}">${p.align ? `<wp:align>${esc(p.align)}</wp:align>` : `<wp:posOffset>${emu(p.offset)}</wp:posOffset>`}</wp:${tag}>`;
  const wrap =
    b.wrap === 'none' ? '<wp:wrapNone/>' : b.wrap === 'topAndBottom' ? '<wp:wrapTopAndBottom/>' : `<wp:wrap${b.wrap[0].toUpperCase()}${b.wrap.slice(1)} wrapText="bothSides"/>`;
  const fill = b.fill && !isLineGeom(b.geom) ? `<a:solidFill><a:srgbClr val="${b.fill}"/></a:solidFill>` : '<a:noFill/>';
  const ln = b.line ? `<a:ln w="${emu(b.line.w)}"><a:solidFill><a:srgbClr val="${b.line.color}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>';
  const va = b.vAlign === 'center' ? 'ctr' : b.vAlign === 'bottom' ? 'b' : 't';
  const xfrmAttrs = `${b.rot ? ` rot="${Math.round(b.rot * 60000)}"` : ''}${b.flipH ? ' flipH="1"' : ''}${b.flipV ? ' flipV="1"' : ''}`;
  const name = b.geom ? '形状' : '文本框';
  const txbx = b.paragraphs.length ? `<wps:txbx><w:txbxContent>${b.paragraphs.map(paraXml).join('')}</w:txbxContent></wps:txbx>` : '';
  return (
    `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor ${distAttrs(b.dist)} simplePos="0" relativeHeight="${251658240 + n}" behindDoc="${b.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>${pos('positionH', b.posH)}${pos('positionV', b.posV)}` +
    `<wp:extent cx="${emu(b.w)}" cy="${emu(b.h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}` +
    `<wp:docPr id="${n}" name="${name} ${n}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${WPS_NS}"><wps:wsp><wps:cNvSpPr${b.geom ? '' : ' txBox="1"'}/>` +
    `<wps:spPr><a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="${emu(b.w)}" cy="${emu(b.h)}"/></a:xfrm><a:prstGeom prst="${esc(b.geom ?? 'rect')}"><a:avLst/></a:prstGeom>${fill}${ln}</wps:spPr>` +
    txbx +
    `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${emu(b.inset.l)}" tIns="${emu(b.inset.t)}" rIns="${emu(b.inset.r)}" bIns="${emu(b.inset.b)}" anchor="${va}" anchorCtr="0" upright="1"><a:noAutofit/></wps:bodyPr>` +
    '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent>'
  );
}

/** 浮动图片 → wp:anchor + pic:pic（位置、大小、环绕、衬于文字下方按模型写） */
function floatImageXml(b: TextBox, rid: string): string {
  const n = ++drawingSeq;
  const emu = (pt: number) => Math.round(pt * EMU_PER_PT);
  const pos = (tag: string, p: TextBox['posH']) =>
    `<wp:${tag} relativeFrom="${esc(p.rel)}">${p.align ? `<wp:align>${esc(p.align)}</wp:align>` : `<wp:posOffset>${emu(p.offset)}</wp:posOffset>`}</wp:${tag}>`;
  const wrap =
    b.wrap === 'none' ? '<wp:wrapNone/>' : b.wrap === 'topAndBottom' ? '<wp:wrapTopAndBottom/>' : `<wp:wrap${b.wrap[0].toUpperCase()}${b.wrap.slice(1)} wrapText="bothSides"/>`;
  return (
    `<w:drawing><wp:anchor ${distAttrs(b.dist)} simplePos="0" relativeHeight="${251658240 + n}" behindDoc="${b.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>${pos('positionH', b.posH)}${pos('positionV', b.posV)}` +
    `<wp:extent cx="${emu(b.w)}" cy="${emu(b.h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}` +
    `<wp:docPr id="${n}" name="图片 ${n}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}">` +
    `<pic:nvPicPr><pic:cNvPr id="${n}" name="图片 ${n}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/>${pictureCropXml(b.image!)}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm${pictureTransformAttrs(b.image!)}><a:off x="0" y="0"/><a:ext cx="${emu(b.w)}" cy="${emu(b.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>'
  );
}

/** 批注：写出时每条批注（含回复）的范围标记跟随它所在的那条 */
let commentIds: (id: string) => string[] = (id) => [id];
let revSeq = 0;

function runXml(r: Run): string {
  if (r.mark) {
    const ids = commentIds(r.mark.id);
    if (r.mark.kind === 'cs') return ids.map((id) => `<w:commentRangeStart w:id="${esc(id)}"/>`).join('');
    return ids.map((id) => `<w:commentRangeEnd w:id="${esc(id)}"/><w:r><w:commentReference w:id="${esc(id)}"/></w:r>`).join('');
  }
  const body = runBodyXml(r);
  const rev = r.props.rev;
  if (!rev || !body) return body;
  const tag = rev.kind === 'ins' ? 'ins' : 'del';
  return `<w:${tag} w:id="${++revSeq}" w:author="${esc(rev.author)}"${rev.date ? ` w:date="${esc(rev.date)}"` : ''}>${body}</w:${tag}>`;
}

function runBodyXml(r: Run): string {
  const rpr = runPr(r.props);
  const T = r.props.rev?.kind === 'del' ? 'delText' : 't';
  if (r.opaque) {
    // 整篇重写时，引用了其它部件的对象写不出来（部件不在新文件里）
    if (!patchCtx && REL_REF.test(r.opaque.xml)) return '';
    return r.opaque.level === 'para' ? r.opaque.xml : `<w:r>${rpr}${r.opaque.xml}</w:r>`;
  }
  if (r.box && patchCtx && r.box.raw && patchCtx.boxUnchanged(r.box)) return `<w:r>${rpr}${r.box.raw}</w:r>`;
  if (r.box?.image) {
    if (patchCtx && r.box.raw) {
      const image = r.box.image;
      const rid = image.rawId && image.id !== image.rawId ? patchCtx.newImageRel(image.id) : null;
      return `<w:r>${rpr}${patchPictureDrawing(r.box.raw, image, rid, r.box)}</w:r>`;
    }
    const rid = patchCtx ? patchCtx.newImageRel(r.box.image.id) : imageRel(r.box.image.id);
    return rid ? `<w:r>${rpr}${floatImageXml(r.box, rid)}</w:r>` : '';
  }
  if (r.box) return `<w:r>${rpr}${textBoxXml(r.box)}</w:r>`;
  if (r.image) {
    if (patchCtx && r.image.raw) {
      const rid = r.image.rawId && r.image.id !== r.image.rawId ? patchCtx.newImageRel(r.image.id) : null;
      return `<w:r>${rpr}${patchPictureDrawing(r.image.raw, r.image, rid)}</w:r>`;
    }
    const rid = patchCtx ? patchCtx.newImageRel(r.image.id) : imageRel(r.image.id);
    return rid ? `<w:r>${rpr}${drawingXml(r.image, rid)}</w:r>` : '';
  }
  if (r.fc) return fieldMarkXml(r.fc, rpr);
  if (r.field) return `<w:fldSimple w:instr=" ${r.field} \\* MERGEFORMAT "><w:r>${rpr}<w:t>1</w:t></w:r></w:fldSimple>`;
  let out = '';
  let buf = '';
  const flush = () => {
    if (buf) out += `<w:${T} xml:space="preserve">${esc(buf)}</w:${T}>`;
    buf = '';
  };
  for (const ch of r.text) {
    if (ch === '\t') {
      flush();
      out += '<w:tab/>';
    } else if (ch === '\n') {
      flush();
      out += '<w:br/>';
    } else if (ch === '\f') {
      flush();
      out += '<w:br w:type="page"/>';
    } else if (ch === '\v') {
      flush();
      out += '<w:br w:type="column"/>';
    } else if (ch >= '\uf000' && ch <= '\uf0ff') {
      // 符号字体的字符写回 w:sym
      flush();
      out += `<w:sym w:font="${esc(r.props.fontLatin)}" w:char="${ch.charCodeAt(0).toString(16).toUpperCase()}"/>`;
    } else buf += ch;
  }
  flush();
  return `<w:r>${rpr}${out}</w:r>`;
}

/** 当前正在写出的文档中分节段落的节属性（writeDocx 开始时设置） */
let sectionXml: (p: Paragraph) => string = () => '';

function pPrXml(p: Paragraph): string {
  const pr = p.props;
  const jc = { left: 'left', center: 'center', right: 'right', both: 'both', distribute: 'distribute' }[pr.align];
  let ind = `<w:ind w:left="${twip(pr.indentLeft)}" w:right="${twip(pr.indentRight)}"`;
  if (pr.firstLineChars !== null) ind += pr.firstLineChars >= 0 ? ` w:firstLineChars="${Math.round(pr.firstLineChars * 100)}"` : ` w:hangingChars="${Math.round(-pr.firstLineChars * 100)}"`;
  if (pr.firstLine >= 0) ind += ` w:firstLine="${twip(pr.firstLine)}"`;
  else ind += ` w:hanging="${twip(-pr.firstLine)}"`;
  ind += '/>';
  let sp = `<w:spacing w:before="${twip(pr.spaceBefore)}"${pr.spaceBeforeAuto ? ' w:beforeAutospacing="1"' : ''} w:after="${twip(pr.spaceAfter)}"${pr.spaceAfterAuto ? ' w:afterAutospacing="1"' : ''}`;
  if (pr.spaceBeforeLines !== null) sp += ` w:beforeLines="${Math.round(pr.spaceBeforeLines * 100)}"`;
  if (pr.spaceAfterLines !== null) sp += ` w:afterLines="${Math.round(pr.spaceAfterLines * 100)}"`;
  const l = pr.line;
  sp += l.rule === 'auto' ? ` w:line="${Math.round(l.value * 240)}" w:lineRule="auto"` : ` w:line="${twip(l.value)}" w:lineRule="${l.rule}"`;
  sp += '/>';
  return (
    '<w:pPr>' +
    (pr.style ? `<w:pStyle w:val="${pr.style}"/>` : pr.styleId && customStyleIds.has(pr.styleId) ? `<w:pStyle w:val="${esc(pr.styleId)}"/>` : '') +
    (pr.keepNext ? '<w:keepNext/>' : '') +
    (pr.keepLines ? '<w:keepLines/>' : '') +
    (pr.pageBreakBefore ? '<w:pageBreakBefore/>' : '') +
    `<w:widowControl w:val="${pr.widowControl ? 1 : 0}"/>` +
    (pr.num ? `<w:numPr><w:ilvl w:val="${pr.num.level}"/><w:numId w:val="${pr.num.id}"/></w:numPr>` : '') +
    (pr.frame
      ? `<w:framePr w:wrap="around" w:vAnchor="${pr.frame.vAnchor}" w:hAnchor="${pr.frame.hAnchor}"${pr.frame.xAlign ? ` w:xAlign="${pr.frame.xAlign}"` : ` w:x="${twip(pr.frame.x)}"`} w:y="${twip(pr.frame.y)}"/>`
      : '') +
    (pr.borders
      ? `<w:pBdr>${(['top', 'left', 'bottom', 'right'] as const)
          .filter((k) => pr.borders![k])
          .map((k) => `<w:${k} w:val="single" w:sz="${Math.round(pr.borders![k]!.w * 8)}" w:space="${Math.round(pr.borders![k]!.space)}" w:color="${pr.borders![k]!.color}"/>`)
          .join('')}</w:pBdr>`
      : '') +
    (pr.tabs?.length
      ? `<w:tabs>${pr.tabs.map((t) => `<w:tab w:val="${t.align}" w:leader="${t.leader}" w:pos="${twip(t.pos)}"/>`).join('')}</w:tabs>`
      : '') +
    (pr.snapToGrid ? '' : '<w:snapToGrid w:val="0"/>') +
    sp +
    ind +
    `<w:jc w:val="${jc}"/>` +
    (pr.outline !== undefined ? `<w:outlineLvl w:val="${pr.outline}"/>` : '') +
    rPrXml(p.mark) +
    (pr.sectionBreak ? sectionXml(p) : '') +
    '</w:pPr>'
  );
}

const FMT_OUT: Record<NumFmt, string> = {
  decimal: 'decimal',
  chineseCounting: 'chineseCountingThousand',
  ideographTraditional: 'ideographTraditional',
  upperLetter: 'upperLetter',
  lowerLetter: 'lowerLetter',
  upperRoman: 'upperRoman',
  lowerRoman: 'lowerRoman',
  decimalEnclosedCircle: 'decimalEnclosedCircle',
  bullet: 'bullet',
  none: 'none',
};

/** 编号自己的字符格式（只写设置了的项） */
function labelRPr(p: Partial<RunProps> | undefined): string {
  if (!p) return '';
  const fonts = (p.fontLatin ? ` w:ascii="${esc(p.fontLatin)}" w:hAnsi="${esc(p.fontLatin)}"` : '') + (p.fontEA ? ` w:eastAsia="${esc(p.fontEA)}"` : '');
  return (
    '<w:rPr>' +
    (fonts ? `<w:rFonts${fonts}/>` : '') +
    (p.bold === undefined ? '' : p.bold ? '<w:b/>' : '<w:b w:val="0"/>') +
    (p.italic === undefined ? '' : p.italic ? '<w:i/>' : '<w:i w:val="0"/>') +
    (p.color && p.color !== 'auto' ? `<w:color w:val="${p.color}"/>` : '') +
    (p.size ? `<w:sz w:val="${Math.round(p.size * 2)}"/>` : '') +
    '</w:rPr>'
  );
}

function numberingXml(lists: NumList[]): string {
  const lvl = (l: NumLevel, i: number) => {
    let ind = `<w:ind w:left="${twip(l.indentLeft)}"`;
    if (l.firstLineChars !== null) ind += ` w:firstLineChars="${Math.round(l.firstLineChars * 100)}"`;
    ind += l.firstLine < 0 ? ` w:hanging="${twip(-l.firstLine)}"/>` : ` w:firstLine="${twip(l.firstLine)}"/>`;
    return (
      `<w:lvl w:ilvl="${i}"><w:start w:val="${l.start}"/><w:numFmt w:val="${FMT_OUT[l.fmt]}"/>` +
      (l.suffix !== 'tab' ? `<w:suff w:val="${l.suffix}"/>` : '') +
      `<w:lvlText w:val="${esc(l.text)}"/><w:lvlJc w:val="left"/><w:pPr>${ind}</w:pPr>${labelRPr(l.label)}</w:lvl>`
    );
  };
  return (
    `${XML_HEAD}<w:numbering ${NS}>` +
    lists.map((l, i) => `<w:abstractNum w:abstractNumId="${i}"><w:multiLevelType w:val="multilevel"/>${l.levels.slice(0, 9).map(lvl).join('')}</w:abstractNum>`).join('') +
    lists.map((l, i) => `<w:num w:numId="${l.id}"><w:abstractNumId w:val="${i}"/></w:num>`).join('') +
    '</w:numbering>'
  );
}

/** 域的开始部分（begin、域代码、separate）或结束部分（end）：读入的原样写回，改过代码的按代码生成（嵌套的 {…} 写成嵌套域） */
function fieldMarkXml(fc: NonNullable<Run['fc']>, rpr: string): string {
  if (fc.raw) return fc.raw;
  const char = (type: string, extra = '') => `<w:r>${rpr}<w:fldChar w:fldCharType="${type}"${extra}/></w:r>`;
  if (fc.kind === 'end') return char('end');
  // 嵌套域：begin 代码 separate 结果 end（结果是上次更新时算的，Word 算外层域时要用）
  let k = 0;
  const open: number[] = [];
  const code = instrPieces(` ${fc.instr ?? ''} `)
    .map((p) => {
      if (p.t === 'text') return `<w:r>${rpr}<w:instrText xml:space="preserve">${esc(p.v)}</w:instrText></w:r>`;
      if (p.t === 'begin') {
        open.push(k++);
        return char('begin');
      }
      const v = fc.inner?.[open.pop() ?? -1];
      return char('separate') + (v ? `<w:r>${rpr}<w:t xml:space="preserve">${esc(v)}</w:t></w:r>` : '') + char('end');
    })
    .join('');
  return char('begin', fc.locked ? ' w:fldLock="1"' : '') + code + char('separate');
}

/** 文本段 → XML：连续的、链接相同的文本段包在一个 w:hyperlink 里；配不上对的域标记不写（只剩结果文字） */
function runsXml(all: Run[]): string {
  const paired = all.some((r) => r.fc) ? pairedMarkers(all) : null;
  const runs = paired ? all.filter((r) => !r.fc || paired.has(r)) : all;
  let out = '';
  for (let i = 0; i < runs.length; ) {
    const link = runs[i].props.link;
    if (!link || runs[i].mark || runs[i].opaque?.level === 'para') {
      out += runXml(runs[i]);
      i++;
      continue;
    }
    let j = i;
    let inner = '';
    const same = (l: typeof link | undefined) => l?.rid === link.rid && l?.anchor === link.anchor && l?.url === link.url;
    while (j < runs.length && same(runs[j].props.link) && !runs[j].mark && runs[j].opaque?.level !== 'para') inner += runXml(runs[j++]);
    // 读入的链接保留原来的关系 id；新加的外部链接按地址登记关系
    const rid = link.rid ?? (link.url ? (patchCtx ? patchCtx.newLinkRel(link.url) : linkRel(link.url)) : undefined);
    out += `<w:hyperlink${rid ? ` r:id="${esc(rid)}"` : ''}${link.anchor ? ` w:anchor="${esc(link.anchor)}"` : ''}>${inner}</w:hyperlink>`;
    i = j;
  }
  return out;
}

/** 原图片的 w:drawing 改成新的显示大小（只改 wp:extent 和 a:ext，其它照旧） */
function resizeDrawing(raw: string, w: number, h: number): string {
  const cx = Math.round(w * EMU_PER_PT);
  const cy = Math.round(h * EMU_PER_PT);
  const old = raw.match(/<wp:extent\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
  if (old && Math.abs(Number(old[1]) - cx) <= 1 && Math.abs(Number(old[2]) - cy) <= 1) return raw;
  return raw.replace(/(<wp:extent\b[^>]*?)cx="\d+"([^>]*?)cy="\d+"/, `$1cx="${cx}"$2cy="${cy}"`).replace(/(<a:ext\b[^>]*?)cx="\d+"([^>]*?)cy="\d+"/, `$1cx="${cx}"$2cy="${cy}"`);
}

/** Patch picture-owned fields; keep document ids, alt text, effects and extension data. */
function patchPictureDrawing(raw: string, image: RunImage, rid: string | null, box?: TextBox): string {
  const set = (tag: string, key: string, value: string | undefined) => {
    const re = new RegExp(`\\s${key}="[^"]*"`);
    if (value === undefined) return tag.replace(re, '');
    return re.test(tag) ? tag.replace(re, ` ${key}="${value}"`) : tag.replace(/\/?>(?=$)/, (end) => ` ${key}="${value}"${end}`);
  };
  let xml = resizeDrawing(raw, image.w, image.h);
  xml = xml.replace(/<a:xfrm\b[^>]*>/, (tag) => {
    tag = set(tag, 'rot', image.rot ? String(Math.round(image.rot * 60000)) : undefined);
    tag = set(tag, 'flipH', image.flipH ? '1' : undefined);
    return set(tag, 'flipV', image.flipV ? '1' : undefined);
  });
  xml = xml.replace(/<pic:blipFill\b[^>]*>[\s\S]*?<\/pic:blipFill>/, (fill) => {
    let out = fill.replace(/<a:srcRect\b[^>]*\/>/, '');
    if (image.crop) out = out.replace(/<a:(stretch|tile)\b/, `${pictureCropXml(image)}<a:$1`);
    if (rid) out = out.replace(/<a:blip\b[^>]*>/, (tag) => set(tag, 'r:embed', esc(rid)));
    return out;
  });
  if (box) {
    const emu = (n: number) => String(Math.round(n * EMU_PER_PT));
    xml = xml.replace(/<wp:anchor\b[^>]*>/, (tag) => {
      tag = set(tag, 'behindDoc', box.behind ? '1' : '0');
      for (const [key, n] of Object.entries(box.dist ?? { t: 0, b: 0, l: 0, r: 0 })) tag = set(tag, `dist${key.toUpperCase()}`, emu(n));
      return tag;
    });
    for (const [tag, pos] of [['positionH', box.posH], ['positionV', box.posV]] as const) {
      xml = xml.replace(new RegExp(`<wp:${tag}\\b[^>]*>[\\s\\S]*?<\\/wp:${tag}>`), `<wp:${tag} relativeFrom="${esc(pos.rel)}">${pos.align ? `<wp:align>${esc(pos.align)}</wp:align>` : `<wp:posOffset>${emu(pos.offset)}</wp:posOffset>`}</wp:${tag}>`);
    }
    const wrap = box.wrap === 'none' ? '<wp:wrapNone/>' : box.wrap === 'topAndBottom' ? '<wp:wrapTopAndBottom/>' : `<wp:wrap${box.wrap[0].toUpperCase()}${box.wrap.slice(1)} wrapText="bothSides"/>`;
    // Preserve tight/through polygons when the mode did not change.
    if (!xml.includes(`<wp:wrap${box.wrap[0].toUpperCase()}${box.wrap.slice(1)}`))
      xml = xml.replace(/<wp:wrap(?:None|Square|Tight|Through|TopAndBottom)\b[^>]*(?:\/>|>[\s\S]*?<\/wp:wrap(?:None|Square|Tight|Through|TopAndBottom)>)/, wrap);
  }
  return xml;
}

function paraXml(p: Paragraph): string {
  if (p.hfTable) return patchCtx ? p.hfTable.xml : tableXml(p.hfTable.table);
  if (patchCtx) return paraPatchXml(p);
  return `<w:p>${pPrXml(p)}${runsXml(p.runs)}</w:p>`;
}

export function tableXml(t: Table): string {
  const b = t.borders ? 'single' : 'none';
  const sz = Math.max(2, Math.round((t.borderSize ?? 0.5) * 8));
  const border = (n: string) => `<w:${n} w:val="${b}" w:sz="${sz}" w:space="0" w:color="auto"/>`;
  let x =
    '<w:tbl><w:tblPr>' +
    tableFloatXml(t.float) +
    (t.widthPct ? `<w:tblW w:w="${Math.round(t.widthPct * 5000)}" w:type="pct"/>` : `<w:tblW w:w="${twip(t.cols.reduce((a, c) => a + c, 0))}" w:type="dxa"/>`) +
    `<w:jc w:val="${t.align}"/>` +
    (t.align === 'left' && t.indent ? `<w:tblInd w:w="${twip(t.indent)}" w:type="dxa"/>` : '') +
    `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>` +
    '<w:tblLayout w:type="fixed"/>' +
    `<w:tblCellMar><w:top w:w="${twip(t.cellMargin.top)}" w:type="dxa"/><w:left w:w="${twip(t.cellMargin.left)}" w:type="dxa"/><w:bottom w:w="${twip(t.cellMargin.bottom)}" w:type="dxa"/><w:right w:w="${twip(t.cellMargin.right)}" w:type="dxa"/></w:tblCellMar>` +
    '</w:tblPr><w:tblGrid>' +
    t.cols.map((c) => `<w:gridCol w:w="${twip(c)}"/>`).join('') +
    '</w:tblGrid>';
  for (const row of t.rows) {
    x += '<w:tr><w:trPr>';
    if (row.heightRule !== 'auto' && row.height > 0) x += `<w:trHeight w:val="${twip(row.height)}" w:hRule="${row.heightRule}"/>`;
    if (row.header) x += '<w:tblHeader/>';
    x += '</w:trPr>';
    let g = 0;
    for (const c of row.cells) {
      const w = t.cols.slice(g, g + c.gridSpan).reduce((a, v) => a + v, 0);
      g += c.gridSpan;
      x +=
        '<w:tc><w:tcPr>' +
        `<w:tcW w:w="${twip(w)}" w:type="dxa"/>` +
        (c.gridSpan > 1 ? `<w:gridSpan w:val="${c.gridSpan}"/>` : '') +
        (c.vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : c.vMerge === 'continue' ? '<w:vMerge/>' : '') +
        (c.bdr
          ? tcBordersXml(c.bdr)
          : c.borders !== undefined
          ? `<w:tcBorders>${['top', 'left', 'bottom', 'right'].map((n) => `<w:${n} w:val="${c.borders ? 'single' : 'nil'}" w:sz="${sz}" w:space="0" w:color="auto"/>`).join('')}</w:tcBorders>`
          : '') +
        (c.shading ? `<w:shd w:val="clear" w:color="auto" w:fill="${c.shading}"/>` : '') +
        `<w:vAlign w:val="${c.vAlign}"/>` +
        '</w:tcPr>' +
        (c.paragraphs.length ? c.paragraphs : [{ kind: 'p', props: t.rows[0].cells[0].paragraphs[0].props, runs: [], mark: t.rows[0].cells[0].paragraphs[0].mark } as Paragraph])
          .map(paraXml)
          .join('') +
        '</w:tc>';
    }
    x += '</w:tr>';
  }
  return x + '</w:tbl>';
}

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
// 整篇重写时根元素声明的命名空间：除了本编辑器自己写的，还有原样对象（公式、形状、VML、OLE…）常用的前缀
const NS =
  `xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:wps="${WPS_NS}" xmlns:mc="${MC_NS}" ` +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:w14ink="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"';

/** 批注部件：comments.xml（纯文字正文）和记回复关系的 commentsExtended.xml（没有回复时为空串） */
export function commentParts(written: Comment[]): { comments: string; ext: string } {
  // 每条批注最后一段的 paraId：commentsExtended.xml 用它记回复关系（Word 2013 起的批注会话）
  const paraId = (c: Comment) => (0x10000000 + Number(c.id || 0)).toString(16).toUpperCase();
  const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
  // NS 里已经声明了 w14（段落的 paraId 用）
  const comments = written.length
    ? `${XML_HEAD}<w:comments ${NS}>${written
        .map((c) => {
          const lines = c.text.split('\n');
          return (
            `<w:comment w:id="${esc(c.id)}" w:author="${esc(c.author)}"${c.date ? ` w:date="${esc(c.date)}"` : ''} w:initials="${esc(c.initials ?? c.author.slice(0, 2))}">` +
            lines.map((t, i) => `<w:p${i === lines.length - 1 ? ` w14:paraId="${paraId(c)}"` : ''}><w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`).join('') +
            '</w:comment>'
          );
        })
        .join('')}</w:comments>`
    : '';
  const ext = written.some((c) => c.parent)
    ? `${XML_HEAD}<w15:commentsEx xmlns:w15="${W15}">${written
        .map((c) => {
          const parent = c.parent ? written.find((x) => x.id === c.parent) : undefined;
          return `<w15:commentEx w15:paraId="${paraId(c)}"${parent ? ` w15:paraIdParent="${paraId(parent)}"` : ''} w15:done="0"/>`;
        })
        .join('')}</w15:commentsEx>`
    : '';
  return { comments, ext };
}

/** 一个编号列表 → abstractNum + num（abstractNumId 由调用方给） */
export function listXml(l: NumList, abstractId: number): { abstract: string; num: string } {
  const full = numberingXml([l]);
  const abstract = full.match(/<w:abstractNum\b[\s\S]*<\/w:abstractNum>/)![0].replace(/w:abstractNumId="\d+"/, `w:abstractNumId="${abstractId}"`);
  return { abstract, num: `<w:num w:numId="${l.id}"><w:abstractNumId w:val="${abstractId}"/></w:num>` };
}

/** settings.xml 的 w:documentProtection（与 docx-engine 的写法一致） */
function protectionXml(p: NonNullable<Doc['protection']>): string {
  const crypt = p.hash
    ? ` w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="${p.algorithmSid ?? 14}" w:cryptSpinCount="${p.spinCount ?? 100000}" w:hash="${esc(p.hash)}"${p.salt ? ` w:salt="${esc(p.salt)}"` : ''}`
    : '';
  return `<w:documentProtection w:edit="${esc(p.edit)}"${p.enforced ? ' w:enforcement="1"' : ''}${crypt}/>`;
}

/** 水印段落的 XML（放在页眉部件里第一个元素）：文字水印直接生成；图片水印要先在这个部件的关系表里登记图片（imageRel 返回关系 id） */
export function watermarkXml(w: Watermark, page: PageGeom, imageRel: (id: string) => string | null): string {
  if (!isPictureWatermark(w)) return watermarkParagraphXml(w);
  const rid = imageRel(w.image.id);
  if (!rid) return '';
  const spec = { image: { base64: '', mime: 'image/png' as const, widthPx: w.image.widthPx, heightPx: w.image.heightPx }, ...(w.scale !== undefined ? { scale: w.scale } : {}), ...(w.washout === false ? { washout: false } : {}) };
  return pictureWatermarkParagraphXml(spec, rid, { widthPt: page.width - page.marginLeft - page.marginRight, heightPt: page.height - page.marginTop - page.marginBottom });
}

export function writeDocx(doc: Doc): Uint8Array {
  customStyleIds = new Set(Object.keys(doc.styleEdits ?? {}));
  const s = doc.setup;
  const grid = s.grid.type === 'none' ? `<w:docGrid w:linePitch="${twip(s.grid.linePitch)}"/>` : `<w:docGrid w:type="${s.grid.type}" w:linePitch="${twip(s.grid.linePitch)}" w:charSpace="${Math.round(s.grid.charSpace)}"/>`;
  // 页眉页脚：每个不同的段落数组写成一个文件；第一节写出全部引用，之后各节只写与上一节不同的分区，
  // 省略的引用在 Word 中即“与上一节相同”
  const sections = sectionList(doc);
  const parts: { tag: 'hdr' | 'ftr'; file: string; rid: string; ps: Paragraph[]; story: HFStory }[] = [];
  const partOf = new Map<Paragraph[], (typeof parts)[number]>();
  const part = (story: HFStory, ps: Paragraph[]) => {
    let x = partOf.get(ps);
    if (!x) {
      const hdr = isHeaderStory(story);
      const n = parts.filter((q) => q.tag === (hdr ? 'hdr' : 'ftr')).length + 1;
      x = { tag: hdr ? 'hdr' : 'ftr', file: `${hdr ? 'header' : 'footer'}${n}.xml`, rid: `rId${hdr ? 'Header' : 'Footer'}${n}`, ps, story };
      parts.push(x);
      partOf.set(ps, x);
    }
    return x;
  };
  const headersFirst = [...HF_STORIES].sort((a, b) => Number(!isHeaderStory(a)) - Number(!isHeaderStory(b)));
  // 原文件里没有的页眉 / 页脚（占位段落还没写内容）不写引用：写成空页眉页脚的话 Word 会给它留位置，正文就挪了
  const absent = (ps: Paragraph[]) => ps.length === 1 && !!ps[0].hfNone && !ps[0].runs.length;
  const refs = sections.map(({ hf }, i) =>
    headersFirst
      .filter((st) => i === 0 || hf[st] !== sections[i - 1].hf[st])
      .filter((st) => !(absent(hf[st]) && (i === 0 || absent(sections[i - 1].hf[st]))))
      .map((st) => `<w:${isHeaderStory(st) ? 'header' : 'footer'}Reference w:type="${hfType(st)}" r:id="${part(st, hf[st]).rid}"/>`)
      .join(''),
  );
  // 分节段落 → 节序号（与 sectionList 顺序一致）
  const sectionIndex = new Map<Paragraph, number>();
  // 各节 sectPr 的 w:type = 这一节前面那个分节符的类型（第一节没有）
  const startType: (SectionBreak | null)[] = [null];
  for (const b of doc.body) {
    if (b.kind !== 'p' || !b.props.sectionBreak) continue;
    sectionIndex.set(b, sectionIndex.size);
    startType.push(b.props.sectionBreak);
  }
  const sectPr = (i: number) => {
    const g = sections[i].setup;
    const type = startType[i] ?? null;
    return (
      '<w:sectPr>' +
      refs[i] +
      (type && type !== 'nextPage' ? `<w:type w:val="${type}"/>` : '') +
      `<w:pgSz w:w="${twip(g.width)}" w:h="${twip(g.height)}"${g.width > g.height ? ' w:orient="landscape"' : ''}/>` +
      `<w:pgMar w:top="${twip(g.marginTop)}" w:right="${twip(g.marginRight)}" w:bottom="${twip(g.marginBottom)}" w:left="${twip(g.marginLeft)}" w:header="${twip(g.headerDistance)}" w:footer="${twip(g.footerDistance)}" w:gutter="0"/>` +
      colsXml(sections[i].setup.cols) +
      (sections[i].hf.titlePg ? '<w:titlePg/>' : '') +
      grid +
      '</w:sectPr>'
    );
  };
  // 表格内的段落不分节，按最后一节写出
  sectionXml = (p) => sectPr(sectionIndex.get(p) ?? sections.length - 1);
  // 图片：每张不同的图片写成一个 word/media 文件；各部件（正文、每个页眉页脚）有自己的关系表（图片、外部链接）
  const media = new Map<string, { file: string; ext: string; data: Uint8Array }>();
  interface PartRels {
    images: Map<string, string>;
    links: Map<string, string>;
  }
  const newRels = (): PartRels => ({ images: new Map(), links: new Map() });
  const useRels = (rels: PartRels) => {
    imageRel = (id) => {
      let rid = rels.images.get(id);
      if (!rid) {
        let m = media.get(id);
        if (!m) {
          const img = imageData(id);
          if (!img) return null;
          const ext = imageExt(img.mime);
          m = { file: `media/image${media.size + 1}.${ext}`, ext, data: img.data };
          media.set(id, m);
        }
        rels.images.set(id, (rid = `rIdImage${rels.images.size + 1}`));
      }
      return rid;
    };
    linkRel = (url) => {
      let rid = rels.links.get(url);
      if (!rid) rels.links.set(url, (rid = `rIdLink${rels.links.size + 1}`));
      return rid;
    };
  };
  const relsXml = (rels: PartRels) =>
    [...rels.images].map(([id, rid]) => `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${media.get(id)!.file}"/>`).join('') +
    [...rels.links].map(([url, rid]) => `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(url)}" TargetMode="External"/>`).join('');
  drawingSeq = 0;
  const docPartRels = newRels();
  useRels(docPartRels);
  // 批注：回复跟着它回复的那条写范围标记；没有范围标记的（范围被删掉了）不写
  const comments = doc.comments ?? [];
  const anchored = new Set<string>();
  for (const b of doc.body) {
    const ps = b.kind === 'p' ? [b] : b.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs));
    for (const p of ps) for (const r of p.runs) if (r.mark?.kind === 'cs') anchored.add(r.mark.id);
  }
  const rootOf = (c: Comment): string => (c.parent && comments.some((x) => x.id === c.parent) ? rootOf(comments.find((x) => x.id === c.parent)!) : c.id);
  const written = comments.filter((c) => anchored.has(rootOf(c)));
  commentIds = (id) => [id, ...written.filter((c) => c.id !== id && rootOf(c) === id).map((c) => c.id)];
  revSeq = 0;
  const body = doc.body.map((b) => (b.kind === 'p' ? paraXml(b) : tableXml(b))).join('');
  commentIds = (id) => [id];
  const { comments: commentsXml, ext: commentsExXml } = commentParts(written);
  const document = `${XML_HEAD}<w:document ${NS}><w:body>${body}${sectPr(sections.length - 1)}</w:body></w:document>`;
  // 页眉页脚：每个部件按自己的关系表写（图片、链接）；文字水印放在默认页眉的第一个元素
  const partRels = new Map<string, PartRels>();
  const hfXml = new Map<string, string>();
  for (const p of parts) {
    const rels = newRels();
    useRels(rels);
    const wm = p.tag === 'hdr' && doc.watermark ? watermarkXml(doc.watermark, sections[0].setup, imageRel) : '';
    hfXml.set(p.file, `${XML_HEAD}<w:${p.tag} ${NS}>${wm}${p.ps.map(paraXml).join('')}</w:${p.tag}>`);
    partRels.set(p.file, rels);
  }
  useRels(docPartRels);
  const d = doc.defaults;
  const usedStyles = new Set<string>();
  for (const b of doc.body) {
    const ps = b.kind === 'p' ? [b] : b.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs));
    for (const p of ps) if (p.props.style) usedStyles.add(p.props.style);
  }
  let styles =
    `${XML_HEAD}<w:styles ${NS}><w:docDefaults><w:rPrDefault>${rPrXml(d)}</w:rPrDefault>` +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    // 用到的标题 / 目录样式（格式都直接写在段落上，样式只负责名称和大纲级别，Word 的导航窗格和目录靠它们）
    [...usedStyles]
      .sort()
      .map((id) => {
        const n = Number(id.slice(-1));
        return id.startsWith('Heading')
          ? `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="${n - 1}"/></w:pPr></w:style>`
          : `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="toc ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/></w:style>`;
      })
      .join('') +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '</w:styles>';
  // 新建 / 修改的样式（样式管理）：整篇重写时都是新建
  for (const e of Object.values(doc.styleEdits ?? {})) styles = upsertStyleXml(styles, editToUpsert(e));
  // doNotCompress：不压缩标点，与本编辑器的排版一致
  const settings =
    `${XML_HEAD}<w:settings ${NS}><w:zoom w:percent="100"/>${doc.trackRevisions ? '<w:trackRevisions/>' : ''}${doc.protection ? protectionXml(doc.protection) : ''}<w:defaultTabStop w:val="${twip(s.defaultTab)}"/>` +
    (s.evenAndOdd ? '<w:evenAndOddHeaders/>' : '') +
    (s.noPunctKern ? '<w:noPunctuationKerning/>' : '') +
    `<w:characterSpacingControl w:val="${s.compressPunct ? 'compressPunctuation' : 'doNotCompress'}"/>` +
    `<w:compat>${s.noExpandShiftReturn ? '<w:doNotExpandShiftReturn/>' : ''}${s.tableGrid ? '<w:adjustLineHeightInTable/>' : ''}${s.altKinsoku ? '<w:useAltKinsokuLineBreakRules/>' : ''}` +
    `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${s.compatMode ?? 15}"/></w:compat>` +
    '</w:settings>';
  const hasNumbering = doc.lists.length > 0;
  const W = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ct =
    `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    [...new Set([...media.values()].map((m) => m.ext))].map((ext) => `<Default Extension="${ext}" ContentType="image/${ext}"/>`).join('') +
    `<Override PartName="/word/document.xml" ContentType="${W}.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="${W}.styles+xml"/>` +
    `<Override PartName="/word/settings.xml" ContentType="${W}.settings+xml"/>` +
    (hasNumbering ? `<Override PartName="/word/numbering.xml" ContentType="${W}.numbering+xml"/>` : '') +
    (commentsXml ? `<Override PartName="/word/comments.xml" ContentType="${W}.comments+xml"/>` : '') +
    (commentsExXml ? `<Override PartName="/word/commentsExtended.xml" ContentType="${W}.commentsExtended+xml"/>` : '') +
    parts.map((p) => `<Override PartName="/word/${p.file}" ContentType="${W}.${p.tag === 'hdr' ? 'header' : 'footer'}+xml"/>`).join('') +
    '</Types>';
  const rootRels =
    `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>` +
    '</Relationships>';
  const docRels =
    `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>` +
    `<Relationship Id="rIdSettings" Type="${R}/settings" Target="settings.xml"/>` +
    (hasNumbering ? `<Relationship Id="rIdNumbering" Type="${R}/numbering" Target="numbering.xml"/>` : '') +
    (commentsXml ? `<Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/>` : '') +
    (commentsExXml ? '<Relationship Id="rIdCommentsEx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>' : '') +
    parts.map((p) => `<Relationship Id="${p.rid}" Type="${R}/${p.tag === 'hdr' ? 'header' : 'footer'}" Target="${p.file}"/>`).join('') +
    relsXml(docPartRels) +
    '</Relationships>';
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(ct),
    '_rels/.rels': strToU8(rootRels),
    'word/document.xml': strToU8(document),
    'word/styles.xml': strToU8(styles),
    'word/settings.xml': strToU8(settings),
    'word/_rels/document.xml.rels': strToU8(docRels),
  };
  for (const p of parts) {
    files[`word/${p.file}`] = strToU8(hfXml.get(p.file)!);
    const rels = relsXml(partRels.get(p.file)!);
    if (rels) files[`word/_rels/${p.file}.rels`] = strToU8(`${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  }
  for (const m of media.values()) files[`word/${m.file}`] = m.data;
  if (hasNumbering) files['word/numbering.xml'] = strToU8(numberingXml(doc.lists));
  if (commentsXml) files['word/comments.xml'] = strToU8(commentsXml);
  if (commentsExXml) files['word/commentsExtended.xml'] = strToU8(commentsExXml);
  return zipSync(files);
}

/** wp:anchor 的环绕距离属性（磅 → EMU）；没有时用 Word 新建浮动对象的默认值（左右 0.32 厘米） */
function distAttrs(d: TextBox['dist']): string {
  const e = (v: number) => Math.round(v * 12700);
  const x = d ?? { t: 0, b: 0, l: 9, r: 9 };
  return `distT="${e(x.t)}" distB="${e(x.b)}" distL="${e(x.l)}" distR="${e(x.r)}"`;
}

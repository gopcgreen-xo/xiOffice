/**
 * 字节级保留的打开与保存（需求 W12–W14）。
 *
 * 打开：docx-engine（GenOffice，见 editors/word/vendor）扫描 document.xml 得到正文顶层元素的字符区间（docxIndex），
 * 本编辑器的读取器逐个元素读成模型，每块记下 src；同时留一份读入时的模型快照。
 * 保存：正文按 src 分组，和快照完全相同的组用 docx-engine 的 original 块（原文件子串照抄）；改过的组重新生成——
 * 段落引用原样式、只写与样式不同的格式，没改格式的段落 pPr 照抄，表格结构没变时只替换改过的单元格；
 * 模型里没有的元素（正文级书签等）原位照抄。其余部件（样式、编号、主题、设置、媒体、页眉页脚…）没改就逐字节保留，
 * 改了的才重写或追加。什么都没改时返回原文件字节。
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { parseDocx, saveDocx, type ParsedDocFull, type SaveBlock, type StyleUpsert } from '../../vendor/docx-engine/src/index';
import { imageData, imageExt } from '../core/images';
import { sectionList, HF_STORIES, isHeaderStory, type HFStory, type Block, type Comment, type Doc, type PageGeom, type Paragraph, type Table, type TextBox, type Cell } from '../core/model';
import { editToUpsert, removeStyleXml, withEdits } from '../core/styles';
import { fillPartXml } from '../core/merge';
import { isWatermarkChild } from '../../vendor/docx-engine/src/watermark';
import { splitXmlChildren } from '../../vendor/docx-engine/src/generate';
import { commentParts, listXml, paraPatchXml, readDocx, setPatchCtx, styleSource, tableXml, tcBordersXml, watermarkXml, type StyleSource } from './docx';

export interface DocxOrigin {
  bytes: Uint8Array;
  parsed: ParsedDocFull;
  documentXml: string;
  /** 第 i 个元素的完整 XML（含内容控件外壳） */
  full: string[];
  /** 第 i 个元素去掉内容控件外壳后的 XML；不是段落 / 表格时为 null */
  inner: (string | null)[];
  shells: Map<number, { open: string; close: string }>;
  /** docx-engine 自己处理的隐藏元素（正文末尾的 sectPr） */
  hidden: Set<number>;
  /** 读入时每个元素读出的块（快照） */
  model: Map<number, Block[]>;
  /** 页眉页脚部件 → 读入时的段落（快照）和根元素的开始 / 结束标签 */
  hf: Map<string, { json: string; open: string; close: string }>;
  comments: string;
  numIds: Set<number>;
  styles: StyleSource;
  /** 原文件里最大的修订 / 批注 id（新修订接在后面） */
  revBase: number;
  /** 原文本框 / 浮动图片：原始 XML → 读入时的模型 key */
  boxKeys: Map<string, string>;
  /** 读入时的文字水印（JSON；没有为 null） */
  watermark: string;
  /** 读入时的文档保护（JSON；没有为 null） */
  protection: string;
  /** 读入时是否处于修订状态（w:trackRevisions） */
  trackRevisions: boolean;
  /** 读入时最后一节的页面设置、奇偶页不同 */
  setup: string;
  titlePg: boolean;
}

const boxKey = (b: TextBox) => JSON.stringify({ ...b, raw: undefined });
const json = (x: unknown) => JSON.stringify(x);

function eachParagraph(blocks: Block[], fn: (p: Paragraph) => void) {
  for (const b of blocks) {
    if (b.kind === 'p') fn(b);
    else for (const r of b.rows) for (const c of r.cells) for (const p of c.paragraphs) fn(p);
  }
}

/** 打开 docx：返回模型和保存时要用的原文件信息；docx-engine 读不了时 origin 为 null（保存时整篇重写） */
export async function openDocx(bytes: Uint8Array): Promise<{ doc: Doc; origin: DocxOrigin | null }> {
  let parsed: ParsedDocFull;
  try {
    parsed = await parseDocx(bytes);
  } catch {
    return { doc: readDocx(bytes), origin: null };
  }
  const documentXml = parsed.internal.documentXml;
  const { elements } = parsed.extras;
  const byIndex = new Map(parsed.blocks.filter((b) => b.docxIndex !== null).map((b) => [b.docxIndex as number, b]));
  const shells = new Map<number, { open: string; close: string }>();
  const hidden = new Set<number>();
  const full: string[] = [];
  const inner: (string | null)[] = [];
  elements.forEach((e, i) => {
    const xml = documentXml.slice(e.start, e.end);
    full.push(xml);
    const blk = byIndex.get(i);
    if (blk?.hidden) hidden.add(i);
    let child = xml;
    if (blk?.sdtShell) {
      shells.set(i, { open: blk.sdtShell.openXml, close: blk.sdtShell.closeXml });
      child = xml.slice(blk.sdtShell.openXml.length, xml.length - blk.sdtShell.closeXml.length);
    }
    inner.push(!hidden.has(i) && e.end > e.start && (e.name === 'w:p' || e.name === 'w:tbl') ? child : null);
  });
  const doc = readDocx(bytes, { elements: inner });
  if (parsed.protection) doc.protection = parsed.protection;
  const model = new Map<number, Block[]>();
  for (const b of doc.body) if (b.src !== undefined) (model.get(b.src) ?? model.set(b.src, []).get(b.src)!).push(b);
  for (const [k, v] of model) model.set(k, structuredClone(v));

  const files = unzipSync(bytes);
  const hf = new Map<string, { json: string; open: string; close: string }>();
  const boxKeys = new Map<string, string>();
  const seen = new Set<Paragraph[]>();
  for (const s of sectionList(doc)) {
    for (const story of HF_STORIES) {
      const ps = s.hf[story];
      if (seen.has(ps)) continue;
      seen.add(ps);
      const path = ps.find((p) => p.hfPart)?.hfPart;
      if (!path || hf.has(path) || !files[path]) continue;
      const xml = strFromU8(files[path]);
      const open = xml.match(/^[\s\S]*?<w:(hdr|ftr)\b[^>]*>/)?.[0];
      if (!open) continue;
      hf.set(path, { json: json(ps), open, close: xml.slice(xml.lastIndexOf('</w:')) });
    }
  }
  const allParas: Paragraph[] = [];
  eachParagraph(doc.body, (p) => allParas.push(p));
  for (const p of allParas) for (const r of p.runs) if (r.box?.raw) boxKeys.set(r.box.raw, boxKey(r.box));
  let revBase = 0;
  for (const m of documentXml.matchAll(/<w:(?:ins|del|moveFrom|moveTo)\b[^>]*\bw:id="(\d+)"/g)) revBase = Math.max(revBase, Number(m[1]));
  const last = sectionList(doc).at(-1);
  return {
    doc,
    origin: {
      bytes,
      parsed,
      documentXml,
      full,
      inner,
      shells,
      hidden,
      model,
      hf,
      comments: json(doc.comments ?? []),
      numIds: new Set(doc.lists.filter((l) => !l.preset || l.preset.startsWith('docx-')).map((l) => l.id)),
      styles: styleSource(bytes),
      revBase: revBase + 1000,
      boxKeys,
      watermark: json(doc.watermark ?? null),
      protection: json(doc.protection ?? null),
      trackRevisions: !!doc.trackRevisions,
      setup: json(doc.setup),
      titlePg: !!last?.hf.titlePg,
    },
  };
}

// ---------------- 表格：结构没变时只替换改过的单元格 ----------------

interface Range {
  name: string;
  start: number;
  end: number;
}

/** [from, to) 里的顶层元素（不进入子元素） */
function topElements(xml: string, from: number, to: number): Range[] {
  const out: Range[] = [];
  const re = /<(\/?)([A-Za-z][\w:.-]*)((?:"[^"]*"|'[^']*'|[^"'>])*)>|<!--[\s\S]*?-->/g;
  re.lastIndex = from;
  let depth = 0;
  let cur: Range | null = null;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) && m.index < to) {
    if (!m[2]) continue;
    const closing = m[1] === '/';
    const self = !closing && m[3].endsWith('/');
    if (closing) {
      depth--;
      if (depth === 0 && cur) {
        cur.end = m.index + m[0].length;
        out.push(cur);
        cur = null;
      }
    } else if (self) {
      if (depth === 0) out.push({ name: m[2], start: m.index, end: m.index + m[0].length });
    } else {
      if (depth === 0) cur = { name: m[2], start: m.index, end: -1 };
      depth++;
    }
  }
  return out;
}

const contentOf = (xml: string, r: Range) => ({ from: xml.indexOf('>', r.start) + 1, to: xml.lastIndexOf('</', r.end - 1) });

/** 表格去掉单元格内容和单元格格式后的样子（比较结构用） */
const tableShape = (t: Table) =>
  json({ ...t, src: undefined, rows: t.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, paragraphs: undefined, shading: undefined, borders: undefined, bdr: undefined, edges: undefined, vAlign: undefined })) })) });

/** 单元格格式里本编辑器能改的几项（底纹、框线、垂直对齐） */
const cellFormat = (c: Cell) => json({ shading: c.shading, borders: c.borders, bdr: c.bdr, vAlign: c.vAlign });

/** CT_TcPr 的子元素顺序 */
const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'];

/**
 * 只改单元格格式：原 tcPr 里改过的那几个子元素（tcBorders / shd / vAlign）换成新的，其余子元素逐字节保留；没有 tcPr 时新加一个。
 */
export function patchTcPr(tcPr: string, cell: Cell, before: Cell): string {
  const open = tcPr.match(/^<w:tcPr\b[^>]*?(?=\/?>)/)?.[0] ?? '<w:tcPr';
  const inner = tcPr.endsWith('/>') ? '' : tcPr.slice(tcPr.indexOf('>') + 1, tcPr.lastIndexOf('</'));
  const kids = topElements(inner, 0, inner.length).map((r) => ({ name: r.name.replace(/^w:/, ''), xml: inner.slice(r.start, r.end) }));
  const set = (name: string, xml: string | null) => {
    const i = kids.findIndex((k) => k.name === name);
    if (i >= 0) kids.splice(i, 1);
    if (xml === null) return;
    const rank = TCPR_ORDER.indexOf(name);
    const at = kids.findIndex((k) => TCPR_ORDER.indexOf(k.name) > rank);
    kids.splice(at < 0 ? kids.length : at, 0, { name, xml });
  };
  if (json(cell.bdr) !== json(before.bdr) || cell.borders !== before.borders) {
    const sz = 4;
    set(
      'tcBorders',
      cell.bdr
        ? tcBordersXml(cell.bdr)
        : cell.borders !== undefined
          ? `<w:tcBorders>${['top', 'left', 'bottom', 'right'].map((n) => `<w:${n} w:val="${cell.borders ? 'single' : 'nil'}" w:sz="${sz}" w:space="0" w:color="auto"/>`).join('')}</w:tcBorders>`
          : null,
    );
  }
  if (cell.shading !== before.shading) set('shd', cell.shading ? `<w:shd w:val="clear" w:color="auto" w:fill="${cell.shading}"/>` : null);
  if (cell.vAlign !== before.vAlign) set('vAlign', `<w:vAlign w:val="${cell.vAlign}"/>`);
  return kids.length ? `${open}>${kids.map((k) => k.xml).join('')}</w:tcPr>` : '';
}

function tablePatchXml(t: Table, orig?: { model: Table; xml: string }): string {
  if (orig && json(t) === json(orig.model)) return orig.xml;
  if (orig && tableShape(t) === tableShape(orig.model)) {
    const xml = orig.xml;
    const body = contentOf(xml, { name: 'w:tbl', start: 0, end: xml.length });
    const rows = topElements(xml, body.from, body.to).filter((e) => e.name === 'w:tr');
    const edits: { from: number; to: number; text: string }[] = [];
    t.rows.forEach((row, r) => {
      const tr = rows[r];
      if (!tr) return;
      const rc = contentOf(xml, tr);
      const cells = topElements(xml, rc.from, rc.to).filter((e) => e.name === 'w:tc');
      row.cells.forEach((cell, c) => {
        const tc = cells[c];
        const before = orig.model.rows[r]?.cells[c];
        if (!tc || !before) return;
        const cc = contentOf(xml, tc);
        const kids = topElements(xml, cc.from, cc.to);
        if (cellFormat(cell) !== cellFormat(before)) {
          const pr = kids.find((k) => k.name === 'w:tcPr');
          const text = patchTcPr(pr ? xml.slice(pr.start, pr.end) : '<w:tcPr/>', cell, before);
          edits.push(pr ? { from: pr.start, to: pr.end, text } : { from: cc.from, to: cc.from, text });
        }
        if (json(cell.paragraphs) === json(before.paragraphs)) return;
        const paras = kids.filter((k) => k.name === 'w:p');
        // 单元格内容 = tcPr 之后到结束；段落逐个配对（个数相同时按位置）
        const startAt = kids.find((k) => k.name !== 'w:tcPr')?.start ?? cc.to;
        const pair = paras.length === before.paragraphs.length;
        const text = cell.paragraphs
          .map((p, k) => paraPatchXml(p, pair && paras[k] && before.paragraphs[k] ? { model: before.paragraphs[k], xml: xml.slice(paras[k].start, paras[k].end) } : undefined))
          .join('');
        edits.push({ from: startAt, to: cc.to, text });
      });
    });
    let out = xml;
    for (const e of edits.sort((a, b) => b.from - a.from)) out = out.slice(0, e.from) + e.text + out.slice(e.to);
    return out;
  }
  // 结构变了：重新生成；表格级格式没变时沿用原 tblPr（表格样式、条件格式等）
  let x = tableXml(t);
  if (orig) {
    const same = json({ align: t.align, indent: t.indent, borders: t.borders, borderSize: t.borderSize, cellMargin: t.cellMargin }) === json({ align: orig.model.align, indent: orig.model.indent, borders: orig.model.borders, borderSize: orig.model.borderSize, cellMargin: orig.model.cellMargin });
    const tblPr = orig.xml.match(/<w:tblPr\b[\s\S]*?<\/w:tblPr>/)?.[0];
    if (same && tblPr) x = x.replace(/<w:tblPr\b[\s\S]*?<\/w:tblPr>/, () => tblPr.replace(/<w:tblW\b[^>]*\/>/, `<w:tblW w:w="${Math.round(t.cols.reduce((a, c) => a + c, 0) * 20)}" w:type="dxa"/>`));
  }
  return x;
}

// ---------------- 节属性 ----------------

const tw = (pt: number) => Math.round(pt * 20);

/** 改 sectPr 的纸张、页边距（只动 pgSz / pgMar 的这几个属性） */
function patchGeom(sectPr: string, g: PageGeom): string {
  const set = (tag: string, attrs: Record<string, number>, s: string) => {
    const re = new RegExp(`<w:${tag}\\b[^>]*?/>`);
    const old = s.match(re)?.[0];
    let el = old ?? `<w:${tag}/>`;
    for (const [k, v] of Object.entries(attrs)) {
      const a = new RegExp(`\\sw:${k}="[^"]*"`);
      el = a.test(el) ? el.replace(a, ` w:${k}="${v}"`) : el.replace(/\s*\/>$/, ` w:${k}="${v}"/>`);
    }
    return old ? s.replace(old, el) : s.replace(/(<w:sectPr\b[^>]*>)/, `$1${el}`);
  };
  const landscape = g.width > g.height;
  let out = set('pgSz', { w: tw(g.width), h: tw(g.height) }, sectPr);
  out = landscape ? (/w:orient=/.test(out) ? out.replace(/w:orient="[^"]*"/, 'w:orient="landscape"') : out.replace(/<w:pgSz\b/, '<w:pgSz w:orient="landscape"')) : out.replace(/\s*w:orient="landscape"/, '');
  // 左页边距在模型里含装订线：写回时扣掉原来的装订线
  const gutter = Number(sectPr.match(/<w:pgMar\b[^>]*w:gutter="(\d+)"/)?.[1] ?? 0);
  return set('pgMar', { top: tw(g.marginTop), bottom: tw(g.marginBottom), left: Math.max(0, tw(g.marginLeft) - gutter), right: tw(g.marginRight), header: tw(g.headerDistance), footer: tw(g.footerDistance) }, out);
}

const geomOf = (s: PageGeom): PageGeom => ({ width: s.width, height: s.height, marginTop: s.marginTop, marginBottom: s.marginBottom, marginLeft: s.marginLeft, marginRight: s.marginRight, headerDistance: s.headerDistance, footerDistance: s.footerDistance });

// ---------------- 命名空间 ----------------

/** 生成的 XML 可能用到的前缀 */
const NS_URI: Record<string, string> = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  w10: 'urn:schemas-microsoft-com:office:word',
};

/** 开始标签上声明的命名空间 */
function declaredNs(openTag: string): Map<string, string> {
  return new Map([...openTag.matchAll(/ xmlns:([\w.-]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
}

/** 生成的块用到了根元素没声明的前缀（原文件根元素上声明得少时）：在块自己的开始标签上补声明 */
function withNs(xml: string, declared: Map<string, string>): string {
  const need = Object.entries(NS_URI).filter(([p]) => !declared.has(p) && !xml.includes(`xmlns:${p}=`) && new RegExp(`[<\\s]${p}:`).test(xml));
  if (!need.length) return xml;
  return xml.replace(/^<[\w:]+/, (m) => m + need.map(([p, u]) => ` xmlns:${p}="${u}"`).join(''));
}

// ---------------- 保存 ----------------

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_W = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** 字节级保留保存：只重写改过的部分 */
/** 页眉部件 XML：去掉里面原有的水印段落，把 wm（空串 = 不要水印）放在最前面 */
export function setWatermarkXml(xml: string, wm: string): string {
  const open0 = xml.match(/^[\s\S]*?<w:hdr\b[^>]*>/)?.[0];
  if (!open0) return xml;
  const selfClosed = open0.endsWith('/>');
  const open = selfClosed ? `${open0.slice(0, -2)}>` : open0;
  const end = selfClosed ? xml.length : xml.lastIndexOf('</w:hdr>');
  const inner = selfClosed ? '' : xml.slice(open0.length, end);
  const kept = splitXmlChildren(inner).filter((k) => !isWatermarkChild(k));
  const body = (wm ? withNs(wm, declaredNs(open)) : '') + kept.map((k) => k.xml).join('');
  return `${open}${body}</w:hdr>`;
}

/** 一个部件新增的关系：新图片、新外部链接 */
interface RelSet {
  images: { id: string; rid: string }[];
  links: { url: string; rid: string }[];
}

/** 保存过程中的状态：各部件新增的关系（null = 正文），cur 为正在写的部件 */
interface SaveState {
  rels: Map<string | null, RelSet>;
  cur: string | null;
  /** 原文件里没有页眉 / 页脚的中间节：分节段落 → 要加进它 sectPr 的页眉页脚引用 */
  midRefs: Map<Paragraph, string>;
  /** 在 cur 这个部件里登记一张图片，返回关系 id（水印要用） */
  newImage: (id: string) => string | null;
}

export async function patchDocx(doc: Doc, origin: DocxOrigin): Promise<Uint8Array> {
  const st: SaveState = { rels: new Map(), cur: null, midRefs: new Map(), newImage: () => null };
  const relSet = () => st.rels.get(st.cur) ?? st.rels.set(st.cur, { images: [], links: [] }).get(st.cur)!;
  const missingStyles = new Map<string, string>();
  const catalog = withEdits(origin.styles, doc.styleEdits);
  let relSeq = 0;
  // A saved document may already contain our rIdWz ids. Reopening and editing
  // it again must allocate fresh ids, otherwise Word rejects duplicate relations.
  const usedRelIds = new Set<string>();
  for (const [name, data] of Object.entries(unzipSync(origin.bytes))) {
    if (name.endsWith('.rels')) for (const m of strFromU8(data).matchAll(/\bId="([^"]+)"/g)) usedRelIds.add(m[1]);
  }
  const nextRelId = () => {
    let id: string;
    do { id = `rIdWz${++relSeq}`; } while (usedRelIds.has(id));
    usedRelIds.add(id);
    return id;
  };
  const newImage = (id: string): string | null => {
    const set = relSet();
    const hit = set.images.find((x) => x.id === id);
    if (hit) return hit.rid;
    if (!imageData(id)) return null;
    const rid = nextRelId();
    set.images.push({ id, rid });
    return rid;
  };
  st.newImage = newImage;
  const elementSect = (i: number) => origin.inner[i]?.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0] ?? null;
  setPatchCtx({
    base: (p, r) => catalog.base(p, r),
    styleIdOf: (c) => {
      const id = origin.styles.idOf(c);
      if (id) return id;
      missingStyles.set(c, c);
      return c;
    },
    newImageRel: (id) => newImage(id),
    newLinkRel: (url) => {
      const set = relSet();
      const hit = set.links.find((x) => x.url === url);
      if (hit) return hit.rid;
      const rid = nextRelId();
      set.links.push({ url, rid });
      return rid;
    },
    boxUnchanged: (b) => !!b.raw && origin.boxKeys.get(b.raw) === boxKey(b),
    sectPr: (p) => {
      if (p.src === undefined) return null;
      const s = elementSect(p.src);
      if (!s) return null;
      const refs = st.midRefs.get(p);
      const geom = p.sectionPage ? patchGeom(s, p.sectionPage) : s;
      return refs ? geom.replace(/(<w:sectPr\b[^>]*>)/, `$1${refs}`) : geom;
    },
    revBase: origin.revBase,
  });
  try {
    return await save(doc, origin, st, missingStyles);
  } finally {
    setPatchCtx(null);
  }
}

/** 页眉页脚分区在 OOXML 中的类型 */
const hfTypeOf = (story: HFStory) => (story.endsWith('First') ? 'first' : story.endsWith('Even') ? 'even' : 'default');
const isAbsentHf = (ps: Paragraph[]) => ps.length === 1 && !!ps[0].hfNone && !ps[0].runs.length;

interface NewPart {
  ps: Paragraph[];
  tag: 'hdr' | 'ftr';
  /** 包内路径，如 word/wz-header1.xml */
  file: string;
  rid: string;
}

/**
 * 原文件里没有这个页眉 / 页脚、现在写了内容：要新建部件。引用加在这个数组第一次出现的那一节的 sectPr 里
 * （之后共用同一数组的节在 Word 里就是“与上一节相同”，不需要自己的引用）。
 * 返回新部件和最后一节要加的引用；中间节的引用记在 st.midRefs。
 */
function planNewParts(doc: Doc, st: SaveState, wm: boolean): { parts: Map<Paragraph[], NewPart>; trailingRefs: string } {
  const parts = new Map<Paragraph[], NewPart>();
  let trailingRefs = '';
  for (const s of sectionList(doc)) {
    let refs = '';
    for (const story of HF_STORIES) {
      const ps = s.hf[story];
      if (ps.some((p) => p.hfPart) || parts.has(ps)) continue;
      // 原来没有的页眉：写了内容才建；要加水印时，会显示的页眉（默认；首页不同时的首页；奇偶页不同时的偶数页）也要建
      const needWm = wm && isHeaderStory(story) && (story === 'header' || (story === 'headerFirst' ? s.hf.titlePg : doc.setup.evenAndOdd));
      if (isAbsentHf(ps) && !needWm) continue;
      const tag = isHeaderStory(story) ? 'hdr' : 'ftr';
      const n = [...parts.values()].filter((x) => x.tag === tag).length + 1;
      const word = tag === 'hdr' ? 'header' : 'footer';
      const part: NewPart = { ps, tag, file: `word/wz-${word}${n}.xml`, rid: `rIdWz${tag === 'hdr' ? 'Hdr' : 'Ftr'}${n}` };
      parts.set(ps, part);
      refs += `<w:${word}Reference w:type="${hfTypeOf(story)}" r:id="${part.rid}"/>`;
    }
    if (!refs) continue;
    if (s.index === null) trailingRefs = refs;
    else st.midRefs.set(doc.body[s.index] as Paragraph, refs);
  }
  return { parts, trailingRefs };
}

async function save(doc: Doc, origin: DocxOrigin, st: SaveState, missingStyles: Map<string, string>): Promise<Uint8Array> {
  const wmChanged = json(doc.watermark ?? null) !== origin.watermark;
  const plan = planNewParts(doc, st, wmChanged && !!doc.watermark);
  // ---- 正文 ----
  const groups: { src?: number; blocks: Block[] }[] = [];
  for (const b of doc.body) {
    const last = groups.at(-1);
    if (last && b.src !== undefined && last.src === b.src) last.blocks.push(b);
    else groups.push({ src: b.src, blocks: [b] });
  }
  // 模型里没有的元素（正文级书签等）：跟在它前面最近的段落 / 表格元素后面照抄
  const passive: number[] = [];
  origin.full.forEach((_, i) => {
    if (origin.inner[i] === null && !origin.hidden.has(i)) passive.push(i);
  });
  const ownerOf = (i: number) => {
    for (let k = i - 1; k >= 0; k--) if (origin.inner[k] !== null) return k;
    return -1;
  };
  const emitted = new Set<number>();
  const out: SaveBlock[] = [];
  const emitPassive = (owner: number) => {
    for (const i of passive) if (!emitted.has(i) && ownerOf(i) === owner) (emitted.add(i), out.push({ kind: 'original', docxIndex: i }));
  };
  emitPassive(-1);
  const rootNs = declaredNs(origin.documentXml.match(/<w:document\b[^>]*>/)?.[0] ?? '');
  const genBlock = (b: Block, orig: Block | undefined, xml: string | undefined) =>
    withNs(
      b.kind === 'p'
        ? paraPatchXml(b, orig?.kind === 'p' && xml ? { model: orig, xml } : undefined)
        : tablePatchXml(b, orig?.kind === 'table' && xml ? { model: orig, xml } : undefined),
      rootNs,
    );
  for (const g of groups) {
    if (g.src !== undefined && !emitted.has(g.src)) {
      const src = g.src;
      emitted.add(src);
      const orig = origin.model.get(src);
      if (orig && json(g.blocks) === json(orig)) out.push({ kind: 'original', docxIndex: src });
      else {
        const single = orig?.length === 1 && g.blocks.length >= 1;
        const xml = g.blocks.map((b, k) => genBlock(b, orig?.[k], single && k === 0 ? (origin.inner[src] ?? undefined) : undefined)).join('');
        const shell = origin.shells.get(src);
        out.push({ kind: 'xml', xml: shell ? shell.open + xml + shell.close : xml, docxIndex: src });
      }
      emitPassive(src);
    } else out.push({ kind: 'xml', xml: g.blocks.map((b) => genBlock(b, undefined, undefined)).join('') });
  }
  for (const i of passive) if (!emitted.has(i)) out.push({ kind: 'original', docxIndex: i });

  // ---- 页眉页脚：内容变了的部件按原根元素重新写 ----
  const partXml: Record<string, string> = {};
  const done = new Set<Paragraph[]>();
  for (const s of sectionList(doc)) {
    for (const story of HF_STORIES) {
      const ps = s.hf[story];
      if (done.has(ps)) continue;
      done.add(ps);
      const path = ps.find((p) => p.hfPart)?.hfPart;
      const o = path ? origin.hf.get(path) : undefined;
      if (path && o && json(ps) !== o.json) {
        const ns = declaredNs(o.open.match(/<w:(?:hdr|ftr)\b[^>]*>/)?.[0] ?? '');
        st.cur = path;
        partXml[path] = o.open + ps.map((p) => withNs(p.hfTable ? p.hfTable.xml : paraPatchXml(p), ns)).join('') + o.close;
        st.cur = null;
      }
    }
  }

  // ---- 新建的页眉 / 页脚部件（原文件里没有，现在写了内容） ----
  const newPartXml = new Map<string, string>();
  const W_NS = new Map([
    ['w', NS_URI.w],
    ['r', NS_URI.r],
  ]);
  for (const np of plan.parts.values()) {
    st.cur = np.file;
    const wmXml = wmChanged && doc.watermark && np.tag === 'hdr' ? withNs(watermarkXml(doc.watermark, doc.setup, (id) => st.newImage(id)), W_NS) : '';
    const body = wmXml + np.ps.map((p) => withNs(p.hfTable ? p.hfTable.xml : paraPatchXml(p), W_NS)).join('');
    st.cur = null;
    newPartXml.set(np.file, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:${np.tag} xmlns:w="${NS_URI.w}" xmlns:r="${NS_URI.r}">${body}</w:${np.tag}>`);
  }

  // ---- 最后一节的页面设置、首页不同 ----
  const options: Parameters<typeof saveDocx>[2] = { partXml };
  const trailing = origin.parsed.blocks.find((b) => b.hidden && b.originalXml?.includes('<w:sectPr'))?.originalXml;
  const lastHf = sectionList(doc).at(-1)?.hf;
  if (trailing) {
    let sect = trailing;
    if (json(doc.setup) !== origin.setup) sect = patchGeom(sect, geomOf(doc.setup));
    if (lastHf && !!lastHf.titlePg !== origin.titlePg)
      sect = lastHf.titlePg ? sect.replace(/(<w:sectPr\b[^>]*>[\s\S]*?)(<w:pgSz\b)/, '$1<w:titlePg/>$2') : sect.replace(/<w:titlePg\b[^>]*\/>/, '');
    if (plan.trailingRefs) sect = sect.replace(/(<w:sectPr\b[^>]*>)/, `$1${plan.trailingRefs}`);
    if (sect !== trailing) options.trailingSectPr = sect;
  }
  // ---- 水印：变了才交给 docx-engine 写（它负责默认页眉里的水印段落） ----
  // ---- 文档保护：变了才交给 docx-engine 写 settings.xml ----
  const protectionChanged = json(doc.protection ?? null) !== origin.protection;
  if (protectionChanged) options.protection = doc.protection ?? null;
  // 修订开关（审阅状态）：变了才改 settings.xml
  const trackChanged = !!doc.trackRevisions !== origin.trackRevisions;
  const evenOdd = JSON.parse(origin.setup).evenAndOdd;
  if (!!doc.setup.evenAndOdd !== !!evenOdd) options.evenAndOddHeaders = !!doc.setup.evenAndOdd;

  // ---- 批注：变了才按模型重写 ----
  const comments = doc.comments ?? [];
  const commentsChanged = json(comments) !== origin.comments;

  // ---- 编号：新列表追加到 numbering.xml ----
  const newLists = doc.lists.filter((l) => !origin.numIds.has(l.id) && usesList(doc, l.id));
  const edited = Object.values(doc.styleEdits ?? {}).map(editToUpsert);
  if (missingStyles.size || edited.length) {
    const upserts: StyleUpsert[] = [...missingStyles.keys()].map((c) => {
      const h = c.match(/^Heading(\d)$/);
      const t = c.match(/^TOC(\d)$/);
      return { styleId: c, type: 'paragraph', name: h ? `heading ${h[1]}` : t ? `toc ${t[1]}` : c };
    });
    options.styleUpserts = [...upserts, ...edited];
  }

  // 什么都没改：直接给原文件（docx-engine 合并域代码段落后，自己的“未改动”判断可能对不上）
  if (out.every((b) => b.kind === 'original') && Object.keys(partXml).length === 0 && !options.trailingSectPr && options.evenAndOddHeaders === undefined && !commentsChanged && !newLists.length && !missingStyles.size && !edited.length && !st.rels.size && !plan.parts.size && !wmChanged && !protectionChanged && !trackChanged && !doc.xmlFill && !doc.styleDeletes?.length) {
    const order = out.map((b) => (b.kind === 'original' ? b.docxIndex : -1));
    const all = origin.full.map((_, i) => i).filter((i) => !origin.hidden.has(i));
    if (json(order) === json(all)) return origin.bytes;
  }
  let bytes = await saveDocx(origin.parsed, out, options);

  // ---- 补充 docx-engine 不管的部件：新图片、编号、批注；原来就没被引用的关系照样保留 ----
  const files = unzipSync(bytes);
  const orig = unzipSync(origin.bytes);
  let notesEdited = false;
  // 水印变了：所有会显示的、原文件里已有的页眉部件里去掉旧水印、放上新的（新建的部件在上面已经带上了）
  if (wmChanged) {
    const done = new Set<string>();
    for (const s of sectionList(doc)) {
      for (const story of ['header', 'headerFirst', 'headerEven'] as const) {
        if (story === 'headerFirst' && !s.hf.titlePg) continue;
        if (story === 'headerEven' && !doc.setup.evenAndOdd) continue;
        const path = s.hf[story].find((p) => p.hfPart)?.hfPart;
        if (!path || done.has(path) || !files[path]) continue;
        done.add(path);
        st.cur = path;
        const wm = doc.watermark ? watermarkXml(doc.watermark, s.setup, (id) => st.newImage(id)) : '';
        st.cur = null;
        files[path] = strToU8(setWatermarkXml(strFromU8(files[path]), wm));
        notesEdited = true;
      }
    }
  }
  // 填过数据（模板与批量生成）：脚注 / 尾注里的 {{字段}} 对原始 XML 做同样的替换
  // 删掉的自定义样式：从 styles.xml 去掉（别的样式 basedOn 它的改成它的上一级）
  if (doc.styleDeletes?.length && files['word/styles.xml']) {
    let sx = strFromU8(files['word/styles.xml']);
    for (const d of doc.styleDeletes) sx = removeStyleXml(sx, d.id, d.parent);
    files['word/styles.xml'] = strToU8(sx);
    notesEdited = true;
  }
  if (trackChanged && files['word/settings.xml']) {
    files['word/settings.xml'] = strToU8(setTrackRevisionsXml(strFromU8(files['word/settings.xml']), !!doc.trackRevisions));
    notesEdited = true;
  }
  if (doc.xmlFill) {
    const { values, blank } = doc.xmlFill;
    for (const path of ['word/footnotes.xml', 'word/endnotes.xml']) {
      if (!files[path]) continue;
      const before = strFromU8(files[path]);
      const after = fillPartXml(before, (k) => (Object.prototype.hasOwnProperty.call(values, k) ? values[k] : blank ? '' : undefined), new Set());
      if (after !== before) {
        files[path] = strToU8(after);
        notesEdited = true;
      }
    }
  }
  const relsPath = 'word/_rels/document.xml.rels';
  const origRels = orig[relsPath] ? strFromU8(orig[relsPath]) : null;
  let rels = files[relsPath] ? strFromU8(files[relsPath]) : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  let restored = false;
  if (origRels && origRels !== rels) {
    const tags = (x: string) => x.match(/<Relationship\b[^>]*\/>/g) ?? [];
    const now = new Set(tags(rels));
    const before = new Set(tags(origRels));
    const idOf = (t: string) => t.match(/\bId="([^"]*)"/)?.[1] ?? '';
    // 原文件里这个关系有没有被引用（正文里任何 r: 属性的值等于这个 id）
    const referenced = (id: string) => [...origin.documentXml.matchAll(/\br:[A-Za-z]+="([^"]*)"/g)].some((m) => m[1] === id);
    const dangling = [...before].filter((t) => !now.has(t) && !referenced(idOf(t)));
    if (dangling.length) {
      // 从原关系表出发：去掉被用户删掉引用的（docx-engine 清理的那些），加上新增的
      const removed = [...before].filter((t) => !now.has(t) && !dangling.includes(t));
      let x = origRels;
      for (const t of removed) x = x.replace(t, '');
      const added = [...now].filter((t) => !before.has(t));
      if (added.length) x = x.replace('</Relationships>', `${added.join('')}</Relationships>`);
      rels = x;
      restored = true;
      // 被一起删掉的媒体文件也放回去
      for (const t of dangling) {
        if (/TargetMode="External"/.test(t)) continue;
        const target = t.match(/\bTarget="([^"]*)"/)?.[1];
        const path = target ? (target.startsWith('/') ? target.slice(1) : `word/${target}`) : '';
        if (path && orig[path] && !files[path]) files[path] = orig[path];
      }
    }
  }
  if (!st.rels.size && !commentsChanged && !newLists.length && !restored && !plan.parts.size && !notesEdited) return bytes;
  let ct = strFromU8(files['[Content_Types].xml']);
  const addRel = (id: string, type: string, target: string) => (rels = rels.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`));
  const addDefault = (ext: string, type: string) => {
    if (!new RegExp(`Extension="${ext}"`, 'i').test(ct)) ct = ct.replace('</Types>', `<Default Extension="${ext}" ContentType="${type}"/></Types>`);
  };
  const addOverride = (part: string, type: string) => {
    if (!ct.includes(`PartName="/${part}"`)) ct = ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  };
  const relTarget = (type: string) => rels.match(new RegExp(`<Relationship\\b[^>]*Type="${REL_NS}/${type}"[^>]*Target="([^"]+)"`))?.[1];

  for (const np of plan.parts.values()) {
    files[np.file] = strToU8(newPartXml.get(np.file)!);
    const word = np.tag === 'hdr' ? 'header' : 'footer';
    addRel(np.rid, `${REL_NS}/${word}`, np.file.slice('word/'.length));
    addOverride(np.file, `${CT_W}.${word}+xml`);
  }
  // 新图片、新外部链接：正文的关系写进 document.xml.rels，页眉页脚部件的写进它自己的关系文件
  const stamp = Date.now().toString(36);
  const mediaOf = new Map<string, string>();
  const media = (id: string) => {
    let name = mediaOf.get(id);
    if (!name) {
      const d = imageData(id)!;
      const ext = imageExt(d.mime);
      name = `media/wz-image${stamp}-${mediaOf.size + 1}.${ext}`;
      mediaOf.set(id, name);
      files[`word/${name}`] = new Uint8Array(d.data);
      addDefault(ext, d.mime);
    }
    return name;
  };
  const attrEsc = (v: string) => v.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  for (const [part, set] of st.rels) {
    const entries = [
      ...set.images.map((x) => `<Relationship Id="${x.rid}" Type="${REL_NS}/image" Target="${media(x.id)}"/>`),
      ...set.links.map((x) => `<Relationship Id="${x.rid}" Type="${REL_NS}/hyperlink" Target="${attrEsc(x.url)}" TargetMode="External"/>`),
    ].join('');
    if (!entries) continue;
    if (part === null) {
      rels = rels.replace('</Relationships>', `${entries}</Relationships>`);
      continue;
    }
    const slash = part.lastIndexOf('/');
    const relPath = `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
    const cur = files[relPath] ? strFromU8(files[relPath]) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    files[relPath] = strToU8(cur.replace('</Relationships>', `${entries}</Relationships>`));
  }

  if (newLists.length) {
    const target = relTarget('numbering');
    const path = target ? `word/${target.replace(/^\/?word\//, '')}` : 'word/numbering.xml';
    let xml = files[path] ? strFromU8(files[path]) : '';
    if (!xml) {
      xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`;
      addRel('rIdWzNumbering', `${REL_NS}/numbering`, 'numbering.xml');
      addOverride('word/numbering.xml', `${CT_W}.numbering+xml`);
    }
    let abs = Math.max(-1, ...[...xml.matchAll(/w:abstractNumId="(\d+)"/g)].map((m) => Number(m[1])));
    const parts = newLists.map((l) => listXml(l, ++abs));
    // 所有 abstractNum 必须在 num 前面
    const firstNum = xml.search(/<w:num\b/);
    const absXml = parts.map((p) => p.abstract).join('');
    xml = firstNum >= 0 ? xml.slice(0, firstNum) + absXml + xml.slice(firstNum) : xml.replace('</w:numbering>', `${absXml}</w:numbering>`);
    xml = xml.replace('</w:numbering>', `${parts.map((p) => p.num).join('')}</w:numbering>`);
    files[path] = strToU8(xml);
  }

  if (commentsChanged) {
    const written = writtenComments(doc, comments);
    const { comments: cx, ext } = commentParts(written);
    const target = relTarget('comments');
    const path = target ? `word/${target.replace(/^\/?word\//, '')}` : 'word/comments.xml';
    if (cx) {
      files[path] = strToU8(cx);
      if (!target) {
        addRel('rIdWzComments', `${REL_NS}/comments`, 'comments.xml');
        addOverride('word/comments.xml', `${CT_W}.comments+xml`);
      }
    } else if (files[path]) files[path] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`);
    const exTarget = rels.match(/<Relationship\b[^>]*Type="http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/commentsExtended"[^>]*Target="([^"]+)"/)?.[1];
    const exPath = exTarget ? `word/${exTarget.replace(/^\/?word\//, '')}` : 'word/commentsExtended.xml';
    if (ext) {
      files[exPath] = strToU8(ext);
      if (!exTarget) {
        addRel('rIdWzCommentsEx', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml');
        addOverride('word/commentsExtended.xml', `${CT_W}.commentsExtended+xml`);
      }
    }
  }
  files[relsPath] = strToU8(rels);
  files['[Content_Types].xml'] = strToU8(ct);
  bytes = zipSync(files);
  return bytes;
}

function usesList(doc: Doc, id: number): boolean {
  let used = false;
  eachParagraph(doc.body, (p) => {
    if (p.props.num?.id === id) used = true;
  });
  return used;
}

/** 有范围标记的批注（含回复）才写 */
function writtenComments(doc: Doc, comments: Comment[]): Comment[] {
  const anchored = new Set<string>();
  eachParagraph(doc.body, (p) => {
    for (const r of p.runs) if (r.mark?.kind === 'cs') anchored.add(r.mark.id);
  });
  const rootOf = (c: Comment): string => (c.parent && comments.some((x) => x.id === c.parent) ? rootOf(comments.find((x) => x.id === c.parent)!) : c.id);
  return comments.filter((c) => anchored.has(rootOf(c)));
}

/** CT_Settings 里排在 w:trackRevisions 后面的元素（OOXML 规定了子元素顺序，新元素要插在它们前面） */
const AFTER_TRACK = 'doNotTrackMoves doNotTrackFormatting documentProtection autoFormatOverride styleLockTheme styleLockQFSet defaultTabStop autoHyphenation consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps showEnvelope summaryLength clickAndTypeStyle defaultTableStyle evenAndOddHeaders bookFoldRevPrinting bookFoldPrinting bookFoldPrintingSheets drawingGridHorizontalSpacing drawingGridVerticalSpacing displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGridOrigin drawingGridHorizontalOrigin drawingGridVerticalOrigin doNotShadeFormData noPunctuationKerning characterSpacingControl printTwoOnOne strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore savePreviewPicture doNotValidateAgainstSchema saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr compat docVars rsids mathPr attachedSchema themeFontLang clrSchemeMapping doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade captions readModeInkLockDown smartTagType schemaLibrary shapeDefaults doNotEmbedSmartTags decimalSymbol listSeparator'.split(' ');

/** 打开 / 关闭 settings.xml 的 w:trackRevisions（Word 的“修订”开关，存在文件里，下次打开仍在修订状态） */
export function setTrackRevisionsXml(xml: string, on: boolean): string {
  const out = xml.replace(/<w:trackRevisions\b[^>]*\/>|<w:trackRevisions\b[^>]*>[\s\S]*?<\/w:trackRevisions>/g, '');
  if (!on) return out;
  // 插在第一个排在它后面的元素前面（扩展命名空间 w14 / w15 / w16… 的元素都在最后）
  const re = new RegExp(`<(?:w:(?:${AFTER_TRACK.join('|')})|w1\\d[a-z]*:[A-Za-z]+)[\\s/>]`, 'g');
  re.lastIndex = Math.max(0, out.search(/<w:settings\b/));
  const m = re.exec(out);
  if (m) return `${out.slice(0, m.index)}<w:trackRevisions/>${out.slice(m.index)}`;
  return out.replace('</w:settings>', '<w:trackRevisions/></w:settings>');
}

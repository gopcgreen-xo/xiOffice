/**
 * 保存：在原 xlsx 上修补（fflate 解压 → 只改编辑器建模的部分 → 压缩），其它部件与节点原样保留。
 *
 * 改过的工作表：重写 `<sheetData>`、`<cols>`、`<mergeCells>`、`<dimension>`、冻结窗格；按插入 / 删除行列的记录平移
 * 条件格式、数据验证、超链接、自动筛选、分页符的区域与公式，以及 workbook.xml 里的定义名称。
 * 文字写成共享字符串（追加到 sharedStrings.xml，原有条目不动）；styles.xml 只追加条目；
 * 有公式时设 fullCalcOnLoad，让 Excel 打开时全部重算；删除 calcChain.xml（格子变了，旧的计算链会让 Excel 报错）。
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { syncDrawings } from './charts';
import { syncPivot } from './pivotParts';
import type { PivotParts } from './pivot';
import { REL, relTarget } from './parts';
import { syncNotes } from './notes';
import { readDrawings } from './drawing';
import type { Pkg } from './parts';
import { colName, resolvePart, unescapeXml, xmlAttr, xmlText } from '@wz/editor-shared/xlsx';
import { mapRefs, quoteSheet, renameSheetRefs, shiftFormula, shiftSqref, withinBand } from './formula/refs';
import { toFileFormula } from './formula/parse';
import { FILE_ERRORS } from './formula/eval';
import { dxfForColor, writeStyles, BLANK_STYLES_XML, type StyleBook } from './styles';
import { parseRef, pxToChars, type SCell, type SheetData, type Workbook } from './workbook';
import { cfRuleXml, rangeRef } from './cfEdit';
import { detachDamagedPrinterSettings } from './printerSettings';

const esc = (s: string) =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_SHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CT_SST = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
const REL_SST = `${NS_REL}/sharedStrings`;

/** 新建 / 从 .xls、.csv 转来的工作簿用的最小 xlsx 包（工作表按 xl/worksheets/sheetN.xml 命名） */
export function blankPackage(names: string[]): Record<string, Uint8Array> {
  const sheets = names.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${CT_SHEET}"/>`).join('');
  const files: Record<string, string> = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets}</Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><bookViews><workbookView/></bookViews><sheets>${names
      .map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join('')}</sheets><calcPr calcId="191029"/></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
      .join('')}<Relationship Id="rId${names.length + 1}" Type="${NS_REL}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': BLANK_STYLES_XML,
  };
  names.forEach((_, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"${i === 0 ? ' tabSelected="1"' : ''}/></sheetViews><sheetFormatPr defaultRowHeight="14.25"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
  });
  const out: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) out[k] = strToU8(v);
  return out;
}

// ---------- 通用：按 schema 顺序放元素 ----------

const WS_ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges',
  'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting',
  'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
  'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF',
  'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
];

const elRe = (tag: string) => new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`);

/** 替换 worksheet 的子元素；没有时插到按 schema 顺序该在的位置；content 为空串时删除 */
export function putElement(xml: string, tag: string, content: string, order = WS_ORDER, root = 'worksheet'): string {
  const re = elRe(tag);
  if (re.test(xml)) return xml.replace(re, () => content);
  if (!content) return xml;
  for (const next of order.slice(order.indexOf(tag) + 1)) {
    const m = xml.match(new RegExp(`<${next}\\b`));
    if (m?.index !== undefined) return xml.slice(0, m.index) + content + xml.slice(m.index);
  }
  return xml.replace(new RegExp(`</${root}>\\s*$`), `${content}</${root}>`);
}

// ---------- 共享字符串 ----------

class SharedStrings {
  private items: string[];
  private index = new Map<string, number>();
  private added = false;
  constructor(private xml: string) {
    this.items = [...xml.matchAll(/<si>[\s\S]*?<\/si>|<si\/>/g)].map((m) => m[0]);
    this.items.forEach((si, i) => {
      // 只复用纯文本条目（富文本的格式不能套到新文字上）
      const plain = si.match(/^<si><t(?:\s[^>]*)?>([\s\S]*?)<\/t><\/si>$/);
      if (plain) {
        const text = unescapeXml(plain[1]);
        if (!this.index.has(text)) this.index.set(text, i);
      }
    });
  }
  /** 原文件第 i 项的纯文字（核对沿用的序号是否还对得上） */
  textAt(i: number): string | undefined {
    const si = this.items[i];
    return si === undefined ? undefined : xmlText(si.replace(/^<si>|<\/si>$/g, ''));
  }
  get(text: string): number {
    let i = this.index.get(text);
    if (i === undefined) {
      i = this.items.length;
      const space = /^\s|\s$|\n/.test(text) ? ' xml:space="preserve"' : '';
      this.items.push(`<si><t${space}>${esc(text)}</t></si>`);
      this.index.set(text, i);
      this.added = true;
    }
    return i;
  }
  get changed() {
    return this.added;
  }
  get size() {
    return this.items.length;
  }
  toXml(): string {
    const head = this.xml.match(/^[\s\S]*?<sst\b[^>]*?>/)?.[0];
    const body = this.items.join('');
    if (!head) return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" uniqueCount="${this.items.length}">${body}</sst>`;
    // count（引用总数）是可选属性，去掉；uniqueCount 按实际条目数
    const open = head
      .replace(/\scount="\d*"/, '')
      .replace(/\suniqueCount="\d*"/, '')
      .replace(/<sst\b/, `<sst uniqueCount="${this.items.length}"`)
      .replace(/\/>$/, '>');
    const tail = this.xml.match(/<extLst\b[\s\S]*<\/extLst>/)?.[0] ?? '';
    return `${open}${body}${tail}</sst>`;
  }
}

// ---------- 工作表 ----------

function cellXml(ref: string, cell: SCell, sst: SharedStrings): string {
  const s = cell.s ? ` s="${cell.s}"` : '';
  const xa = cell.xa ? ` ${cell.xa}` : '';
  const v = cell.v;
  const ff = cell.f ? esc(toFileFormula(cell.f)) : '';
  const f = cell.fraw ? cell.fraw : cell.f ? (cell.fa ? `<f t="array" ref="${cell.fa}">${ff}</f>` : `<f>${ff}</f>`) : '';
  if (f) {
    // 没算出来的、或者是编辑器自己的错误值（#CIRC!）：不写缓存值，Excel 打开时会重算
    if (v === '' || v === undefined || (cell.k === 'e' && !FILE_ERRORS.has(String(v)))) return `<c r="${ref}"${s}${xa}>${f}</c>`;
    if (cell.k === 'e') return `<c r="${ref}"${s} t="e"${xa}>${f}<v>${esc(String(v))}</v></c>`;
    if (typeof v === 'string') return `<c r="${ref}"${s} t="str"${xa}>${f}<v>${esc(v)}</v></c>`;
    if (typeof v === 'boolean') return `<c r="${ref}"${s} t="b"${xa}>${f}<v>${v ? 1 : 0}</v></c>`;
    return `<c r="${ref}"${s}${xa}>${f}<v>${v}</v></c>`;
  }
  if (cell.k === 'e') return `<c r="${ref}"${s} t="e"${xa}><v>${esc(String(v))}</v></c>`;
  if (cell.k === 'd') return `<c r="${ref}"${s} t="d"${xa}><v>${esc(String(v))}</v></c>`;
  if (typeof v === 'boolean') return `<c r="${ref}"${s} t="b"${xa}><v>${v ? 1 : 0}</v></c>`;
  if (typeof v === 'number') return `<c r="${ref}"${s}${xa}><v>${Number.isFinite(v) ? v : 0}</v></c>`;
  const idx = cell.ss !== undefined && sst.textAt(cell.ss) === v ? cell.ss : sst.get(v);
  return `<c r="${ref}"${s} t="s"${xa}><v>${idx}</v></c>`;
}

function sheetDataXml(sheet: SheetData, sst: SharedStrings): { xml: string; dim: string } {
  const rowSet = new Set<number>();
  sheet.cells.forEach((row, r) => row && row.some(Boolean) && rowSet.add(r));
  sheet.blankStyles?.forEach((row, r) => row && row.some((x) => x !== undefined) && rowSet.add(r));
  for (const k of Object.keys(sheet.rowDefs ?? {})) rowSet.add(Number(k));
  const rows = [...rowSet].sort((a, b) => a - b);
  let minR = Infinity;
  let maxR = -1;
  let minC = Infinity;
  let maxC = -1;
  const out: string[] = [];
  for (const r of rows) {
    const rd = sheet.rowDefs?.[r];
    let attrs = ` r="${r + 1}"`;
    if (rd?.s !== undefined) attrs += ` s="${rd.s}" customFormat="1"`;
    if (rd?.ht !== undefined) attrs += ` ht="${rd.ht}"`;
    if (rd?.hidden) attrs += ' hidden="1"';
    if (rd?.custom) attrs += ' customHeight="1"';
    if (rd?.level) attrs += ` outlineLevel="${rd.level}"`;
    if (rd?.collapsed) attrs += ' collapsed="1"';
    if (rd?.extra) attrs += ` ${rd.extra}`;
    const cells: string[] = [];
    const row = sheet.cells[r];
    const brow = sheet.blankStyles?.[r];
    const n = Math.max(row?.length ?? 0, brow?.length ?? 0);
    for (let c = 0; c < n; c++) {
      const cell = row?.[c];
      const ref = `${colName(c)}${r + 1}`;
      if (cell) cells.push(cellXml(ref, cell, sst));
      else if (brow?.[c] !== undefined) cells.push(brow[c] ? `<c r="${ref}" s="${brow[c]}"/>` : `<c r="${ref}"/>`);
      else continue;
      minR = Math.min(minR, r);
      maxR = Math.max(maxR, r);
      minC = Math.min(minC, c);
      maxC = Math.max(maxC, c);
    }
    out.push(cells.length ? `<row${attrs}>${cells.join('')}</row>` : `<row${attrs}/>`);
  }
  const dim = maxR < 0 ? 'A1' : minR === maxR && minC === maxC ? `${colName(minC)}${minR + 1}` : `${colName(minC)}${minR + 1}:${colName(maxC)}${maxR + 1}`;
  return { xml: out.length ? `<sheetData>${out.join('')}</sheetData>` : '<sheetData/>', dim };
}

function colsXml(sheet: SheetData): string {
  const defs = sheet.colDefs ?? [];
  if (!defs.length) return '';
  const def = pxToChars(sheet.defaultColWidth);
  return `<cols>${defs
    .map((d) => {
      let a = ` min="${d.min + 1}" max="${d.max + 1}" width="${d.width ?? def}"`;
      if (d.style !== undefined) a += ` style="${d.style}"`;
      if (d.hidden) a += ' hidden="1"';
      if (d.width !== undefined) a += ' customWidth="1"';
      if (d.level) a += ` outlineLevel="${d.level}"`;
      if (d.collapsed) a += ' collapsed="1"';
      if (d.extra) a += ` ${d.extra}`;
      return `<col${a}/>`;
    })
    .join('')}</cols>`;
}

function mergesXml(sheet: SheetData): string {
  if (!sheet.merges.length) return '';
  const ref = (m: { r0: number; c0: number; r1: number; c1: number }) => `${colName(m.c0)}${m.r0 + 1}:${colName(m.c1)}${m.r1 + 1}`;
  return `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${ref(m)}"/>`).join('')}</mergeCells>`;
}

/** 冻结窗格写到第一个 sheetView 里 */
function patchFreeze(xml: string, sheet: SheetData): string {
  const { rows, cols } = sheet.freeze;
  const pane = xml.match(/<pane\b[^>]*>/)?.[0];
  const frozen = pane && /frozen/i.test(xmlAttr(pane, 'state') ?? '');
  const curR = frozen ? Math.floor(Number(xmlAttr(pane!, 'ySplit') ?? 0)) : 0;
  const curC = frozen ? Math.floor(Number(xmlAttr(pane!, 'xSplit') ?? 0)) : 0;
  if (curR === rows && curC === cols) return xml;
  const active = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
  const newPane =
    rows || cols
      ? `<pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${colName(cols)}${rows + 1}" activePane="${active}" state="frozen"/><selection pane="${active}" activeCell="${colName(cols)}${rows + 1}" sqref="${colName(cols)}${rows + 1}"/>`
      : '';
  const view = xml.match(/<sheetView\b[^>]*\/>|<sheetView\b[^>]*>[\s\S]*?<\/sheetView>/)?.[0];
  if (!view) {
    if (!newPane) return xml;
    return putElement(xml, 'sheetViews', `<sheetViews><sheetView workbookViewId="0">${newPane}</sheetView></sheetViews>`);
  }
  const open = view.match(/^<sheetView\b[^>]*?(?=\/?>)/)![0];
  const inner = view.endsWith('/>') ? '' : view.replace(/^<sheetView\b[^>]*>/, '').replace(/<\/sheetView>$/, '');
  // 去掉旧的 pane 和各窗格的 selection（只留 topLeft 窗格的）
  const rest = inner.replace(/<pane\b[^>]*\/>/g, '').replace(/<selection\b[^>]*\/>/g, (s) => (/pane="/.test(s) ? '' : s));
  return xml.replace(view, () => `${open}>${newPane}${rest}</sheetView>`);
}

/** 按插入 / 删除行列的记录，平移编辑器不建模的区域与公式 */
function replayShifts(xml: string, sheet: SheetData): string {
  let out = xml;
  for (const sh of sheet.shifts ?? []) {
    const spec = { sheet: sheet.name, axis: sh.axis, at: sh.at, n: sh.n, band: sh.band };
    const sq = (s: string) => shiftSqref(s, sh);
    const fm = (s: string) => esc(shiftFormula(unescapeXml(s), sheet.name, spec));
    // 条件格式（含 x14 扩展）、数据验证
    out = out.replace(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (all, attrs: string, inner: string) => {
      const ref = xmlAttr(attrs, 'sqref');
      if (ref === undefined) return all;
      const next = sq(ref);
      if (!next) return '';
      return `<conditionalFormatting${attrs.replace(/sqref="[^"]*"/, `sqref="${next}"`)}>${inner.replace(/<formula>([\s\S]*?)<\/formula>/g, (_m, f: string) => `<formula>${fm(f)}</formula>`)}</conditionalFormatting>`;
    });
    out = out.replace(/<dataValidation\b([^>]*?)(\/>|>([\s\S]*?)<\/dataValidation>)/g, (all, attrs: string, _end: string, inner?: string) => {
      const ref = xmlAttr(attrs, 'sqref');
      if (ref === undefined) return all;
      const next = sq(ref);
      if (!next) return '';
      const a = attrs.replace(/sqref="[^"]*"/, `sqref="${next}"`);
      if (inner === undefined) return `<dataValidation${a}/>`;
      return `<dataValidation${a}>${inner.replace(/<formula([12])>([\s\S]*?)<\/formula\1>/g, (_m, i: string, f: string) => `<formula${i}>${fm(f)}</formula${i}>`)}</dataValidation>`;
    });
    out = out.replace(/<dataValidations\b([^>]*)>([\s\S]*?)<\/dataValidations>/g, (_all, attrs: string, inner: string) => {
      const n = (inner.match(/<dataValidation\b/g) ?? []).length;
      return n ? `<dataValidations${attrs.replace(/count="\d+"/, `count="${n}"`)}>${inner}</dataValidations>` : '';
    });
    out = out.replace(/<xm:sqref>([\s\S]*?)<\/xm:sqref>/g, (_m, s: string) => `<xm:sqref>${sq(s)}</xm:sqref>`);
    out = out.replace(/<xm:f>([\s\S]*?)<\/xm:f>/g, (_m, f: string) => `<xm:f>${fm(f)}</xm:f>`);
    // 超链接、自动筛选、忽略错误、受保护区域
    out = out.replace(/<hyperlink\b[^>]*\/>/g, (tag) => {
      const ref = xmlAttr(tag, 'ref');
      if (ref === undefined) return tag;
      const next = sq(ref);
      return next ? tag.replace(/ref="[^"]*"/, `ref="${next}"`) : '';
    });
    out = out.replace(/<hyperlinks>\s*<\/hyperlinks>/, '');
    out = out.replace(/<autoFilter\b[^>]*?ref="([^"]*)"/, (all, ref: string) => {
      const next = sq(ref);
      return next ? all.replace(`ref="${ref}"`, `ref="${next}"`) : all;
    });
    out = out.replace(/<(ignoredError|protectedRange)\b[^>]*?sqref="([^"]*)"/g, (all, _t, ref: string) => all.replace(`sqref="${ref}"`, `sqref="${sq(ref) || ref}"`));
    // 分页符（插入 / 删除单元格时不动）
    const brk = sh.axis === 'row' ? 'rowBreaks' : 'colBreaks';
    if (!sh.band) out = out.replace(new RegExp(`<${brk}\\b[^>]*>[\\s\\S]*?</${brk}>`), (all) =>
      all.replace(/<brk\b([^>]*?)id="(\d+)"/g, (tag, pre: string, id: string) => {
        // brk id 是分页符下面那一行 / 右边那一列（从 1 开始的第 id 行之后）
        const i = Number(id);
        if (i < sh.at) return tag;
        return `<brk${pre}id="${Math.max(sh.at, i + sh.n)}"`;
      }),
    );
  }
  return out;
}

/** 自动筛选：按模型重写 <autoFilter>（没有时删除） */
function autoFilterXml(sheet: SheetData, styles?: StyleBook): string {
  const af = sheet.autoFilter;
  if (!af) return '';
  const ref = `${colName(af.ref.c0)}${af.ref.r0 + 1}:${colName(af.ref.c1)}${af.ref.r1 + 1}`;
  const cols = Object.entries(af.cols)
    .sort((a, b) => Number(a[0]) - Number(b[0]))
    .map(([k, f]) => {
      if (f.color) {
        // 按颜色：引用一个差异格式；无填充 / 自动色时没有 dxfId
        const dxf = f.color.dxf ?? (f.color.color && styles ? dxfForColor(styles, f.color.color, f.color.cell) : undefined);
        return `<filterColumn colId="${k}"><colorFilter${dxf !== undefined ? ` dxfId="${dxf}"` : ''}${f.color.cell ? '' : ' cellColor="0"'}/></filterColumn>`;
      }
      if (f.values || f.dates) {
        const values = f.values ?? [];
        const blank = values.includes('');
        const vals = values
          .filter((v) => v !== '')
          .map((v) => `<filter val="${esc(v)}"/>`)
          .join('');
        const dates = (f.dates ?? [])
          .map((g) => `<dateGroupItem year="${g.y}"${g.m !== undefined ? ` month="${g.m}"` : ''}${g.d !== undefined ? ` day="${g.d}"` : ''} dateTimeGrouping="${g.d !== undefined ? 'day' : g.m !== undefined ? 'month' : 'year'}"/>`)
          .join('');
        return `<filterColumn colId="${k}"><filters${blank ? ' blank="1"' : ''}>${vals}${dates}</filters></filterColumn>`;
      }
      if (f.custom) {
        const rules = f.custom.rules.map((r) => `<customFilter${r.op === 'equal' ? '' : ` operator="${r.op}"`} val="${esc(r.val)}"/>`).join('');
        return `<filterColumn colId="${k}"><customFilters${f.custom.and ? ' and="1"' : ''}>${rules}</customFilters></filterColumn>`;
      }
      return '';
    })
    .join('');
  return cols ? `<autoFilter ref="${ref}">${cols}</autoFilter>` : `<autoFilter ref="${ref}"/>`;
}

const PR_ORDER = ['tabColor', 'outlinePr', 'pageSetUpPr'];

/** 改 <sheetPr> 的一个子元素（content 为空串时删除），sheetPr 不存在时新建 */
function putSheetPr(xml: string, tag: string, content: string): string {
  const pr = xml.match(/<sheetPr\b[^>]*\/>|<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/)?.[0];
  if (!pr) return content ? putElement(xml, 'sheetPr', `<sheetPr>${content}</sheetPr>`) : xml;
  const open = pr.match(/^<sheetPr\b[^>]*?(?=\/?>)/)![0];
  const inner = pr.endsWith('/>') ? '' : pr.replace(/^<sheetPr\b[^>]*>/, '').replace(/<\/sheetPr>$/, '');
  const next = putElement(`<x>${inner}</x>`, tag, content, PR_ORDER, 'x').replace(/^<x>|<\/x>$/g, '');
  return xml.replace(pr, () => (next ? `${open}>${next}</sheetPr>` : `${open}/>`));
}

function setAttr(tag: string, name: string, value: string | undefined): string {
  const re = new RegExp(`\\s${name}="[^"]*"`);
  if (value === undefined) return tag.replace(re, '');
  return re.test(tag) ? tag.replace(re, () => ` ${name}="${value}"`) : tag.replace(/\s*(\/?>)$/, (_m, end: string) => ` ${name}="${value}"${end}`);
}

/** 页面设置：方向、纸张、缩放 / 调整为几页、网格线 */
function patchPage(xml: string, sheet: SheetData): string {
  const p = sheet.page;
  if (!p) return xml;
  let out = xml;
  const had = /<pageSetup\b/.test(out);
  let ps = out.match(/<pageSetup\b[^>]*>/)?.[0] ?? '<pageSetup/>';
  const needPs = had || !!p.landscape || !!p.paper || !!p.scale || p.fitW !== undefined || p.fitH !== undefined;
  ps = setAttr(ps, 'paperSize', p.paper ? String(p.paper) : undefined);
  ps = setAttr(ps, 'orientation', p.landscape ? 'landscape' : 'portrait');
  const fit = p.fitW !== undefined || p.fitH !== undefined;
  ps = setAttr(ps, 'scale', p.scale && !fit ? String(p.scale) : undefined);
  ps = setAttr(ps, 'fitToWidth', fit ? String(p.fitW ?? 1) : undefined);
  ps = setAttr(ps, 'fitToHeight', fit ? String(p.fitH ?? 1) : undefined);
  if (had) out = out.replace(/<pageSetup\b[^>]*>/, () => ps);
  else if (needPs) {
    if (!/<pageMargins\b/.test(out)) out = putElement(out, 'pageMargins', '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');
    out = putElement(out, 'pageSetup', ps);
  }
  if (p.margins) {
    const m = p.margins;
    out = putElement(out, 'pageMargins', `<pageMargins left="${m.left}" right="${m.right}" top="${m.top}" bottom="${m.bottom}" header="${m.header}" footer="${m.footer}"/>`);
  }
  // 页眉页脚：只改奇数页（默认）的，首页 / 偶数页等其它子元素原样保留
  const hf0 = out.match(/<headerFooter\b[^>]*\/>|<headerFooter\b[^>]*>[\s\S]*?<\/headerFooter>/)?.[0];
  if (hf0 || p.header || p.footer) {
    const open = hf0?.match(/^<headerFooter\b[^>]*?(?=\/?>)/)?.[0] ?? '<headerFooter';
    let inner = hf0 && !hf0.endsWith('/>') ? hf0.replace(/^<headerFooter\b[^>]*>/, '').replace(/<\/headerFooter>$/, '') : '';
    inner = inner.replace(/<oddHeader>[\s\S]*?<\/oddHeader>|<oddHeader\/>/, '').replace(/<oddFooter>[\s\S]*?<\/oddFooter>|<oddFooter\/>/, '');
    const head = (p.header ? `<oddHeader>${esc(p.header)}</oddHeader>` : '') + (p.footer ? `<oddFooter>${esc(p.footer)}</oddFooter>` : '');
    const next = head + inner;
    out = putElement(out, 'headerFooter', next ? `${open}>${next}</headerFooter>` : '');
  }
  const brks = (tag: 'rowBreaks' | 'colBreaks', list: number[] | undefined, max: number) =>
    (out = putElement(out, tag, list?.length ? `<${tag} count="${list.length}" manualBreakCount="${list.length}">${list.map((id) => `<brk id="${id}" max="${max}" man="1"/>`).join('')}</${tag}>` : ''));
  brks('rowBreaks', p.rowBreaks, 16383);
  brks('colBreaks', p.colBreaks, 1048575);
  out = putSheetPr(out, 'pageSetUpPr', fit ? '<pageSetUpPr fitToPage="1"/>' : '');
  // 打印选项：网格线、水平 / 垂直居中
  const po = out.match(/<printOptions\b[^>]*\/>/)?.[0];
  let next = po ?? '<printOptions/>';
  next = setAttr(next, 'gridLines', p.gridLines ? '1' : undefined);
  next = setAttr(next, 'horizontalCentered', p.centerH ? '1' : undefined);
  next = setAttr(next, 'verticalCentered', p.centerV ? '1' : undefined);
  const empty = /^<printOptions\s*\/>$/.test(next);
  if (po) out = out.replace(po, () => (empty ? '' : next));
  else if (!empty) out = putElement(out, 'printOptions', next);
  return out;
}

/** 条件格式的改动：先按原 XML 里的序号去掉删掉的规则（此时序号还没变），平移之后再追加新建的规则 */
function removeCfRules(xml: string, removed: number[] | undefined): string {
  if (!removed?.length) return xml;
  const gone = new Set(removed);
  let n = 0;
  return xml
    .replace(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (_all, attrs: string, inner: string) => {
      let kept = 0;
      const rest = inner.replace(/<cfRule\b[^>]*?(?:\/>|>[\s\S]*?<\/cfRule>)/g, (r) => {
        if (gone.has(n++)) return '';
        kept++;
        return r;
      });
      return kept ? `<conditionalFormatting${attrs}>${rest}</conditionalFormatting>` : '';
    });
}

function addCfRules(xml: string, sheet: SheetData): string {
  const fresh = (sheet.cf ?? []).filter((r) => r.idx === undefined);
  if (!fresh.length) return xml;
  const blocks = fresh.map((r) => `<conditionalFormatting sqref="${r.ranges.map(rangeRef).join(' ')}">${cfRuleXml(r)}</conditionalFormatting>`).join('');
  const last = [...xml.matchAll(/<\/conditionalFormatting>/g)].pop();
  if (last?.index !== undefined) {
    const at = last.index + last[0].length;
    return xml.slice(0, at) + blocks + xml.slice(at);
  }
  return putElement(xml, 'conditionalFormatting', blocks);
}

/** 表格部件里的 <autoFilter>：条件换成当前的，区域沿用文件里原来的文字；筛选被取消时去掉这个元素 */
function patchTableFilter(xml: string, sheet: SheetData, origRef: string): string {
  const re = /<autoFilter\b[^>]*?(?:\/>|>[\s\S]*?<\/autoFilter>)/;
  const af = sheet.autoFilter?.table ? autoFilterXml(sheet) : '';
  const next = af && origRef ? af.replace(/ref="[^"]*"/, `ref="${origRef}"`) : af;
  if (re.test(xml)) return xml.replace(re, () => next);
  if (!next) return xml;
  return xml.replace(/<(sortState|tableColumns)\b/, (m) => `${next}${m}`);
}

/** 有分组时 <sheetFormatPr> 要记下最大级别（outlineLevelRow / outlineLevelCol），Excel 才会画分组栏 */
function patchOutlineLevels(xml: string, sheet: SheetData): string {
  const rowMax = Math.max(0, ...Object.values(sheet.rowDefs ?? {}).map((d) => d.level ?? 0));
  const colMax = Math.max(0, ...(sheet.colDefs ?? []).map((d) => d.level ?? 0));
  const tag = xml.match(/<sheetFormatPr\b[^>]*\/?>/)?.[0];
  if (!tag) {
    if (!rowMax && !colMax) return xml;
    const a = `${rowMax ? ` outlineLevelRow="${rowMax}"` : ''}${colMax ? ` outlineLevelCol="${colMax}"` : ''}`;
    return putElement(xml, 'sheetFormatPr', `<sheetFormatPr defaultRowHeight="14.25"${a}/>`);
  }
  let next = setAttr(tag, 'outlineLevelRow', rowMax ? String(rowMax) : undefined);
  next = setAttr(next, 'outlineLevelCol', colMax ? String(colMax) : undefined);
  return next === tag ? xml : xml.replace(tag, () => next);
}

/** <sheetProtection> / <workbookProtection>：属性按模型里的顺序写，没有 = 去掉元素 */
function protectionXml(tag: string, attrs: Record<string, string> | undefined): string {
  if (!attrs) return '';
  return `<${tag}${Object.entries(attrs).map(([k, v]) => ` ${k}="${esc(v)}"`).join('')}/>`;
}

function writeSheet(xml: string, sheet: SheetData, sst: SharedStrings, styles?: StyleBook): string {
  const data = sheetDataXml(sheet, sst);
  let out = replayShifts(removeCfRules(xml, sheet.cfRemoved), sheet);
  out = addCfRules(out, sheet);
  out = putElement(out, 'sheetData', data.xml);
  out = putElement(out, 'dimension', `<dimension ref="${data.dim}"/>`);
  out = putElement(out, 'cols', colsXml(sheet));
  out = patchOutlineLevels(out, sheet);
  out = putElement(out, 'mergeCells', mergesXml(sheet));
  out = putElement(out, 'autoFilter', autoFilterXml(sheet, styles));
  out = putSheetPr(out, 'tabColor', sheet.tabColor ? `<tabColor rgb="FF${sheet.tabColor.slice(1).toUpperCase()}"/>` : '');
  out = patchFreeze(out, sheet);
  out = patchPage(out, sheet);
  out = putElement(out, 'sheetProtection', protectionXml('sheetProtection', sheet.protection?.attrs));
  return out;
}

/**
 * 插入 / 删除行列后，工作表关联的部件里的位置也要跟着移：批注（ref）、批注的 VML 形状（x:Row / x:Column）、
 * 图片和图表的锚点（xdr:row / xdr:col）。被删掉的行列里的批注去掉，锚点贴到删除处。
 */
function shiftParts(files: Record<string, Uint8Array>, sheetPath: string, shifts: NonNullable<SheetData['shifts']>) {
  const relsPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels');
  if (!files[relsPath]) return;
  const at1 = (i: number, sh: (typeof shifts)[number]) => (sh.n < 0 && i >= sh.at && i < sh.at - sh.n ? -1 : i >= sh.at ? i + sh.n : i);
  for (const m of strFromU8(files[relsPath]).matchAll(/<Relationship\b[^>]*>/g)) {
    const type = xmlAttr(m[0], 'Type') ?? '';
    const target = xmlAttr(m[0], 'Target');
    if (!target || xmlAttr(m[0], 'TargetMode') === 'External') continue;
    const part = resolvePart(sheetPath, target);
    if (!files[part]) continue;
    // 只处理批注、批注框、图片 / 图表锚点；其它关联部件（打印机设置 .bin、表格、超链接等）不能当文字解码再写回：
    // 二进制内容会被 UTF-8 替换字符破坏，Excel 打不开
    if (!/\/(comments|vmlDrawing|drawing)$/.test(type)) continue;
    let x = strFromU8(files[part]);
    for (const sh of shifts) {
      if (/\/comments$/.test(type)) {
        x = x.replace(/<comment\b[^>]*>[\s\S]*?<\/comment>/g, (c) => {
          const ref = xmlAttr(c, 'ref');
          const g = ref ? parseRef(ref) : null;
          if (!g) return c;
          if (!withinBand(sh.axis === 'row' ? g.c0 : g.r0, sh.axis === 'row' ? g.c0 : g.r0, sh.band)) return c;
          const i = at1(sh.axis === 'row' ? g.r0 : g.c0, sh);
          if (i < 0) return '';
          const next = sh.axis === 'row' ? `${colName(g.c0)}${i + 1}` : `${colName(i)}${g.r0 + 1}`;
          return c.replace(/ref="[^"]*"/, `ref="${next}"`);
        });
      } else if (/\/vmlDrawing$/.test(type) && sh.band) {
        // 插入 / 删除单元格：只移动 band 里的批注框（行列号成对出现在 x:ClientData 里）
        x = x.replace(/<x:Row>(\d+)<\/x:Row>(\s*)<x:Column>(\d+)<\/x:Column>/g, (all, r: string, sp: string, c: string) => {
          const [rr, cc] = [Number(r), Number(c)];
          if (!withinBand(sh.axis === 'row' ? cc : rr, sh.axis === 'row' ? cc : rr, sh.band)) return all;
          const i = Math.max(0, at1(sh.axis === 'row' ? rr : cc, sh));
          return sh.axis === 'row' ? `<x:Row>${i}</x:Row>${sp}<x:Column>${c}</x:Column>` : `<x:Row>${r}</x:Row>${sp}<x:Column>${i}</x:Column>`;
        });
      } else if (/\/drawing$/.test(type) && sh.band) {
        // 图片、图表不随插入 / 删除单元格移动（同 Excel）
        continue;
      } else if (/\/vmlDrawing$/.test(type)) {
        const tag = sh.axis === 'row' ? 'x:Row' : 'x:Column';
        x = x.replace(new RegExp(`<${tag}>(\\d+)</${tag}>`, 'g'), (_m, v: string) => `<${tag}>${Math.max(0, at1(Number(v), sh))}</${tag}>`);
      } else if (/\/drawing$/.test(type)) {
        const tag = sh.axis === 'row' ? 'xdr:row' : 'xdr:col';
        x = x.replace(new RegExp(`<${tag}>(\\d+)</${tag}>`, 'g'), (_m, v: string) => {
          const i = at1(Number(v), sh);
          return `<${tag}>${i < 0 ? sh.at : i}</${tag}>`;
        });
      } else continue;
    }
    files[part] = strToU8(x);
  }
}

/** 复制工作表时以原表 XML 为底：去掉指向其它部件的关系（图表、图片、批注、表格、带 r:id 的超链接） */
function stripRels(xml: string): string {
  return xml
    .replace(/<(drawing|legacyDrawing|legacyDrawingHF|picture)\b[^>]*\/>/g, '')
    .replace(/<(tableParts|oleObjects|controls)\b[\s\S]*?<\/\1>/g, '')
    .replace(/<hyperlink\b[^>]*r:id="[^"]*"[^>]*\/>/g, '')
    .replace(/<hyperlinks>\s*<\/hyperlinks>/, '')
    .replace(/(<pageSetup\b[^>]*?)\s+r:id="[^"]*"/, '$1')
    .replace(/\stabSelected="1"/, '');
}

const BLANK_SHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="14.25"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;

// ---------- 工作簿 ----------

function removeCalcChain(files: Record<string, Uint8Array>, relsPath: string) {
  const chain = Object.keys(files).find((p) => /^xl\/calcChain\.xml$/i.test(p));
  if (!chain) return;
  delete files[chain];
  const rels = files[relsPath] ? strFromU8(files[relsPath]) : '';
  files[relsPath] = strToU8(rels.replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g, ''));
  const ct = strFromU8(files['[Content_Types].xml']);
  files['[Content_Types].xml'] = strToU8(ct.replace(/<Override\b[^>]*calcChain\.xml[^>]*\/>/g, ''));
}

const WB_ORDER = [
  'fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames',
  'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst',
];

/** 工作簿里有公式时让 Excel 打开时全部重算（编辑器算出的缓存值可能和 Excel 不一致，或者还没算） */
function setFullCalc(wb: string): string {
  const pr = wb.match(/<calcPr\b[^>]*\/?>/)?.[0];
  if (pr) {
    if (/fullCalcOnLoad=/.test(pr)) return wb.replace(pr, () => pr.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"'));
    return wb.replace(pr, () => pr.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"'));
  }
  return putElement(wb, 'calcPr', '<calcPr fullCalcOnLoad="1"/>', WB_ORDER, 'workbook');
}

/** 每张表保存后的位置信息（保存成功后由 markSaved 写回模型） */
interface SheetPlan {
  path: string;
  rid: string;
  sheetId: number;
}
const plans = new WeakMap<Workbook, SheetPlan[]>();
const pivotPlans = new WeakMap<Workbook, Map<number, PivotParts>>();

const absRange = (g: { r0: number; c0: number; r1: number; c1: number }) => `$${colName(g.c0)}$${g.r0 + 1}:$${colName(g.c1)}$${g.r1 + 1}`;

/**
 * 定义名称：按表的新位置改 localSheetId（删掉的表的名称去掉）；表改名、插删行列后改写引用；
 * 用户定义的名称按模型写；打印标题 / 打印区域 / 筛选区域按模型重新生成。
 */
function syncDefinedNames(wb: string, book: Workbook): string {
  const byOrig = new Map<number, number>();
  book.sheets.forEach((s, i) => s.origIndex !== undefined && byOrig.set(s.origIndex, i));
  const renamed = book.sheets.filter((s) => s.fileName && s.fileName !== s.name);
  const shifted = book.sheets.filter((s) => s.shifts?.length);
  const alive = new Set(book.sheets.flatMap((s) => [s.name.toLowerCase(), (s.fileName ?? s.name).toLowerCase()]));
  const kept: string[] = [];
  const userDone = new Set<string>();
  for (const m of wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    let attrs = m[1];
    const name = unescapeXml(xmlAttr(attrs, 'name') ?? '');
    const local = xmlAttr(attrs, 'localSheetId');
    let idx: number | undefined;
    if (local !== undefined) {
      idx = byOrig.get(Number(local));
      if (idx === undefined) continue;
      attrs = attrs.replace(/localSheetId="\d+"/, `localSheetId="${idx}"`);
    }
    if (/^_xlnm\.(Print_Titles|Print_Area|_FilterDatabase)$/.test(name)) continue;
    let f = unescapeXml(m[2]);
    if (!name.startsWith('_xlnm.')) {
      const user = (book.names ?? []).find((n) => n.name.toLowerCase() === name.toLowerCase() && n.local === idx);
      if (!user) continue;
      userDone.add(`${name.toLowerCase()}|${idx ?? ''}`);
      f = toFileFormula(user.f);
    } else {
      // 其它内置名称：表改名 / 删除后跟着改
      for (const s of renamed) f = renameSheetRefs(f, s.fileName!, s.name);
      f = mapRefs(f, (ref) => (ref.sheet !== undefined && !alive.has(ref.sheet.toLowerCase()) ? null : undefined));
    }
    for (const s of shifted) for (const sh of s.shifts!) f = shiftFormula(f, undefined, { sheet: s.name, axis: sh.axis, at: sh.at, n: sh.n, band: sh.band });
    kept.push(`<definedName${attrs}>${esc(f)}</definedName>`);
  }
  for (const n of book.names ?? []) {
    if (userDone.has(`${n.name.toLowerCase()}|${n.local ?? ''}`)) continue;
    kept.push(`<definedName name="${esc(n.name)}"${n.local !== undefined ? ` localSheetId="${n.local}"` : ''}>${esc(toFileFormula(n.f))}</definedName>`);
  }
  book.sheets.forEach((s, i) => {
    const q = quoteSheet(s.name);
    if (s.autoFilter && !s.tableFilter) kept.push(`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${esc(`${q}!${absRange(s.autoFilter.ref)}`)}</definedName>`);
    if (s.page?.area) kept.push(`<definedName name="_xlnm.Print_Area" localSheetId="${i}">${esc(`${q}!${absRange(s.page.area)}`)}</definedName>`);
    if (s.page?.titleRows) kept.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${esc(`${q}!$${s.page.titleRows[0] + 1}:$${s.page.titleRows[1] + 1}`)}</definedName>`);
  });
  return putElement(wb, 'definedNames', kept.length ? `<definedNames>${kept.join('')}</definedNames>` : '', WB_ORDER, 'workbook');
}

/**
 * 编辑后的工作簿 → xlsx 字节。book.source 为原文件（没有时按新建处理）；
 * 只有改过的工作表会重写，其余部件逐字节保留。不修改模型（保存成功后调用 markSaved）。
 */
export function saveXlsx(book: Workbook): Uint8Array {
  const fresh = !book.source || book.kind !== 'xlsx';
  const files = fresh ? blankPackage(book.sheets.map((s) => s.name)) : unzipSync(book.source!);
  const text = (p: string) => (files[p] ? strFromU8(files[p]) : '');
  const relsPath = 'xl/_rels/workbook.xml.rels';
  let rels = text(relsPath);
  let ct = text('[Content_Types].xml');
  let wb = text('xl/workbook.xml');
  const origWb = wb;
  const relTargets = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({
    id: xmlAttr(m[0], 'Id') ?? '',
    type: xmlAttr(m[0], 'Type') ?? '',
    path: resolvePart('xl/workbook.xml', xmlAttr(m[0], 'Target') ?? ''),
  }));
  const fileSheets = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => ({ tag: m[0], rid: xmlAttr(m[0], 'r:id') ?? '', id: Number(xmlAttr(m[0], 'sheetId')) }));

  // 每张表的部件：新建包按顺序；原文件里已有的沿用；新加 / 复制的表分配新部件
  const usedIds = new Set(fileSheets.map((s) => s.id));
  let nextRel = Math.max(0, ...relTargets.map((r) => Number(r.id.replace(/\D/g, '')) || 0)) + 1;
  let nextPart = 1;
  const existing = (s: SheetData) => !fresh && !!s.path && !!s.rid && !!files[s.path];
  const plan: SheetPlan[] = book.sheets.map((s, i) => {
    if (fresh) return { path: `xl/worksheets/sheet${i + 1}.xml`, rid: `rId${i + 1}`, sheetId: i + 1 };
    if (existing(s)) return { path: s.path!, rid: s.rid!, sheetId: s.sheetId ?? 0 };
    while (files[`xl/worksheets/sheet${nextPart}.xml`]) nextPart++;
    const path = `xl/worksheets/sheet${nextPart++}.xml`;
    const rid = `rId${nextRel++}`;
    let sheetId = 1;
    while (usedIds.has(sheetId)) sheetId++;
    usedIds.add(sheetId);
    files[path] = strToU8(s.copyOf && files[s.copyOf] ? stripRels(text(s.copyOf)) : BLANK_SHEET);
    rels = rels.replace(/<\/Relationships>\s*$/, `<Relationship Id="${rid}" Type="${NS_REL}/worksheet" Target="${path.replace(/^xl\//, '')}"/></Relationships>`);
    ct = ct.replace(/<\/Types>\s*$/, `<Override PartName="/${path}" ContentType="${CT_SHEET}"/></Types>`);
    return { path, rid, sheetId };
  });
  const added = book.sheets.map((s) => !fresh && !existing(s));

  // 删除的表：去掉部件、它的关系文件、工作簿关系和内容类型
  if (!fresh) {
    const keep = new Set(plan.map((p) => p.rid));
    for (const fs of fileSheets) {
      if (keep.has(fs.rid)) continue;
      const target = relTargets.find((r) => r.id === fs.rid);
      if (!target) continue;
      delete files[target.path];
      delete files[target.path.replace(/([^/]+)$/, '_rels/$1.rels')];
      rels = rels.replace(new RegExp(`<Relationship\\b[^>]*Id="${fs.rid}"[^>]*/>`), '');
      ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${target.path.replace(/\./g, '\\.')}"[^>]*/>`), '');
    }
  }

  // 透视表缓存：工作簿里已有的缓存编号不能重复
  const usedCacheIds = new Set([...wb.matchAll(/<pivotCache\b[^>]*\bcacheId="(\d+)"/g)].map((m) => Number(m[1])));
  const pivotParts = new Map<number, PivotParts>();
  const newCaches: PivotParts[] = [];
  let sstPath = relTargets.find((r) => r.type.endsWith('/sharedStrings'))?.path;
  const sst = new SharedStrings(sstPath ? text(sstPath) : '');
  book.sheets.forEach((s, i) => {
    if (!(fresh || s.dirty || added[i])) return;
    // 表格自带的筛选写在表格部件里，工作表自己不能再写 autoFilter（两个重叠会让 Excel 报错）
    const tf = !fresh && !added[i] && s.tableFilter && files[s.tableFilter.path] ? s.tableFilter : undefined;
    files[plan[i].path] = strToU8(writeSheet(text(plan[i].path), { ...s, ...(tf && { autoFilter: undefined }), shifts: fresh || added[i] ? [] : s.shifts }, sst, book.styleBook));
    if (tf) files[tf.path] = strToU8(patchTableFilter(text(tf.path), s, tf.ref));
    if (!fresh && !added[i] && s.shifts?.length) shiftParts(files, plan[i].path, s.shifts);
    // 批注：和文件里一致就不动；有增删改才重写 comments / VML
    const pk: Pkg = { files, ct };
    syncNotes(pk, plan[i].path, s, (x, tag, content) => putElement(x, tag, content));
    syncDrawings(pk, plan[i].path, s, book, (x, tag, content) => putElement(x, tag, content));
    // 数据透视表：写成 Excel 真部件（缓存 + 透视表）；新缓存稍后在 workbook.xml 里登记
    const pv = syncPivot(pk, plan[i].path, s, book, usedCacheIds);
    if (pv) {
      pivotParts.set(i, pv.parts);
      if (pv.isNew) newCaches.push(pv.parts);
    }
    ct = pk.ct;
  });

  if (sst.changed) {
    if (!sstPath) {
      sstPath = 'xl/sharedStrings.xml';
      rels = rels.replace(/<\/Relationships>\s*$/, `<Relationship Id="rId${nextRel++}" Type="${REL_SST}" Target="sharedStrings.xml"/></Relationships>`);
      ct = ct.replace(/<\/Types>\s*$/, `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT_SST}"/></Types>`);
    }
    files[sstPath] = strToU8(sst.toXml());
  }

  if (book.styleBook?.dirty || fresh) {
    const stylesPath = relTargets.find((r) => r.type.endsWith('/styles'))?.path ?? 'xl/styles.xml';
    if (book.styleBook) files[stylesPath] = strToU8(writeStyles(text(stylesPath), book.styleBook));
  }

  // 工作表列表（名称、顺序、隐藏）：和文件里一样时不动
  const sameList =
    fileSheets.length === book.sheets.length &&
    book.sheets.every((s, i) => fileSheets[i].rid === plan[i].rid && unescapeXml(xmlAttr(fileSheets[i].tag, 'name') ?? '') === s.name && (xmlAttr(fileSheets[i].tag, 'state') ?? undefined) === s.hidden);
  if (!sameList) {
    const sheetsXml = `<sheets>${book.sheets
      .map((s, i) => {
        let tag = fileSheets.find((f) => f.rid === plan[i].rid)?.tag ?? `<sheet name="" sheetId="${plan[i].sheetId}" r:id="${plan[i].rid}"/>`;
        tag = setAttr(tag, 'name', esc(s.name));
        tag = setAttr(tag, 'state', s.hidden);
        return tag.endsWith('/>') ? tag : tag.replace(/\s*>$/, '/>');
      })
      .join('')}</sheets>`;
    wb = wb.replace(/<sheets\b[^>]*>[\s\S]*?<\/sheets>|<sheets\s*\/>/, () => sheetsXml);
    // 当前标签不能超出范围、不能是隐藏的表
    const firstVisible = Math.max(0, book.sheets.findIndex((s) => !s.hidden));
    wb = wb.replace(/<workbookView\b[^>]*>/, (tag) => {
      const active = Number(xmlAttr(tag, 'activeTab') ?? 0);
      const ok = active < book.sheets.length && !book.sheets[active]?.hidden;
      const t = setAttr(tag, 'activeTab', ok ? (active ? String(active) : undefined) : firstVisible ? String(firstVisible) : undefined);
      return setAttr(t, 'firstSheet', undefined);
    });
  }
  const anyDirty = fresh || !sameList || !!book.namesDirty || book.sheets.some((s, i) => s.dirty || added[i]);
  if (anyDirty) {
    wb = syncDefinedNames(wb, book);
    wb = putElement(wb, 'workbookProtection', protectionXml('workbookProtection', book.wbProtection?.attrs), WB_ORDER, 'workbook');
    const hasFormula = book.sheets.some((s) => s.cells.some((row) => row?.some((c) => c?.f || c?.fraw)));
    if (hasFormula) wb = setFullCalc(wb);
  }
  if (newCaches.length) {
    const entries = [...wb.matchAll(/<pivotCache\b[^>]*\/>/g)].map((m) => m[0]);
    for (const c of newCaches) {
      const rid = `rId${nextRel++}`;
      rels = rels.replace(/<\/Relationships>\s*$/, `<Relationship Id="${rid}" Type="${REL.pivotCacheDef}" Target="${relTarget('xl/workbook.xml', c.cacheDef)}"/></Relationships>`);
      entries.push(`<pivotCache cacheId="${c.cacheId}" r:id="${rid}"/>`);
    }
    wb = putElement(wb, 'pivotCaches', `<pivotCaches>${entries.join('')}</pivotCaches>`, WB_ORDER, 'workbook');
  }
  if (wb !== origWb) files['xl/workbook.xml'] = strToU8(wb);
  files[relsPath] = strToU8(rels);
  files['[Content_Types].xml'] = strToU8(ct);
  if (anyDirty) removeCalcChain(files, relsPath);
  plans.set(book, plan);
  pivotPlans.set(book, pivotParts);
  // Old versions could UTF-8-encode printer-driver binaries while shifting
  // rows. Retain healthy binaries, and detach only positively identified damage.
  detachDamagedPrinterSettings(files);
  return zipSync(files, { level: 6 });
}

/** 保存成功后：当前内容成为新的“原文件”，记下各表的部件位置，清掉改动标记与平移记录 */
export function markSaved(book: Workbook, bytes: Uint8Array): void {
  const plan = plans.get(book);
  const pivots = pivotPlans.get(book);
  book.source = bytes;
  book.kind = 'xlsx';
  delete book.namesDirty;
  if (book.styleBook) {
    book.styleBook.dirty = false;
    delete book.styleBook.newDxfs;
  }
  let saved: Record<string, Uint8Array> | undefined;
  book.sheets.forEach((s, i) => {
    // 图表 / 图片：改过的工作表按新文件重新读一遍（部件路径、第几个锚点都变了）
    if (s.dirty && plan?.[i] && (s.drawings?.length || s.hadDrawings)) {
      saved ??= unzipSync(bytes);
      const fresh = readDrawings(saved, plan[i].path).drawings;
      if (fresh.length) s.drawings = fresh;
      else delete s.drawings;
      s.hadDrawings = fresh.length > 0;
    }
    const pp = pivots?.get(i);
    if (pp && s.pivot) s.pivot = { ...s.pivot, parts: pp };
    if (plan?.[i]) {
      s.path = plan[i].path;
      s.rid = plan[i].rid;
      s.sheetId = plan[i].sheetId;
    }
    s.origIndex = i;
    s.fileName = s.name;
    delete s.copyOf;
    s.dirty = false;
    s.shifts = [];
    // 条件格式：删掉的已经从文件里去掉，新建的已经写入，序号按文件里的新顺序重排
    delete s.cfRemoved;
    if (s.cf) {
      const old = s.cf.filter((r) => r.idx !== undefined).sort((a, b) => a.idx! - b.idx!);
      [...old, ...s.cf.filter((r) => r.idx === undefined)].forEach((r, k) => (r.idx = k));
    }
  });
}

// ---------- CSV ----------

/** 一张表 → CSV 文字（显示的文字，同 Excel 另存为 CSV），行尾 CRLF */
export function saveCsv(sheet: SheetData, delim = ','): string {
  const lines: string[] = [];
  for (let r = 0; r < sheet.rows; r++) {
    const row = sheet.cells[r];
    let last = -1;
    if (row) for (let c = row.length - 1; c >= 0; c--) if (row[c]) {
      last = c;
      break;
    }
    const fields: string[] = [];
    for (let c = 0; c <= last; c++) {
      const t = row?.[c]?.t ?? '';
      fields.push(t.includes(delim) || /["\r\n]/.test(t) || /^\s|\s$/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
    }
    lines.push(fields.join(delim));
  }
  return lines.join('\r\n') + (lines.length ? '\r\n' : '');
}

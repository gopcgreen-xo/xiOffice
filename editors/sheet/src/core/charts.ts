/**
 * 图表：新建、修改（类型 / 标题 / 数据）、移动缩放、删除，并写进 xlsx 的 drawing / chart 部件。
 *
 * 写出时尽量少动：和文件里一致的图表一个字节不改；只改了位置大小 → 只换锚点；只改标题或数据且类型没变 → 在原 chart XML 上
 * 改标题和系列引用（保留原来的样式）；改了类型或增删系列 → 按新类型整个重新生成该图表（样式回到默认）。
 * 新图表按 Excel 的标准结构生成。图表引用的区域随插入 / 删除行列自动调整（见 ops.ts）。
 */
import { strFromU8, strToU8 } from 'fflate';
import { colName } from '@wz/editor-shared/xlsx';
import { readDrawings, type Anchor, type ChartSeries, type ChartSpec, type SheetDrawing } from './drawing';
import { listRefs, quoteSheet } from './formula/refs';
import type { Tx } from './ops';
import { addOverride, addRel, CT, dropPart, esc, nextPart, readRels, REL, removeRel, type Pkg } from './parts';
import type { Range, SheetData, Workbook } from './workbook';

const EMU_PX = 9525;
export const CHART_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E'];

export const CHART_TYPES: { type: ChartSpec['type']; label: string }[] = [
  { type: 'col', label: '柱形图' },
  { type: 'bar', label: '条形图' },
  { type: 'line', label: '折线图' },
  { type: 'area', label: '面积图' },
  { type: 'pie', label: '饼图' },
];

// ---------- 取值 ----------

/** 引用（Sheet1!$A$2:$A$9）里的格子：显示文字与数值 */
export function refCells(book: Workbook, ref: string | undefined): { t: string; v: number | null }[] | null {
  if (!ref) return null;
  const [r] = listRefs(ref);
  if (!r) return null;
  const sheet = r.sheet === undefined ? book.sheets[0] : book.sheets.find((s) => s.name.toLowerCase() === r.sheet!.toLowerCase());
  if (!sheet) return null;
  const a = r.a;
  const b = r.b ?? a;
  const out: { t: string; v: number | null }[] = [];
  for (let row = Math.min(a.r ?? 0, b.r ?? 0); row <= Math.max(a.r ?? 0, b.r ?? 0); row++) {
    for (let c = Math.min(a.c ?? 0, b.c ?? 0); c <= Math.max(a.c ?? 0, b.c ?? 0); c++) {
      const cell = sheet.cells[row]?.[c];
      out.push({ t: cell?.t ?? '', v: typeof cell?.v === 'number' ? cell.v : null });
    }
  }
  return out;
}

const absCell = (r: number, c: number) => `$${colName(c)}$${r + 1}`;
const refOf = (sheet: SheetData, r0: number, c0: number, r1: number, c1: number) => `${quoteSheet(sheet.name)}!${absCell(r0, c0)}${r0 === r1 && c0 === c1 ? '' : `:${absCell(r1, c1)}`}`;

export interface ChartSource {
  /** 第一行是系列名称（byRows 时是第一列） */
  header: boolean;
  /** 第一列是分类标签（byRows 时是第一行） */
  labels: boolean;
  /** 按行取系列（默认按列：每一列一个系列） */
  byRows: boolean;
}

/** 猜这片数据怎么画：第一行都是文字且下面有数字 → 有标题行；第一列是文字 → 分类标签 */
export function guessSource(sheet: SheetData, g: Range): ChartSource {
  const isNum = (r: number, c: number) => typeof sheet.cells[r]?.[c]?.v === 'number';
  const isText = (r: number, c: number) => typeof sheet.cells[r]?.[c]?.v === 'string' && sheet.cells[r][c]!.v !== '';
  let headerText = g.r1 > g.r0;
  for (let c = g.c0 + 1; c <= g.c1; c++) if (!isText(g.r0, c)) headerText = false;
  if (g.c1 === g.c0) headerText = g.r1 > g.r0 && isText(g.r0, g.c0) && isNum(g.r0 + 1, g.c0);
  const r = g.r0 + (headerText ? 1 : 0);
  const labels = g.c1 > g.c0 && isText(r, g.c0);
  return { header: headerText, labels, byRows: false };
}

/** 数据区域 → 系列（引用写成 Sheet1!$B$2:$B$5，文字 / 数值缓存取当前格子） */
export function seriesFromRange(book: Workbook, si: number, g: Range, src: ChartSource): ChartSeries[] {
  const sheet = book.sheets[si];
  const out: ChartSeries[] = [];
  const cells = (ref: string | undefined) => refCells(book, ref) ?? [];
  if (!src.byRows) {
    const c0 = g.c0 + (src.labels ? 1 : 0);
    const r0 = g.r0 + (src.header ? 1 : 0);
    if (r0 > g.r1 || c0 > g.c1) return [];
    const catRef = src.labels ? refOf(sheet, r0, g.c0, g.r1, g.c0) : undefined;
    for (let c = c0; c <= g.c1; c++) {
      const nameRef = src.header ? refOf(sheet, g.r0, c, g.r0, c) : undefined;
      const valRef = refOf(sheet, r0, c, g.r1, c);
      out.push({ name: nameRef ? (cells(nameRef)[0]?.t ?? '') : `系列${out.length + 1}`, ...(nameRef && { nameRef }), ...(catRef && { catRef }), valRef, cats: cells(catRef).map((x) => x.t), vals: cells(valRef).map((x) => x.v ?? 0) });
    }
  } else {
    const r0 = g.r0 + (src.labels ? 1 : 0);
    const c0 = g.c0 + (src.header ? 1 : 0);
    if (r0 > g.r1 || c0 > g.c1) return [];
    const catRef = src.labels ? refOf(sheet, g.r0, c0, g.r0, g.c1) : undefined;
    for (let r = r0; r <= g.r1; r++) {
      const nameRef = src.header ? refOf(sheet, r, g.c0, r, g.c0) : undefined;
      const valRef = refOf(sheet, r, c0, r, g.c1);
      out.push({ name: nameRef ? (cells(nameRef)[0]?.t ?? '') : `系列${out.length + 1}`, ...(nameRef && { nameRef }), ...(catRef && { catRef }), valRef, cats: cells(catRef).map((x) => x.t), vals: cells(valRef).map((x) => x.v ?? 0) });
    }
  }
  return out;
}

/** 按当前格子刷新系列的名称和缓存（图表引用的格子变了） */
export function refreshSeries(book: Workbook, spec: ChartSpec): ChartSpec {
  return {
    ...spec,
    series: spec.series.map((s) => {
      const vals = refCells(book, s.valRef);
      const cats = refCells(book, s.catRef);
      const name = refCells(book, s.nameRef);
      return { ...s, ...(name && { name: name[0]?.t ?? s.name }), ...(vals && { vals: vals.map((x) => x.v ?? 0) }), ...(cats && { cats: cats.map((x) => x.t) }) };
    }),
  };
}

// ---------- 生成 chart XML ----------

const C_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const strRef = (f: string | undefined, vals: string[]) =>
  f
    ? `<c:strRef><c:f>${esc(f)}</c:f><c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`
    : `<c:strLit><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strLit>`;
const numRef = (f: string | undefined, vals: number[]) =>
  f
    ? `<c:numRef><c:f>${esc(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`
    : `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`;

const catXml = (s: ChartSeries) => {
  if (!s.catRef && !s.cats.length) return '';
  return `<c:cat>${strRef(s.catRef, s.cats)}</c:cat>`;
};
const valXml = (s: ChartSeries) => `<c:val>${numRef(s.valRef, s.vals)}</c:val>`;
const txXml = (s: ChartSeries) => `<c:tx>${strRef(s.nameRef, [s.name])}</c:tx>`;

function serXml(type: ChartSpec['type'], s: ChartSeries, i: number): string {
  const color = CHART_COLORS[i % CHART_COLORS.length];
  const head = `<c:idx val="${i}"/><c:order val="${i}"/>${txXml(s)}`;
  const fill = `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
  if (type === 'line') return `<c:ser>${head}<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>${catXml(s)}${valXml(s)}<c:smooth val="0"/></c:ser>`;
  if (type === 'pie') return `<c:ser>${head}${catXml(s)}${valXml(s)}</c:ser>`;
  if (type === 'area') return `<c:ser>${head}${fill}${catXml(s)}${valXml(s)}</c:ser>`;
  return `<c:ser>${head}${fill}<c:invertIfNegative val="0"/>${catXml(s)}${valXml(s)}</c:ser>`;
}

const titleXml = (t: string) =>
  `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="1400" b="0"/><a:t>${esc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;

/** 按图表说明生成完整的 chart 部件 XML（标准结构） */
export function chartXml(spec: ChartSpec): string {
  const type = spec.type === 'other' ? 'col' : spec.type;
  const series = type === 'pie' ? spec.series.slice(0, 1) : spec.series;
  const sers = series.map((s, i) => serXml(type, s, i)).join('');
  const ax = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  let plot: string;
  if (type === 'pie') plot = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`;
  else {
    const chart =
      type === 'line'
        ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${ax}</c:lineChart>`
        : type === 'area'
          ? `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${ax}</c:areaChart>`
          : `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="150"/>${ax}</c:barChart>`;
    const horizontal = type === 'bar';
    plot =
      chart +
      `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="${horizontal ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
      `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="${horizontal ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
  }
  const legend = series.length > 1 || type === 'pie' ? '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>' : '';
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace ${C_NS}><c:roundedCorners val="0"/><c:chart>` +
    `${spec.title ? titleXml(spec.title) : ''}<c:autoTitleDeleted val="${spec.title ? 0 : 1}"/><c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`
  );
}

/** 在原 chart XML 上改标题和系列的引用 / 缓存（类型、系列个数没变时用）；返回 null = 改不了，要整个重新生成 */
export function patchChartXml(xml: string, spec: ChartSpec): string | null {
  const sers = [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)];
  if (sers.length !== spec.series.length) return null;
  let out = xml;
  // 系列：名称 / 分类 / 值的引用和缓存换成当前的，其它（颜色、标记、标签）原样
  sers.forEach((m, i) => {
    const s = spec.series[i];
    let x = m[0];
    const swap = (tag: 'tx' | 'cat' | 'val' | 'yVal', inner: string) => {
      const re = new RegExp(`<c:${tag}>[\\s\\S]*?</c:${tag}>`);
      if (re.test(x)) x = x.replace(re, () => `<c:${tag}>${inner}</c:${tag}>`);
    };
    swap('tx', strRef(s.nameRef, [s.name]));
    if (s.catRef || s.cats.length) {
      if (/<c:cat>/.test(x)) swap('cat', strRef(s.catRef, s.cats));
      else x = x.replace(/<c:val>/, () => `${catXml(s)}<c:val>`);
    } else x = x.replace(/<c:cat>[\s\S]*?<\/c:cat>/, '');
    swap('val', numRef(s.valRef, s.vals));
    swap('yVal', numRef(s.valRef, s.vals));
    out = out.replace(m[0], () => x);
  });
  // 标题
  const had = /<c:title>[\s\S]*?<\/c:title>/.test(out.replace(/<c:plotArea>[\s\S]*<\/c:plotArea>/, ''));
  const chartHead = out.match(/<c:chart>/);
  if (!chartHead) return null;
  const outside = out.replace(/<c:plotArea>[\s\S]*<\/c:plotArea>/, '<c:plotArea/>');
  const oldTitle = outside.match(/<c:title>[\s\S]*?<\/c:title>/)?.[0];
  if (spec.title) {
    if (had && oldTitle) {
      // 标题文字换掉，格式保留：只有一段文字时改 <a:t>，否则整个重写
      const runs = [...oldTitle.matchAll(/<a:t>[\s\S]*?<\/a:t>/g)];
      if (runs.length === 1 && !/<c:strRef>/.test(oldTitle)) out = out.replace(oldTitle, () => oldTitle.replace(/<a:t>[\s\S]*?<\/a:t>/, `<a:t>${esc(spec.title!)}</a:t>`));
      else out = out.replace(oldTitle, () => titleXml(spec.title!));
    } else out = out.replace('<c:chart>', () => `<c:chart>${titleXml(spec.title!)}`);
    out = out.replace(/<c:autoTitleDeleted val="1"\/>/, '<c:autoTitleDeleted val="0"/>');
  } else if (oldTitle) {
    out = out.replace(oldTitle, '');
    out = /<c:autoTitleDeleted\b/.test(out) ? out.replace(/<c:autoTitleDeleted val="0"\/>/, '<c:autoTitleDeleted val="1"/>') : out;
  }
  return out;
}

// ---------- 修改（带撤销） ----------

const cloneDrawings = (s: SheetData) => (s.drawings ??= []);

/** 新建图表，放在指定位置（默认放在选区右边）。返回它在 drawings 里的序号 */
export function addChart(tx: Tx, si: number, spec: ChartSpec, from: Anchor, to: Anchor): number {
  tx.snap(si);
  const sheet = tx.book.sheets[si];
  const list = cloneDrawings(sheet);
  list.push({ kind: 'chart', from, to, chart: spec, name: `图表 ${list.filter((d) => d.kind === 'chart').length + 1}` });
  return list.length - 1;
}

/** 移动 / 缩放（图片、图表都行）。to 缺省时保留 ext（一个锚点的图） */
export function moveDrawing(tx: Tx, si: number, i: number, from: Anchor, to: Anchor): void {
  const sheet = tx.book.sheets[si];
  const d = sheet.drawings?.[i];
  if (!d) return;
  tx.snap(si);
  sheet.drawings![i] = { ...d, from, to };
  delete sheet.drawings![i].ext;
}

export function deleteDrawing(tx: Tx, si: number, i: number): void {
  const sheet = tx.book.sheets[si];
  if (!sheet.drawings?.[i]) return;
  tx.snap(si);
  sheet.drawings = sheet.drawings.filter((_, k) => k !== i);
  if (!sheet.drawings.length) delete sheet.drawings;
}

export function updateChart(tx: Tx, si: number, i: number, spec: ChartSpec): void {
  const sheet = tx.book.sheets[si];
  const d = sheet.drawings?.[i];
  if (!d || d.kind !== 'chart') return;
  tx.snap(si);
  sheet.drawings![i] = { ...d, chart: spec };
}

// ---------- 写出 ----------

const anchorXml = (tag: 'from' | 'to', a: Anchor) => `<xdr:${tag}><xdr:col>${a.c}</xdr:col><xdr:colOff>${Math.round(a.cOff * EMU_PX)}</xdr:colOff><xdr:row>${a.r}</xdr:row><xdr:rowOff>${Math.round(a.rOff * EMU_PX)}</xdr:rowOff></xdr:${tag}>`;

const sameAnchor = (a: Anchor | undefined, b: Anchor | undefined) => (!a && !b) || (!!a && !!b && a.c === b.c && a.r === b.r && Math.abs(a.cOff - b.cOff) < 0.5 && Math.abs(a.rOff - b.rOff) < 0.5);
const sameExt = (a: SheetDrawing['ext'], b: SheetDrawing['ext']) => (!a && !b) || (!!a && !!b && Math.abs(a.w - b.w) < 0.5 && Math.abs(a.h - b.h) < 0.5);
const chartKey = (c: ChartSpec | undefined) => JSON.stringify(c && { t: c.type, ti: c.title ?? '', s: c.series.map((s) => [s.name, s.nameRef, s.catRef, s.valRef, s.cats, s.vals]) });

const DRAWING_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';

type PutElement = (xml: string, tag: string, content: string) => string;

/** 把 sheet.drawings 的改动写进包里（和文件一致就不动） */
export function syncDrawings(pkg: Pkg, sheetPath: string, sheet: SheetData, book: Workbook, put: PutElement): void {
  const files = pkg.files;
  const model = sheet.drawings ?? [];
  const cur = readDrawings(files, sheetPath).drawings;
  const key = (d: SheetDrawing) => `${d.part}#${d.ord}`;
  const byKey = new Map(model.filter((d) => d.part !== undefined).map((d) => [key(d), d]));
  const deleted = cur.filter((c) => !byKey.has(key(c)));
  const fresh = model.filter((d) => d.part === undefined);
  const changed = model.filter((d) => {
    if (d.part === undefined) return false;
    const c = cur.find((x) => key(x) === key(d));
    return !!c && (!sameAnchor(c.from, d.from) || !sameAnchor(c.to, d.to) || !sameExt(c.ext, d.ext) || (d.kind === 'chart' && chartKey(c.chart) !== chartKey(d.chart)));
  });
  if (!deleted.length && !fresh.length && !changed.length) return;

  // 新图表要放进哪个 drawing 部件：这张表已有的第一个，没有就新建
  let target = readRels(files, sheetPath).find((r) => r.type === REL.drawing && files[r.path])?.path;
  if (!target && fresh.length) {
    target = nextPart(files, 'xl/drawings', 'drawing', '.xml');
    files[target] = strToU8(`${DRAWING_HEAD}</xdr:wsDr>`);
    addOverride(pkg, target, CT.drawing);
    const id = addRel(files, sheetPath, REL.drawing, target);
    files[sheetPath] = strToU8(put(strFromU8(files[sheetPath]), 'drawing', `<drawing r:id="${id}"/>`));
  }

  const parts = new Set([...cur.map((c) => c.part!), ...(target ? [target] : [])]);
  for (const part of parts) {
    let xml = strFromU8(files[part]);
    let ord = -1;
    xml = xml.replace(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b([^>]*)>([\s\S]*?)<\/xdr:\1>/g, (all, _tag: string, _attrs: string, body: string) => {
      ord++;
      const k = `${part}#${ord}`;
      const gone = deleted.find((d) => key(d) === k);
      if (gone) {
        if (gone.kind === 'chart' && gone.chartPart) {
          const rid = body.match(/<c:chart\b[^>]*r:id="([^"]+)"/)?.[1];
          dropPart(pkg, gone.chartPart);
          if (rid) removeRel(files, part, rid);
        } else {
          const rid = body.match(/<a:blip\b[^>]*r:embed="([^"]+)"/)?.[1];
          if (rid) removeRel(files, part, rid);
        }
        return '';
      }
      const d = byKey.get(k);
      if (!d || !changed.includes(d)) return all;
      // 位置 / 大小：换成两个锚点（保留里面的图形 / 图片内容）
      const inner = body.replace(/<xdr:from>[\s\S]*?<\/xdr:from>|<xdr:to>[\s\S]*?<\/xdr:to>|<xdr:ext\b[^>]*\/>|<xdr:pos\b[^>]*\/>/g, '');
      if (d.kind === 'chart' && d.chart && d.chartPart && files[d.chartPart]) {
        const cur0 = cur.find((x) => key(x) === k);
        if (cur0 && chartKey(cur0.chart) !== chartKey(d.chart)) {
          const old = strFromU8(files[d.chartPart]);
          const spec = refreshSeries(book, d.chart);
          files[d.chartPart] = strToU8(cur0.chart?.type === d.chart.type ? (patchChartXml(old, spec) ?? chartXml(spec)) : chartXml(spec));
        }
      }
      const to = d.to ?? { c: d.from.c + 6, cOff: 0, r: d.from.r + 16, rOff: 0 };
      return `<xdr:twoCellAnchor editAs="oneCell">${anchorXml('from', d.from)}${anchorXml('to', to)}${inner}</xdr:twoCellAnchor>`;
    });
    if (part === target) {
      let nextId = Math.max(1, ...[...xml.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]))) + 1;
      const add = fresh.map((d) => {
        const chartPart = nextPart(files, 'xl/charts', 'chart', '.xml');
        files[chartPart] = strToU8(chartXml(refreshSeries(book, d.chart!)));
        addOverride(pkg, chartPart, CT.chart);
        const rid = addRel(files, part, REL.chart, chartPart);
        const to = d.to ?? { c: d.from.c + 6, cOff: 0, r: d.from.r + 16, rOff: 0 };
        const id = nextId++;
        return (
          `<xdr:twoCellAnchor>${anchorXml('from', d.from)}${anchorXml('to', to)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(d.name ?? `图表 ${id - 1}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
          `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
        );
      });
      xml = xml.replace(/<\/xdr:wsDr>\s*$/, `${add.join('')}</xdr:wsDr>`);
    }
    files[part] = strToU8(xml);
  }
}

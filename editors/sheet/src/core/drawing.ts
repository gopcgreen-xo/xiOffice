/**
 * 工作表里的绘图、批注（笔记）和超链接（只读显示；保存时文件里的原样保留）：
 * - 绘图（xl/drawings）：图片（xdr:pic）和图表（c:chart，柱形 / 条形 / 折线 / 面积 / 饼图），按锚点的单元格位置摆放；
 * - 批注（xl/commentsN.xml）：格子右上角红色三角，悬停看内容；
 * - 超链接：外部网址或本工作簿里的位置。
 */
import { strFromU8 } from 'fflate';
import { resolvePart, unescapeXml, xmlAttr, xmlText } from '@wz/editor-shared/xlsx';
import { parseRef, type Range } from './workbook';
import type { SheetTable } from './tableStyle';

export interface Anchor {
  /** 左上角：列、列内偏移（像素，100% 缩放）、行、行内偏移 */
  c: number;
  cOff: number;
  r: number;
  rOff: number;
}

export interface ChartSeries {
  name: string;
  /** 系列名称所在的格子（Sheet1!$B$1） */
  nameRef?: string;
  /** 分类和值的引用（显示时按当前单元格取值，取不到用文件里的缓存） */
  catRef?: string;
  valRef?: string;
  cats: string[];
  vals: number[];
}

export interface ChartSpec {
  type: 'col' | 'bar' | 'line' | 'area' | 'pie' | 'other';
  title?: string;
  series: ChartSeries[];
}

export interface SheetDrawing {
  from: Anchor;
  /** twoCellAnchor 的右下角；oneCellAnchor / absoluteAnchor 用 ext（像素） */
  to?: Anchor;
  ext?: { w: number; h: number };
  kind: 'image' | 'chart';
  image?: { data: Uint8Array; mime: string };
  chart?: ChartSpec;
  name?: string;
  /** 来自文件时：所在的 drawing 部件、在里面的第几个锚点（保存时靠它对上文件里的元素）；新建的图表没有 */
  part?: string;
  ord?: number;
  /** 图表来自文件时：chart 部件路径 */
  chartPart?: string;
}

export interface SheetNote {
  r: number;
  c: number;
  author: string;
  text: string;
  /** 线程批注（Excel 365 对话式批注）：只读 */
  threaded?: boolean;
}

export interface SheetLink {
  ref: Range;
  url?: string;
  location?: string;
  tooltip?: string;
}

const EMU_PX = 9525;
const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };

type Files = Record<string, Uint8Array>;
const text = (files: Files, p: string) => (files[p] ? strFromU8(files[p]) : '');

/** 部件的关系表：rId → { type, target（完整路径或外部网址）} */
export function relsOf(files: Files, part: string): Map<string, { type: string; target: string; external: boolean }> {
  const relsPath = part.replace(/([^/]+)$/, '_rels/$1.rels');
  const out = new Map<string, { type: string; target: string; external: boolean }>();
  for (const m of text(files, relsPath).matchAll(/<Relationship\b[^>]*>/g)) {
    const id = xmlAttr(m[0], 'Id');
    const target = xmlAttr(m[0], 'Target');
    if (!id || !target) continue;
    const external = xmlAttr(m[0], 'TargetMode') === 'External';
    out.set(id, { type: xmlAttr(m[0], 'Type') ?? '', target: external ? unescapeXml(target) : resolvePart(part, target), external });
  }
  return out;
}

function anchorOf(xml: string | undefined): Anchor | undefined {
  if (!xml) return undefined;
  const n = (tag: string) => Number(xml.match(new RegExp(`<xdr:${tag}>(-?\\d+)</xdr:${tag}>`))?.[1] ?? 0);
  return { c: n('col'), cOff: n('colOff') / EMU_PX, r: n('row'), rOff: n('rowOff') / EMU_PX };
}

function chartOf(xml: string): ChartSpec {
  const plot = xml.match(/<c:plotArea>([\s\S]*?)<\/c:plotArea>/)?.[1] ?? '';
  const kind = plot.match(/<c:(barChart|bar3DChart|lineChart|line3DChart|areaChart|pieChart|pie3DChart|doughnutChart)\b/)?.[1] ?? '';
  const dir = xmlAttr(plot.match(/<c:barDir\b[^>]*>/)?.[0] ?? '', 'val');
  const type: ChartSpec['type'] = /^bar/.test(kind) ? (dir === 'bar' ? 'bar' : 'col') : /^line/.test(kind) ? 'line' : kind === 'areaChart' ? 'area' : /pie|doughnut/.test(kind) ? 'pie' : 'other';
  const titleXml = xml.match(/<c:title>([\s\S]*?)<\/c:title>/)?.[1];
  const title = titleXml ? [...titleXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => unescapeXml(m[1])).join('') || undefined : undefined;
  const cache = (x: string | undefined) => [...(x ?? '').matchAll(/<c:pt\b[^>]*idx="(\d+)"[^>]*>\s*<c:v>([\s\S]*?)<\/c:v>/g)].sort((a, b) => Number(a[1]) - Number(b[1])).map((m) => unescapeXml(m[2]));
  const series = [...plot.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)].map((m): ChartSeries => {
    const s = m[1];
    const tx = s.match(/<c:tx>([\s\S]*?)<\/c:tx>/)?.[1];
    const nameRef = tx?.match(/<c:f>([\s\S]*?)<\/c:f>/)?.[1];
    const cat = s.match(/<c:cat>([\s\S]*?)<\/c:cat>/)?.[1];
    const val = s.match(/<c:val>([\s\S]*?)<\/c:val>/)?.[1] ?? s.match(/<c:yVal>([\s\S]*?)<\/c:yVal>/)?.[1];
    const f = (x: string | undefined) => {
      const v = x?.match(/<c:f>([\s\S]*?)<\/c:f>/)?.[1];
      return v ? unescapeXml(v) : undefined;
    };
    return {
      name: cache(tx)[0] ?? (tx ? xmlText(tx) : ''),
      ...(nameRef && { nameRef: unescapeXml(nameRef) }),
      catRef: f(cat),
      valRef: f(val),
      cats: cache(cat),
      vals: cache(val).map(Number),
    };
  });
  return { type, ...(title ? { title } : {}), series };
}

export function readDrawings(files: Files, sheetPath: string): { drawings: SheetDrawing[]; notes: SheetNote[]; links: SheetLink[]; tables: SheetTable[] } {
  const rels = relsOf(files, sheetPath);
  const sheetXml = text(files, sheetPath);
  const drawings: SheetDrawing[] = [];
  const notes: SheetNote[] = [];
  const links: SheetLink[] = [];
  const tables: SheetTable[] = [];
  // 超链接
  for (const m of sheetXml.matchAll(/<hyperlink\b[^>]*>/g)) {
    const ref = parseRef(xmlAttr(m[0], 'ref') ?? '');
    if (!ref) continue;
    const rid = xmlAttr(m[0], 'r:id');
    const loc = xmlAttr(m[0], 'location');
    const tip = xmlAttr(m[0], 'tooltip');
    links.push({ ref, ...(rid && rels.get(rid)?.external ? { url: rels.get(rid)!.target } : {}), ...(loc ? { location: unescapeXml(loc) } : {}), ...(tip ? { tooltip: unescapeXml(tip) } : {}) });
  }
  for (const rel of rels.values()) {
    // 表格（套用表格格式）：只读样式信息
    if (/\/table$/.test(rel.type)) {
      const x = text(files, rel.target);
      const head = x.match(/<table\b[^>]*>/)?.[0] ?? '';
      const ref = parseRef(xmlAttr(head, 'ref') ?? '');
      if (!ref) continue;
      const info = x.match(/<tableStyleInfo\b[^>]*>/)?.[0] ?? '';
      const on = (n: string) => ['1', 'true'].includes(xmlAttr(info, n) ?? '');
      tables.push({
        path: rel.target,
        ref,
        headerRows: Number(xmlAttr(head, 'headerRowCount') ?? 1),
        totalsRows: Number(xmlAttr(head, 'totalsRowCount') ?? 0),
        ...(xmlAttr(info, 'name') ? { style: xmlAttr(info, 'name') } : {}),
        rowStripes: on('showRowStripes'),
        firstCol: on('showFirstColumn'),
        lastCol: on('showLastColumn'),
        ...(/<autoFilter\b/.test(x) ? { filter: true } : {}),
      });
    }
    // 批注（笔记）
    if (/\/comments$/.test(rel.type)) {
      const x = text(files, rel.target);
      const authors = [...(x.match(/<authors>([\s\S]*?)<\/authors>/)?.[1] ?? '').matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => unescapeXml(m[1]));
      for (const m of x.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)) {
        const ref = parseRef(xmlAttr(m[1], 'ref') ?? '');
        if (!ref) continue;
        notes.push({ r: ref.r0, c: ref.c0, author: authors[Number(xmlAttr(m[1], 'authorId') ?? 0)] ?? '', text: xmlText(m[2]).trim() });
      }
    }
    // 绘图：图片和图表
    if (/\/drawing$/.test(rel.type)) {
      const x = text(files, rel.target);
      const drels = relsOf(files, rel.target);
      let ord = -1;
      for (const m of x.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/xdr:\1>/g)) {
        ord++;
        const body = m[2];
        const from = anchorOf(body.match(/<xdr:from>([\s\S]*?)<\/xdr:from>/)?.[1]) ?? { c: 0, cOff: 0, r: 0, rOff: 0 };
        const to = anchorOf(body.match(/<xdr:to>([\s\S]*?)<\/xdr:to>/)?.[1]);
        const extTag = body.match(/<xdr:ext\b[^>]*>/)?.[0];
        const ext = extTag ? { w: Number(xmlAttr(extTag, 'cx')) / EMU_PX, h: Number(xmlAttr(extTag, 'cy')) / EMU_PX } : undefined;
        const name = xmlAttr(body.match(/<xdr:cNvPr\b[^>]*>/)?.[0] ?? '', 'name');
        const base = { from, ...(to ? { to } : {}), ...(ext ? { ext } : {}), ...(name ? { name: unescapeXml(name) } : {}), part: rel.target, ord };
        const embed = body.match(/<a:blip\b[^>]*r:embed="([^"]+)"/)?.[1];
        if (embed && /<xdr:pic>/.test(body)) {
          const target = drels.get(embed)?.target;
          const data = target ? files[target] : undefined;
          const mime = MIME[(target ?? '').split('.').pop()?.toLowerCase() ?? ''];
          if (data && mime) drawings.push({ ...base, kind: 'image', image: { data, mime } });
          continue;
        }
        const chartId = body.match(/<c:chart\b[^>]*r:id="([^"]+)"/)?.[1];
        const chartPath = chartId ? drels.get(chartId)?.target : undefined;
        if (chartPath && files[chartPath]) drawings.push({ ...base, kind: 'chart', chart: chartOf(text(files, chartPath)), chartPart: chartPath });
      }
    }
  }
  return { drawings, notes, links, tables };
}

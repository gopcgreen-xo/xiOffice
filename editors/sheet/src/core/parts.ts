/**
 * xlsx 包（zip 里的部件）操作工具：关系文件（.rels）、内容类型（[Content_Types].xml）、新部件命名。
 * 批注、图表、透视表这些“新增部件”的写出都用它；只加 / 改 / 删自己的部件和关系，其它部件不碰。
 */
import { strFromU8, strToU8 } from 'fflate';
import { resolvePart, unescapeXml, xmlAttr } from '@wz/editor-shared/xlsx';

export type Files = Record<string, Uint8Array>;

/** 正在写出的包：文件表 + 内容类型文字（saveXlsx 里两者是分开存的，用完再写回） */
export interface Pkg {
  files: Files;
  ct: string;
}

export const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const REL = {
  comments: `${OD}/comments`,
  vml: `${OD}/vmlDrawing`,
  drawing: `${OD}/drawing`,
  chart: `${OD}/chart`,
  pivotTable: `${OD}/pivotTable`,
  pivotCacheDef: `${OD}/pivotCacheDefinition`,
  pivotCacheRec: `${OD}/pivotCacheRecords`,
  image: `${OD}/image`,
  worksheet: `${OD}/worksheet`,
} as const;

export const CT = {
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  pivotCacheDef: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRec: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
} as const;

const text = (files: Files, p: string) => (files[p] ? strFromU8(files[p]) : '');
export const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** 部件 xl/worksheets/sheet1.xml → 它的关系文件 xl/worksheets/_rels/sheet1.xml.rels */
export const relsPathOf = (part: string) => part.replace(/([^/]+)$/, '_rels/$1.rels');

export interface Rel {
  id: string;
  type: string;
  /** 文件里写的 Target 原文 */
  target: string;
  external: boolean;
  /** 内部关系：目标部件的完整路径 */
  path: string;
}

export function readRels(files: Files, part: string): Rel[] {
  const out: Rel[] = [];
  for (const m of text(files, relsPathOf(part)).matchAll(/<Relationship\b[^>]*>/g)) {
    const id = xmlAttr(m[0], 'Id');
    const target = xmlAttr(m[0], 'Target');
    if (!id || !target) continue;
    const external = xmlAttr(m[0], 'TargetMode') === 'External';
    out.push({ id, type: xmlAttr(m[0], 'Type') ?? '', target: unescapeXml(target), external, path: external ? unescapeXml(target) : resolvePart(part, target) });
  }
  return out;
}

/** from 部件指向 to 部件的相对路径（Target 的写法）：xl/worksheets/sheet1.xml → xl/comments1.xml = ../comments1.xml */
export function relTarget(from: string, to: string): string {
  const a = from.split('/').slice(0, -1);
  const b = to.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return [...Array(a.length - i).fill('..'), ...b.slice(i)].join('/');
}

/** 给部件加一条关系（关系文件不存在就新建）；返回 rId */
export function addRel(files: Files, part: string, type: string, targetPart: string): string {
  const path = relsPathOf(part);
  let xml = text(files, path) || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG_REL}"></Relationships>`;
  const used = [...xml.matchAll(/\bId="rId(\d+)"/g)].map((m) => Number(m[1]));
  const id = `rId${Math.max(0, ...used) + 1}`;
  const rel = `<Relationship Id="${id}" Type="${type}" Target="${esc(relTarget(part, targetPart))}"/>`;
  xml = /<Relationships\b[^>]*\/>/.test(xml) ? xml.replace(/<Relationships\b([^>]*)\/>/, `<Relationships$1>${rel}</Relationships>`) : xml.replace(/<\/Relationships>\s*$/, `${rel}</Relationships>`);
  files[path] = strToU8(xml);
  return id;
}

/** 去掉一条关系；关系文件空了就删掉 */
export function removeRel(files: Files, part: string, id: string): void {
  const path = relsPathOf(part);
  if (!files[path]) return;
  const xml = text(files, path).replace(new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*/>`), '');
  if (/<Relationship\b/.test(xml)) files[path] = strToU8(xml);
  else delete files[path];
}

/** dir 目录下第一个没用过的 stem + 序号 + ext：xl/charts/chart3.xml */
export function nextPart(files: Files, dir: string, stem: string, ext: string): string {
  for (let i = 1; ; i++) {
    const p = `${dir}/${stem}${i}${ext}`;
    if (!files[p]) return p;
  }
}

const reEsc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function addOverride(pkg: Pkg, part: string, type: string): void {
  const tag = `<Override PartName="/${part}" ContentType="${type}"/>`;
  if (new RegExp(`PartName="/${reEsc(part)}"`).test(pkg.ct)) return;
  pkg.ct = pkg.ct.replace(/<\/Types>\s*$/, `${tag}</Types>`);
}

export function removeOverride(pkg: Pkg, part: string): void {
  pkg.ct = pkg.ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${reEsc(part)}"[^>]*/>`), '');
}

/** 没有这个扩展名的默认类型就加一条（如 vml） */
export function ensureDefault(pkg: Pkg, ext: string, type: string): void {
  if (new RegExp(`<Default\\b[^>]*Extension="${reEsc(ext)}"`, 'i').test(pkg.ct)) return;
  pkg.ct = pkg.ct.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="${ext}" ContentType="${type}"/>`);
}

/** 删掉一个部件和它的关系文件、内容类型 */
export function dropPart(pkg: Pkg, part: string): void {
  delete pkg.files[part];
  delete pkg.files[relsPathOf(part)];
  removeOverride(pkg, part);
}

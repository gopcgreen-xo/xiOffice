/**
 * 批注（Excel 里的“备注 / 注释”）：新建、修改、删除，并把改动写进 xlsx 的 comments 部件和 VML 绘图部件。
 *
 * 写出时尽量少动：和文件里一致就一个字节不改；改过的批注重新生成（纯文字），没改的批注保留原来的 <comment>（保留富文本格式），
 * VML 里没改的批注框原样保留，只增删有变化的。“线程批注”（Excel 365 的对话式批注，文件里带一份 [Threaded comment] 占位文字）只读。
 */
import { strFromU8, strToU8 } from 'fflate';
import { unescapeXml, xmlAttr, xmlText } from '@wz/editor-shared/xlsx';
import type { SheetNote } from './drawing';
import type { Tx } from './ops';
import { addOverride, addRel, CT, dropPart, ensureDefault, esc, nextPart, readRels, REL, removeRel, type Files, type Pkg } from './parts';
import { parseRef, type SheetData } from './workbook';

const at = (r: number, c: number) => `${r}:${c}`;

export const THREADED_MARK = '[Threaded comment]';
/** 线程批注（文件里是占位文字）：只读 */
export const isThreaded = (n: SheetNote) => !!n.threaded || n.text.startsWith(THREADED_MARK);

// ---------- 修改（带撤销） ----------

/** 新建或修改格子 (r, c) 的批注。文字为空等于删除 */
export function setNote(tx: Tx, si: number, r: number, c: number, text: string, author: string): boolean {
  const sheet = tx.book.sheets[si];
  const old = sheet.notes?.find((n) => n.r === r && n.c === c);
  if (old && isThreaded(old)) return false;
  if (!text.trim()) {
    deleteNote(tx, si, r, c);
    return true;
  }
  tx.snap(si);
  const rest = (sheet.notes ?? []).filter((n) => !(n.r === r && n.c === c));
  sheet.notes = [...rest, { r, c, author: old?.author ?? author, text }].sort((a, b) => a.r - b.r || a.c - b.c);
  return true;
}

export function deleteNote(tx: Tx, si: number, r: number, c: number): void {
  const sheet = tx.book.sheets[si];
  if (!sheet.notes?.some((n) => n.r === r && n.c === c)) return;
  tx.snap(si);
  const rest = sheet.notes.filter((n) => !(n.r === r && n.c === c));
  if (rest.length) sheet.notes = rest;
  else delete sheet.notes;
}

// ---------- 写出 ----------

interface Existing {
  r: number;
  c: number;
  author: string;
  text: string;
  raw: string;
}

function parseComments(xml: string): Existing[] {
  const authors = [...(xml.match(/<authors>([\s\S]*?)<\/authors>/)?.[1] ?? '').matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => unescapeXml(m[1]));
  const out: Existing[] = [];
  for (const m of xml.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)) {
    const ref = parseRef(xmlAttr(m[1], 'ref') ?? '');
    if (!ref) continue;
    out.push({ r: ref.r0, c: ref.c0, author: authors[Number(xmlAttr(m[1], 'authorId') ?? 0)] ?? '', text: xmlText(m[2]).trim(), raw: m[0] });
  }
  return out;
}

const cellRef = (r: number, c: number) => {
  let col = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) col = String.fromCharCode(65 + ((n - 1) % 26)) + col;
  return `${col}${r + 1}`;
};

function shapeXml(id: number, r: number, c: number): string {
  const anchor = `${c + 1}, 15, ${Math.max(r - 1, 0)}, 10, ${c + 3}, 15, ${r + 3}, 16`;
  return (
    `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:60pt;z-index:${id % 1024};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
    `<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>` +
    `<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>` +
    `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`
  );
}

const VML_HEAD = (idmap: number) =>
  `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
  `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${idmap}"/></o:shapelayout>` +
  `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>`;

type PutElement = (xml: string, tag: string, content: string) => string;

/**
 * 把 sheet.notes 同步到包里：和文件里的批注一致就不动；否则重写 comments 部件和 VML，必要时新建 / 删除部件和关系。
 * put = save.ts 的 putElement（按 schema 顺序放 <legacyDrawing>）
 */
export function syncNotes(pkg: Pkg, sheetPath: string, sheet: SheetData, put: PutElement): void {
  const files: Files = pkg.files;
  const rels = readRels(files, sheetPath);
  const cRel = rels.find((r) => r.type === REL.comments && files[r.path]);
  const vRel = rels.find((r) => r.type === REL.vml && files[r.path]);
  const notes = sheet.notes ?? [];
  const existing = cRel ? parseComments(strFromU8(files[cRel.path])) : [];
  const match = (n: SheetNote) => existing.find((e) => e.r === n.r && e.c === n.c && e.text === n.text.trim() && e.author === n.author);
  if (existing.length === notes.length && notes.every((n) => match(n))) return;

  // 全删了：去掉批注部件，VML 里的批注框也去掉（VML 里没别的形状了就整个删）
  if (!notes.length) {
    if (cRel) {
      dropPart(pkg, cRel.path);
      removeRel(files, sheetPath, cRel.id);
    }
    if (vRel) {
      const rest = strFromU8(files[vRel.path]).replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (s) => (/ObjectType="Note"/.test(s) ? '' : s));
      if (/<v:shape\b/.test(rest)) files[vRel.path] = strToU8(rest);
      else {
        dropPart(pkg, vRel.path);
        removeRel(files, sheetPath, vRel.id);
        files[sheetPath] = strToU8(strFromU8(files[sheetPath]).replace(new RegExp(`<legacyDrawing\\b[^>]*r:id="${vRel.id}"[^>]*/>`), ''));
      }
    }
    return;
  }

  // comments 部件
  const authors = [...new Set(notes.map((n) => n.author || '作者'))];
  const list = [...notes].sort((a, b) => a.r - b.r || a.c - b.c);
  const body = list
    .map((n) => {
      const id = authors.indexOf(n.author || '作者');
      const old = match(n);
      if (old) return old.raw.replace(/authorId="\d+"/, `authorId="${id}"`);
      return `<comment ref="${cellRef(n.r, n.c)}" authorId="${id}"><text><t xml:space="preserve">${esc(n.text)}</t></text></comment>`;
    })
    .join('');
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<authors>${authors.map((a) => `<author>${esc(a)}</author>`).join('')}</authors><commentList>${body}</commentList></comments>`;
  let cPath = cRel?.path;
  if (!cPath) {
    cPath = nextPart(files, 'xl', 'comments', '.xml');
    addOverride(pkg, cPath, CT.comments);
    addRel(files, sheetPath, REL.comments, cPath);
  }
  files[cPath] = strToU8(xml);

  // VML：没改的批注框保留，缺的补上，多的去掉
  const have = new Set<string>();
  let vml = vRel ? strFromU8(files[vRel.path]) : '';
  const alive = new Set(notes.map((n) => at(n.r, n.c)));
  vml = vml.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (s) => {
    if (!/ObjectType="Note"/.test(s)) return s;
    const r = Number(s.match(/<x:Row>(\d+)<\/x:Row>/)?.[1]);
    const c = Number(s.match(/<x:Column>(\d+)<\/x:Column>/)?.[1]);
    if (!alive.has(at(r, c))) return '';
    have.add(at(r, c));
    return s;
  });
  if (!vml) {
    // 新绘图的 idmap 要在整个工作簿里唯一（形状编号 = idmap × 1024 + 序号）
    const taken = Object.keys(files)
      .filter((p) => /\.vml$/i.test(p))
      .map((p) => Number(strFromU8(files[p]).match(/<o:idmap\b[^>]*data="(\d+)"/)?.[1] ?? 0));
    vml = VML_HEAD(Math.max(0, ...taken) + 1) + '</xml>';
  }
  const used = [...vml.matchAll(/id="_x0000_s(\d+)"/g)].map((m) => Number(m[1]));
  const idmap = Number(vml.match(/<o:idmap\b[^>]*data="(\d+)"/)?.[1] ?? 1);
  let next = Math.max(1024 * idmap, ...used) + 1;
  const add = notes.filter((n) => !have.has(at(n.r, n.c))).map((n) => shapeXml(next++, n.r, n.c)).join('');
  vml = vml.replace(/<\/xml>\s*$/, `${add}</xml>`);
  let vPath = vRel?.path;
  if (!vPath) {
    vPath = nextPart(files, 'xl/drawings', 'vmlDrawing', '.vml');
    ensureDefault(pkg, 'vml', CT.vml);
    const id = addRel(files, sheetPath, REL.vml, vPath);
    files[sheetPath] = strToU8(put(strFromU8(files[sheetPath]), 'legacyDrawing', `<legacyDrawing r:id="${id}"/>`));
  }
  files[vPath] = strToU8(vml);
}

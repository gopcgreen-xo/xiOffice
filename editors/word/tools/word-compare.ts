/**
 * 排版对照（需求 W19 的验收工具）：同一份 docx，本编辑器排版的结果和 Microsoft Word 导出的 PDF 逐页比较。
 * 比较页数，以及每页正文第一行的开头、最后一行的结尾（页眉页脚按页边距排除）。
 *
 * 用法：data/ 下放 fonts.json（npx vite-node scripts/font-metrics.ts）、 <名字>.docx（或 <名字>.doc）和 Word 导出的 <名字>.pdf（scripts/word-pdf.ps1，
 * 导出卡住时可用 Word 打印到 Microsoft Print to PDF）以及 index.json（名字列表，.doc 写全名 <名字>.doc），
 * npx vite 后打开 /editors/word/tools/word-compare.html（?names=a,b.doc 只比这几份）；结果在页面上，也放在 window.__result（脚本读取）。
 */
import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { openDocx } from '../src/ooxml/patch';
import { isDoc, readDoc } from '../src/doc97/read';
import { layoutDocument, type DocLayout } from '../src/core/layout';
import { CanvasMeasurer, drawPage } from '../src/core/render';
import { loadImages } from '../src/core/images';
import { imageIds } from '../src/core/model';
import { registerFonts, type FaceInfo } from '../src/core/fontfile';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

// index.json 里的名字不带扩展名时是 <名字>.docx；写成 <名字>.doc 时读 Word 97-2003 文件。PDF 都是 <名字>.pdf
const docUrl = (name: string) => `./data/${encodeURIComponent(/\.doc$/i.test(name) ? name : `${name}.docx`)}`;
const pdfUrl = (name: string) => `./data/${encodeURIComponent(name.replace(/\.doc$/i, ''))}.pdf`;
const openAny = async (bytes: Uint8Array) => (isDoc(bytes) ? { doc: readDoc(bytes) } : openDocx(bytes));

const norm = (s: string) => s.replace(/[\s ￼]/g, '');

export interface PageCheck {
  page: number;
  ok: boolean;
  ours: { first: string; last: string };
  word: { first: string; last: string };
}

export interface DocCheck {
  name: string;
  ourPages: number;
  wordPages: number;
  pages: PageCheck[];
  ok: boolean;
}

/** 调试：一份文档两边每页每行的 y 和文字（window.__dump(名字)） */
async function dump(name: string) {
  const [d, p] = await Promise.all([fetch(docUrl(name)), fetch(pdfUrl(name))]);
  const { doc } = await openAny(new Uint8Array(await d.arrayBuffer()));
  const L = layoutDocument(doc, new CanvasMeasurer());
  const ours = L.pages.map((pg) => [...pg.body, ...(pg.notes ?? [])].sort((a, b) => a.y - b.y || a.x - b.x).map((l) => `${l.y.toFixed(1)}+${l.line.height.toFixed(2)} ${norm(l.line.glyphs.map((g) => g.ch).join('')).slice(0, 16)}`));
  const pdf = await pdfjs.getDocument({ data: await p.arrayBuffer() }).promise;
  const word: string[][] = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const h = page.view[3];
    const items = (await page.getTextContent()).items.filter((it): it is { str: string; transform: number[]; height: number } & typeof it => 'str' in it && !!it.str.trim());
    const rows = new Map<number, string>();
    for (const it of items) {
      const y = Math.round((h - it.transform[5]) * 2) / 2;
      const key = [...rows.keys()].find((k) => Math.abs(k - y) < 3) ?? y;
      rows.set(key, (rows.get(key) ?? '') + it.str);
    }
    word.push([...rows.entries()].sort((a, b) => a[0] - b[0]).map(([y, s]) => `${y.toFixed(1)} ${norm(s).slice(0, 16)}`));
  }
  return { ours, word };
}
(window as unknown as { __dump: typeof dump }).__dump = dump;

/** 调试：本编辑器里以 prefix 开头的那一行每个字形的 x（页面坐标） */
async function glyphs(name: string, prefix: string) {
  const d = await fetch(docUrl(name));
  const { doc } = await openAny(new Uint8Array(await d.arrayBuffer()));
  const L = layoutDocument(doc, new CanvasMeasurer());
  for (const pg of L.pages)
    for (const l of pg.body) {
      const text = l.line.glyphs.map((g) => g.ch).join('');
      if (norm(text).startsWith(prefix)) return { x0: l.x, glyphs: l.line.glyphs.map((g) => `${g.ch}:${(l.x + g.x).toFixed(2)}/${g.w.toFixed(2)}`).join(' ') };
    }
  return null;
}
(window as unknown as { __glyphs: typeof glyphs }).__glyphs = glyphs;

/** 调试：第 page 页左边本编辑器、右边 Word 的 PDF，拼成一张图（window.__side(名字, 页)） */
async function side(name: string, page: number, scale = 1.5) {
  const [d, p] = await Promise.all([fetch(docUrl(name)), fetch(pdfUrl(name))]);
  const { doc } = await openAny(new Uint8Array(await d.arrayBuffer()));
  await loadImages(imageIds(doc));
  const L = layoutDocument(doc, new CanvasMeasurer());
  const pg = L.pages[page - 1];
  const w = Math.round(pg.setup.width * scale);
  const h = Math.round(pg.setup.height * scale);
  const c = document.createElement('canvas');
  c.width = w * 2 + 10;
  c.height = h;
  const ctx = c.getContext('2d')!;
  ctx.fillStyle = '#888';
  ctx.fillRect(0, 0, c.width, c.height);
  const ours = document.createElement('canvas');
  ours.width = w;
  ours.height = h;
  drawPage(ours.getContext('2d')!, pg, scale, { story: 'body', selection: [], showMarks: false, print: true });
  ctx.drawImage(ours, 0, 0);
  const pdf = await pdfjs.getDocument({ data: await p.arrayBuffer() }).promise;
  const pp = await pdf.getPage(page);
  const vp = pp.getViewport({ scale });
  const wc = document.createElement('canvas');
  wc.width = Math.round(vp.width);
  wc.height = Math.round(vp.height);
  await pp.render({ canvasContext: wc.getContext('2d')!, viewport: vp }).promise;
  ctx.drawImage(wc, w + 10, 0);
  return c.toDataURL('image/png');
}
(window as unknown as { __side: typeof side }).__side = side;

/** 调试：第 page 页的线段（window.__borders(名字, 页)） */
async function borders(name: string, page: number) {
  const d = await fetch(docUrl(name));
  const { doc } = await openAny(new Uint8Array(await d.arrayBuffer()));
  return layoutDocument(doc, new CanvasMeasurer()).pages[page - 1].borders.map((b) => `${b.x1.toFixed(0)},${b.y1.toFixed(1)}-${b.x2.toFixed(0)},${b.y2.toFixed(1)} ${b.w ?? ''} ${b.color ?? ''}`);
}
(window as unknown as { __borders: typeof borders }).__borders = borders;

/** 本编辑器每页正文的行（按位置从上到下、从左到右），文字 = 显示的字形；纸张以外（被裁掉、看不见）的行不算 */
function ourLines(L: DocLayout): string[][] {
  return L.pages.map((pg) =>
    [...pg.body, ...(pg.notes ?? [])]
      .filter((l) => l.x < pg.setup.width && l.y < pg.setup.height)
      .sort((a, b) => a.y - b.y || a.x - b.x)
      .map((l) => norm(l.line.glyphs.filter((g) => l.x + g.x < pg.setup.width).map((g) => g.ch).join('')))
      .filter(Boolean),
  );
}

/** Word PDF 每页正文的行：按基线 y 合并同一行的文字，去掉页边距以外（页眉页脚）的行 */
async function wordLines(pdfData: ArrayBuffer, L: DocLayout): Promise<string[][]> {
  const pdf = await pdfjs.getDocument({ data: pdfData }).promise;
  const out: string[][] = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const h = page.view[3] - page.view[1];
    const setup = (L.pages[i - 1] ?? L.pages[L.pages.length - 1]).setup;
    const content = await page.getTextContent();
    const items = content.items.filter((it): it is { str: string; transform: number[] } & typeof it => 'str' in it && !!it.str.trim());
    const rows: { y: number; parts: { x: number; s: string }[] }[] = [];
    for (const it of items) {
      const x = it.transform[4];
      const y = h - it.transform[5]; // 从上往下
      if (y < setup.marginTop - 2 || y > h - setup.marginBottom + 6) continue;
      const row = rows.find((r) => Math.abs(r.y - y) < 3);
      if (row) row.parts.push({ x, s: it.str });
      else rows.push({ y, parts: [{ x, s: it.str }] });
    }
    rows.sort((a, b) => a.y - b.y);
    out.push(rows.map((r) => norm(r.parts.sort((a, b) => a.x - b.x).map((p) => p.s).join(''))).filter(Boolean));
  }
  return out;
}

/** 两行“是同一行”：一行的开头 / 结尾在另一行里（表格一行有好几个单元格，PDF 会合成一行） */
function sameStart(a: string, b: string): boolean {
  const k = Math.min(6, a.length, b.length);
  if (!k) return a === b;
  return a.startsWith(b.slice(0, k)) || b.startsWith(a.slice(0, k)) || a.includes(b.slice(0, k)) || b.includes(a.slice(0, k));
}
function sameEnd(a: string, b: string): boolean {
  const k = Math.min(6, a.length, b.length);
  if (!k) return a === b;
  return a.endsWith(b.slice(-k)) || b.endsWith(a.slice(-k)) || a.includes(b.slice(-k)) || b.includes(a.slice(-k));
}

export async function compare(name: string, docx: ArrayBuffer, pdf: ArrayBuffer): Promise<DocCheck> {
  const { doc } = await openAny(new Uint8Array(docx));
  await loadImages(imageIds(doc));
  const L = layoutDocument(doc, new CanvasMeasurer());
  const ours = ourLines(L);
  const word = await wordLines(pdf, L);
  const pages: PageCheck[] = [];
  for (let i = 0; i < Math.max(ours.length, word.length); i++) {
    const o = ours[i] ?? [];
    const w = word[i] ?? [];
    const first = { ours: o[0] ?? '', word: w[0] ?? '' };
    const last = { ours: o[o.length - 1] ?? '', word: w[w.length - 1] ?? '' };
    pages.push({
      page: i + 1,
      ok: sameStart(first.ours, first.word) && sameEnd(last.ours, last.word),
      ours: { first: first.ours.slice(0, 20), last: last.ours.slice(-20) },
      word: { first: first.word.slice(0, 20), last: last.word.slice(-20) },
    });
  }
  return { name, ourPages: ours.length, wordPages: word.length, pages, ok: ours.length === word.length && pages.every((p) => p.ok) };
}

async function main() {
  await document.fonts.ready;
  // 本机字体的度量（npx vite-node scripts/font-metrics.ts 生成）；没有时用画布量的近似值
  const fonts = await fetch('./data/fonts.json').catch(() => null);
  if (fonts?.ok) registerFonts((await fonts.json()) as FaceInfo[]);
  // ?names=a,b.doc 只比较这几份
  const only = new URLSearchParams(location.search).get('names');
  const names: string[] = only ? only.split(',') : await (await fetch('./data/index.json')).json();
  const results: DocCheck[] = [];
  for (const name of names) {
    const [d, p] = await Promise.all([fetch(docUrl(name)), fetch(pdfUrl(name))]);
    try {
      results.push(await compare(name, await d.arrayBuffer(), await p.arrayBuffer()));
    } catch (e) {
      results.push({ name, ourPages: 0, wordPages: 0, pages: [], ok: false });
      console.error(name, e);
    }
  }
  (window as unknown as { __result: DocCheck[] }).__result = results;
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
  document.body.innerHTML = results
    .map(
      (r) =>
        `<h3>${r.ok ? '✔' : '✘'} ${esc(r.name)}：本编辑器 ${r.ourPages} 页 / Word ${r.wordPages} 页</h3><table border="1" cellpadding="4">` +
        r.pages.map((p) => `<tr style="color:${p.ok ? 'green' : 'red'}"><td>${p.page}</td><td>${esc(p.ours.first)} … ${esc(p.ours.last)}</td><td>${esc(p.word.first)} … ${esc(p.word.last)}</td></tr>`).join('') +
        '</table>',
    )
    .join('');
}

void main();

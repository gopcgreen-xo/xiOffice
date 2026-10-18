/**
 * 超链接与书签。
 * 链接是文字格式的一部分（RunProps.link：外部地址 url / 原文件的关系 rid / 文档内书签 anchor），
 * 书签是一对不占宽度的原样对象（bookmarkStart / bookmarkEnd），保存时原样写回。
 */
import { FIELD_CHAR, isAtom, paraText, type Doc, type Opaque, type Paragraph, type Run, type RunProps } from './model';
import { comparePos, type Pos } from './layout';
import { getPara, mapRuns, paragraphPaths, paraLength, runPropsAt } from './ops';

/** Word 的“超链接”字符样式的外观：蓝色加下划线 */
export const LINK_COLOR = '0563C1';

export type LinkTarget = NonNullable<RunProps['link']>;

/** 链接的显示：外部地址或“文档内：书签名” */
export function linkLabel(l: LinkTarget): string {
  return l.anchor ? `文档内：${l.anchor}` : (l.url ?? '（原文件里的链接）');
}

const sameLink = (a?: LinkTarget, b?: LinkTarget) => !!a && !!b && a.rid === b.rid && a.anchor === b.anchor && a.url === b.url;

/** 补全地址：没写协议的补 https://；已有 http / https / mailto / tel / file 等原样 */
export function normalizeUrl(input: string): string {
  const u = input.trim();
  if (!u) return '';
  if (/^[a-z][a-z0-9+.-]*:/i.test(u)) return u;
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(u)) return `mailto:${u}`;
  return `https://${u}`;
}

/** 给选区加链接（外观同 Word：蓝色下划线，原来有别的颜色的也改成链接色） */
export function setLink(doc: Doc, a: Pos, b: Pos, link: LinkTarget): void {
  mapRuns(doc, a, b, (p, r) => (r.opaque || r.mark ? p : { ...p, link: { ...link }, color: LINK_COLOR, underline: true }), false);
}

/** 去掉选区里的链接（链接外观一并还原） */
export function clearLink(doc: Doc, a: Pos, b: Pos): void {
  mapRuns(
    doc,
    a,
    b,
    (p) => {
      if (!p.link) return p;
      const { link: _link, ...rest } = p;
      return p.color === LINK_COLOR && p.underline ? { ...rest, color: 'auto', underline: false } : rest;
    },
    false,
  );
}

/** 光标（或选区起点）所在的链接：返回链接和它的完整范围（同一个链接连续的文字） */
export function linkAt(doc: Doc, pos: Pos): { link: LinkTarget; from: Pos; to: Pos; text: string } | null {
  const p = getPara(doc, pos);
  let acc = 0;
  let hit = -1;
  for (let i = 0; i < p.runs.length; i++) {
    const r = p.runs[i];
    const end = acc + r.text.length;
    // 光标在文字中间或末尾都算（在链接后面立即输入的字属于链接，这里与 Word 一致）
    if (r.props.link && pos.offset > acc && pos.offset <= end) hit = i;
    else if (r.props.link && pos.offset === acc && hit < 0) hit = i;
    acc = end;
    if (hit >= 0 && acc >= pos.offset) break;
  }
  if (hit < 0) return null;
  const link = p.runs[hit].props.link!;
  let lo = hit;
  let hi = hit;
  while (lo > 0 && sameLink(p.runs[lo - 1].props.link, link)) lo--;
  while (hi < p.runs.length - 1 && sameLink(p.runs[hi + 1].props.link, link)) hi++;
  const start = p.runs.slice(0, lo).reduce((n, r) => n + r.text.length, 0);
  const text = p.runs.slice(lo, hi + 1).map((r) => r.text).join('');
  return { link, from: { ...pos, offset: start }, to: { ...pos, offset: start + text.length }, text };
}

// ---------------- 书签 ----------------

/** 书签名规则（Word）：以字母或汉字开头，只含字母、数字、下划线，不超过 40 个字符，不含空格 */
export function checkBookmarkName(name: string): string | null {
  if (!name) return '请输入书签名';
  if ([...name].length > 40) return '书签名最长 40 个字符';
  if (/\s/.test(name)) return '书签名不能含空格';
  if (!/^[A-Za-z一-鿿]/.test(name)) return '书签名要以字母或汉字开头';
  if (!/^[A-Za-z0-9_一-鿿]+$/.test(name)) return '书签名只能含字母、汉字、数字和下划线';
  return null;
}

const escAttr = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const unescAttr = (s: string) => s.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

interface Mark {
  kind: 'start' | 'end';
  id: string;
  name?: string;
}

/** 这个文本段是不是书签标记 */
function markOf(r: Run): Mark | null {
  const o = r.opaque;
  if (!o || o.kind !== 'bookmark') return null;
  const start = /^<w:bookmarkStart\b/.test(o.xml);
  const end = /^<w:bookmarkEnd\b/.test(o.xml);
  if (!start && !end) return null;
  const id = o.xml.match(/\bw:id="([^"]*)"/)?.[1] ?? '';
  const name = start ? unescAttr(o.xml.match(/\bw:name="([^"]*)"/)?.[1] ?? '') : undefined;
  return { kind: start ? 'start' : 'end', id, name };
}

/** 正文、表格、页眉页脚里所有段落（带路径），书签只在正文里找 */
function bodyParas(doc: Doc): { path: number[]; p: Paragraph }[] {
  return paragraphPaths(doc, 'body').map((path) => ({ path, p: getPara(doc, { story: 'body', path }) }));
}

export interface BookmarkInfo {
  name: string;
  /** 起点（文档顺序）；结束标记在别处时 end 为 null */
  start: Pos;
  end: Pos | null;
  /** 书签覆盖的文字（最多 40 个字符，显示用） */
  text: string;
}

/** 文档里的书签（文档顺序）。名字以下划线开头的是 Word 的隐藏书签（目录、交叉引用用），默认不列出。 */
export function listBookmarks(doc: Doc, hidden = false): BookmarkInfo[] {
  const paras = bodyParas(doc);
  const starts: { name: string; id: string; pos: Pos }[] = [];
  const ends = new Map<string, Pos>();
  for (const { path, p } of paras) {
    let acc = 0;
    for (const r of p.runs) {
      const m = markOf(r);
      if (m?.kind === 'start' && m.name) starts.push({ name: m.name, id: m.id, pos: { story: 'body', path, offset: acc } });
      else if (m?.kind === 'end') ends.set(m.id, { story: 'body', path, offset: acc });
      acc += r.text.length;
    }
  }
  return starts
    .filter((s) => hidden || !s.name.startsWith('_'))
    .map((s) => {
      const end = ends.get(s.id) ?? null;
      return { name: s.name, start: s.pos, end, text: end ? textBetween(doc, s.pos, end) : '' };
    });
}

/** 书签覆盖的全部文字（REF 域用；跨段时段与段之间用空格连接）；没有这个书签返回 null */
export function bookmarkText(doc: Doc, name: string): string | null {
  const b = listBookmarks(doc, true).find((x) => x.name === name);
  if (!b) return null;
  return b.end ? textBetween(doc, b.start, b.end, Infinity) : '';
}

function textBetween(doc: Doc, a: Pos, b: Pos, max = 40): string {
  if (comparePos(a, b) > 0) return '';
  const paths = paragraphPaths(doc, 'body');
  let out = '';
  for (const path of paths) {
    const here = { story: 'body' as const, path, offset: 0 };
    if (comparePos(here, { ...a, offset: 0 }) < 0 || comparePos(here, { ...b, offset: 0 }) > 0) continue;
    const t = paraText(getPara(doc, here));
    const from = path.join() === a.path.join() ? a.offset : 0;
    const to = path.join() === b.path.join() ? b.offset : t.length;
    out += t.slice(from, to).replaceAll(FIELD_CHAR, '') + (path.join() === b.path.join() ? '' : ' ');
    if (out.length > max) return `${out.slice(0, max)}…`;
  }
  return out;
}

const nextBookmarkId = (doc: Doc): number => {
  let max = 9999;
  for (const { p } of bodyParas(doc)) for (const r of p.runs) {
    const m = markOf(r);
    if (m) max = Math.max(max, Number(m.id) || 0);
  }
  return max + 1;
};

const bookmarkRun = (xml: string, like: RunProps): Run => {
  const opaque: Opaque = { xml, level: 'para', kind: 'bookmark', w: 0, h: 0 };
  return { text: FIELD_CHAR, props: { ...like }, opaque };
};

/** 在选区两端放一对书签标记；同名书签已存在时先删掉（Word：同名即移动） */
export function addBookmark(doc: Doc, a: Pos, b: Pos, name: string): void {
  const bad = checkBookmarkName(name);
  if (bad) throw new Error(bad);
  if (a.story !== 'body' || b.story !== 'body') throw new Error('书签只能加在正文里');
  [a, b] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  removeBookmark(doc, name);
  const id = nextBookmarkId(doc);
  const like = runPropsAt(doc, a);
  // 先放结束标记再放开始标记：选区在同一段时后放的开始标记不会挪动结束标记的位置
  insertMark(doc, b, bookmarkRun(`<w:bookmarkEnd w:id="${id}"/>`, like));
  insertMark(doc, a, bookmarkRun(`<w:bookmarkStart w:id="${id}" w:name="${escAttr(name)}"/>`, like));
}

function insertMark(doc: Doc, pos: Pos, run: Run): void {
  const p = getPara(doc, pos);
  let acc = 0;
  let i = 0;
  for (; i < p.runs.length; i++) {
    const r = p.runs[i];
    if (acc === pos.offset) break;
    if (pos.offset < acc + r.text.length && !isAtom(r)) {
      const k = pos.offset - acc;
      p.runs.splice(i, 1, { ...r, text: r.text.slice(0, k) }, { ...r, props: { ...r.props }, text: r.text.slice(k) });
      i++;
      acc += k;
      break;
    }
    acc += r.text.length;
  }
  // 同一偏移处已有别的零宽标记时，开始标记放在前面、结束标记放在后面
  p.runs.splice(i, 0, run);
}

/** 删除书签（开始和结束标记）；返回有没有删到 */
export function removeBookmark(doc: Doc, name: string): boolean {
  let id: string | null = null;
  for (const { p } of bodyParas(doc)) for (const r of p.runs) {
    const m = markOf(r);
    if (m?.kind === 'start' && m.name === name) id = m.id;
  }
  if (id === null) return false;
  for (const { p } of bodyParas(doc)) {
    const before = p.runs.length;
    p.runs = p.runs.filter((r) => {
      const m = markOf(r);
      return !(m && m.id === id);
    });
    if (p.runs.length !== before && !paraLength(p)) p.runs = p.runs.filter((r) => r.text);
  }
  return true;
}

/** 链接指向的书签是否存在（文档内链接用） */
export const bookmarkExists = (doc: Doc, name: string): boolean => listBookmarks(doc, true).some((b) => b.name === name);

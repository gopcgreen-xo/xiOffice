/** 编辑操作。直接修改传入的文档对象（调用方负责先存撤销快照）。 */
import {
  FIELD_CHAR,
  isAtom,
  hfList,
  hfPath,
  newParagraph,
  newTable,
  pageGeom,
  paraText,
  sameGeom,
  sameProps,
  sectionList,
  textWidth,
  type Align,
  type Block,
  type Doc,
  type FieldKind,
  type HFStory,
  type PageGeom,
  type Paragraph,
  type ParaProps,
  type Run,
  type RunImage,
  type RunProps,
  type SectionBreak,
  type Story,
  type Table,
  type Cell,
  type CellEdge,
  type CellEdges,
} from './model';
import { comparePos, samePath, type Pos } from './layout';
import { isCJK } from './text';
import { ensureList } from './numbering';

type List = (Block | Paragraph)[];

function container(doc: Doc, story: Story, path: number[]): { list: List; index: number } {
  if (story !== 'body') return { list: hfList(doc, story, path), index: path[path.length - 1] };
  if (path.length === 1) return { list: doc.body, index: path[0] };
  const t = doc.body[path[0]] as Table;
  return { list: t.rows[path[1]].cells[path[2]].paragraphs, index: path[3] };
}

function withIndex(path: number[], i: number): number[] {
  return [...path.slice(0, -1), i];
}

export function getPara(doc: Doc, pos: Pick<Pos, 'story' | 'path'>): Paragraph {
  const { list, index } = container(doc, pos.story, pos.path);
  const b = list[index];
  if (!b || b.kind !== 'p') throw new Error('光标位置无效');
  return b;
}

export function paraLength(p: Paragraph): number {
  return p.runs.reduce((n, r) => n + r.text.length, 0);
}

/** 故事中所有段落的路径（文档顺序，含表格单元格内的段落）。页眉页脚取路径 at 所在那一节的分区。 */
export function paragraphPaths(doc: Doc, story: Story, at: number[] = [0]): number[][] {
  if (story !== 'body') return hfList(doc, story, at).map((_, i) => [...at.slice(0, -1), i]);
  const out: number[][] = [];
  doc.body.forEach((b, bi) => {
    if (b.kind === 'p') out.push([bi]);
    else b.rows.forEach((row, r) => row.cells.forEach((c, ci) => c.paragraphs.forEach((_, pi) => out.push([bi, r, ci, pi]))));
  });
  return out;
}

export function sameContainer(a: Pos, b: Pos): boolean {
  return a.story === b.story && a.path.length === b.path.length && samePath(a.path.slice(0, -1), b.path.slice(0, -1));
}

/** 在 offset 处拆分文本段，返回从 offset 开始的第一个文本段下标。 */
function splitAt(p: Paragraph, offset: number): number {
  let acc = 0;
  for (let i = 0; i < p.runs.length; i++) {
    const r = p.runs[i];
    if (offset === acc) return i;
    if (offset < acc + r.text.length && !isAtom(r)) {
      const k = offset - acc;
      p.runs.splice(i, 1, { ...r, text: r.text.slice(0, k) }, { ...r, props: { ...r.props }, text: r.text.slice(k) });
      return i + 1;
    }
    acc += r.text.length;
  }
  return p.runs.length;
}

/** 合并相邻同格式文本段，删除空文本段。 */
function normalize(p: Paragraph) {
  const out: Paragraph['runs'] = [];
  for (const r of p.runs) {
    if (!r.text) continue;
    const last = out[out.length - 1];
    if (last && !isAtom(last) && !isAtom(r) && sameProps(last.props, r.props)) last.text += r.text;
    else out.push(r);
  }
  p.runs = out;
}

export function runPropsAt(doc: Doc, pos: Pos): RunProps {
  const p = getPara(doc, pos);
  let acc = 0;
  for (const r of p.runs) {
    if (pos.offset > acc && pos.offset <= acc + r.text.length) return r.props;
    acc += r.text.length;
  }
  return p.runs[0]?.props ?? p.mark;
}

/** 插入不含段落分隔的文本（可含 '\t' '\n' '\f'）。 */
export function insertRaw(doc: Doc, pos: Pos, s: string, props?: RunProps | null): Pos {
  if (!s) return pos;
  const p = getPara(doc, pos);
  let base = props;
  if (!base) {
    // 沿用光标前那个字的格式（段首用后面的字，空段落用段落标记）
    let acc = 0;
    for (const r of p.runs) {
      const len = r.text.length;
      if (!isAtom(r) && pos.offset > acc && pos.offset <= acc + len) base = r.props;
      acc += len;
      if (base) break;
    }
  }
  const i = splitAt(p, pos.offset);
  base ??= p.runs[i] && !isAtom(p.runs[i]) ? p.runs[i].props : (p.runs[i - 1]?.props ?? p.mark);
  p.runs.splice(i, 0, { text: s, props: trackedProps(base) });
  normalize(p);
  return { ...pos, offset: pos.offset + s.length };
}

// ---------------- 修订 ----------------

/** 修订模式：打开时输入记为插入、删除记为删除（不真的删掉），修订人为 author */
export const tracking = { on: false, author: '工作台用户' };

/** 修订时间到分钟（同一分钟里连续输入的文字合成一段） */
function revDate(): string {
  return `${new Date().toISOString().slice(0, 16)}:00Z`;
}

/** 新输入文字的格式：去掉沿用来的修订标记；修订模式下标为插入 */
function trackedProps(p: RunProps): RunProps {
  const out: RunProps = { ...p };
  delete out.rev;
  if (tracking.on) out.rev = { kind: 'ins', author: tracking.author, date: revDate() };
  return out;
}

/** 修订模式下的删除：区域里的文字标为删除；自己（任何人）还没接受的插入直接去掉。不合并段落 */
function trackedDelete(doc: Doc, s: Pos, e: Pos): Pos {
  for (const { path, from, to } of paragraphsInRange(doc, s, e)) {
    const p = getPara(doc, { story: s.story, path });
    if (from >= to) continue;
    const i = splitAt(p, from);
    const j = splitAt(p, to);
    const kept: Paragraph['runs'] = [];
    for (let k = i; k < j; k++) {
      const r = p.runs[k];
      if (r.mark) kept.push(r);
      else if (r.props.rev?.kind === 'ins') continue;
      else kept.push(r.props.rev?.kind === 'del' ? r : { ...r, props: { ...r.props, rev: { kind: 'del', author: tracking.author, date: revDate() } } });
    }
    p.runs.splice(i, j - i, ...kept);
    normalize(p);
  }
  return s;
}

/** 位置 off 前 / 后那个字是不是已经标为删除 */
function deletedAt(p: Paragraph, off: number): boolean {
  let acc = 0;
  for (const r of p.runs) {
    if (off >= acc && off < acc + r.text.length) return r.props.rev?.kind === 'del';
    acc += r.text.length;
  }
  return false;
}

function trackedBackspace(doc: Doc, pos: Pos): Pos | null {
  const p = getPara(doc, pos);
  let off = pos.offset;
  // 已经标为删除的字跳过去（同 Word）
  while (off > 0 && deletedAt(p, off - 1)) off--;
  if (off === 0) {
    const paths = paragraphPaths(doc, pos.story, pos.path);
    const i = paths.findIndex((x) => samePath(x, pos.path));
    if (i <= 0) return null;
    const prev = getPara(doc, { story: pos.story, path: paths[i - 1] });
    return { story: pos.story, path: paths[i - 1], offset: paraLength(prev) };
  }
  const from = prevOffset(paraText(p), off);
  trackedDelete(doc, { ...pos, offset: from }, { ...pos, offset: off });
  return { ...pos, offset: from };
}

function trackedDeleteForward(doc: Doc, pos: Pos): Pos | null {
  const p = getPara(doc, pos);
  const text = paraText(p);
  let off = pos.offset;
  while (off < text.length && deletedAt(p, off)) off++;
  if (off >= text.length) {
    const paths = paragraphPaths(doc, pos.story, pos.path);
    const i = paths.findIndex((x) => samePath(x, pos.path));
    if (i < 0 || i + 1 >= paths.length) return null;
    return { story: pos.story, path: paths[i + 1], offset: 0 };
  }
  const to = nextOffset(text, off);
  trackedDelete(doc, { ...pos, offset: off }, { ...pos, offset: to });
  return { ...pos, offset: to };
}

/** 所有段落（正文、表格、各节页眉页脚） */
function everyParagraph(doc: Doc): Paragraph[] {
  const out: Paragraph[] = [];
  for (const b of doc.body) {
    if (b.kind === 'p') out.push(b);
    else for (const r of b.rows) for (const c of r.cells) out.push(...c.paragraphs);
  }
  const seen = new Set<Paragraph[]>();
  for (const sec of sectionList(doc)) {
    for (const k of ['header', 'footer', 'headerFirst', 'footerFirst', 'headerEven', 'footerEven'] as const) {
      if (seen.has(sec.hf[k])) continue;
      seen.add(sec.hf[k]);
      out.push(...sec.hf[k]);
    }
  }
  return out;
}

/**
 * 接受（accept = true）或拒绝修订。给了选区时只处理选区里的，否则全部。返回处理的修订段数。
 * 接受插入 / 拒绝删除 = 去掉标记留下文字；接受删除 / 拒绝插入 = 去掉文字。
 */
export function resolveRevisions(doc: Doc, accept: boolean, range?: [Pos, Pos]): number {
  let n = 0;
  const handle = (p: Paragraph, from = 0, to = Infinity) => {
    if (!p.runs.some((r) => r.props.rev)) return;
    const i = from > 0 ? splitAt(p, from) : 0;
    const j = Number.isFinite(to) ? splitAt(p, to) : p.runs.length;
    const out: Paragraph['runs'] = [];
    for (let k = i; k < j; k++) {
      const r = p.runs[k];
      const rev = r.props.rev;
      if (!rev) {
        out.push(r);
        continue;
      }
      n++;
      const keep = (rev.kind === 'ins') === accept;
      if (keep) {
        const props = { ...r.props };
        delete props.rev;
        out.push({ ...r, props });
      }
    }
    p.runs.splice(i, j - i, ...out);
    normalize(p);
  };
  if (range) {
    const [a, b] = range;
    for (const { path, from, to } of paragraphsInRange(doc, a, b)) handle(getPara(doc, { story: a.story, path }), from, to);
  } else for (const p of everyParagraph(doc)) handle(p);
  return n;
}

/** 文档里有没有修订 */
export function hasRevisions(doc: Doc): boolean {
  return everyParagraph(doc).some((p) => p.runs.some((r) => r.props.rev));
}

/** 正文里所有修订（同一作者紧挨着的插入 / 删除算一处），按文档顺序给出起止位置 */
export function allRevisions(doc: Doc): [Pos, Pos][] {
  const found: [Pos, Pos][] = [];
  for (const path of paragraphPaths(doc, 'body')) {
    const p = getPara(doc, { story: 'body', path });
    let acc = 0;
    let start = -1;
    let key = '';
    const flush = (end: number) => {
      if (start >= 0) found.push([{ story: 'body', path, offset: start }, { story: 'body', path, offset: end }]);
      start = -1;
    };
    for (const r of p.runs) {
      // 同一个人紧挨着的删除 + 插入（选中文字直接改写）算一处修订，不分成两处
      const k = r.props.rev ? r.props.rev.author : '';
      if (k !== key) {
        flush(acc);
        if (k) start = acc;
        key = k;
      }
      acc += r.text.length;
    }
    flush(acc);
  }
  return found;
}

/** 正文里 from 之后（back：之前）的下一处修订：返回它的起止位置 */
export function nextRevision(doc: Doc, from: Pos, back = false): [Pos, Pos] | null {
  const found = allRevisions(doc);
  if (!found.length) return null;
  const at = from.story === 'body' ? from : { story: 'body' as const, path: [0], offset: 0 };
  if (back) return [...found].reverse().find((x) => comparePos(x[0], at) < 0) ?? found[found.length - 1];
  return found.find((x) => comparePos(x[0], at) > 0) ?? found[0];
}

/** 包含位置 pos（含两端）的那一处修订；不在修订上返回 null */
export function revisionAt(doc: Doc, pos: Pos): [Pos, Pos] | null {
  if (pos.story !== 'body') return null;
  return allRevisions(doc).find((x) => comparePos(x[0], pos) <= 0 && comparePos(pos, x[1]) <= 0) ?? null;
}

// ---------------- 批注 ----------------

/** 每条批注（有范围标记的）在正文里的起止位置 */
export function commentRanges(doc: Doc): Map<string, [Pos, Pos]> {
  const start = new Map<string, Pos>();
  const out = new Map<string, [Pos, Pos]>();
  for (const path of paragraphPaths(doc, 'body')) {
    let acc = 0;
    for (const r of getPara(doc, { story: 'body', path }).runs) {
      if (r.mark?.kind === 'cs') start.set(r.mark.id, { story: 'body', path, offset: acc + 1 });
      if (r.mark?.kind === 'ce') {
        const s = start.get(r.mark.id);
        if (s) out.set(r.mark.id, [s, { story: 'body', path, offset: acc }]);
      }
      acc += r.text.length;
    }
  }
  return out;
}

const nextCommentId = (doc: Doc) => String(Math.max(-1, ...(doc.comments ?? []).map((c) => Number(c.id) || 0)) + 1);

/** 给选区加批注（只能在正文里）。返回新批注的 id；选区不在正文里返回 null */
export function addComment(doc: Doc, a: Pos, b: Pos, text: string, author: string): string | null {
  const [s, e] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  if (s.story !== 'body' || e.story !== 'body') return null;
  const id = nextCommentId(doc);
  const atom = (kind: 'cs' | 'ce', at: Pos) => {
    const p = getPara(doc, at);
    const i = splitAt(p, at.offset);
    p.runs.splice(i, 0, { text: FIELD_CHAR, props: { ...(p.runs[i]?.props ?? p.runs[i - 1]?.props ?? p.mark) }, mark: { kind, id } });
  };
  // 先插结束标记，开始标记的位置不受影响
  atom('ce', e);
  atom('cs', s);
  (doc.comments ??= []).push({ id, author, date: new Date().toISOString().slice(0, 19) + 'Z', text });
  return id;
}

export function replyComment(doc: Doc, parent: string, text: string, author: string): string {
  const id = nextCommentId(doc);
  (doc.comments ??= []).push({ id, author, date: new Date().toISOString().slice(0, 19) + 'Z', text, parent });
  return id;
}

/** 删除批注和它的回复（回复只删它自己） */
export function deleteComment(doc: Doc, id: string) {
  const c = doc.comments?.find((x) => x.id === id);
  if (!c) return;
  const ids = new Set([id]);
  if (!c.parent) for (const x of doc.comments ?? []) if (x.parent === id) ids.add(x.id);
  doc.comments = (doc.comments ?? []).filter((x) => !ids.has(x.id));
  if (c.parent) return;
  for (const p of everyParagraph(doc)) {
    const before = p.runs.length;
    p.runs = p.runs.filter((r) => !(r.mark && ids.has(r.mark.id)));
    if (p.runs.length !== before) normalize(p);
  }
}

/** 插入文本，'\n' 或 '\r' 视为新段落（粘贴时使用）。 */
export function insertText(doc: Doc, pos: Pos, text: string, props?: RunProps | null): Pos {
  const parts = text.replace(/\r\n?/g, '\n').split('\n');
  let cur = pos;
  parts.forEach((part, i) => {
    if (i > 0) cur = splitParagraph(doc, cur);
    cur = insertRaw(doc, cur, part, props);
  });
  return cur;
}

function deleteInPara(p: Paragraph, from: number, to: number) {
  const i = splitAt(p, from);
  const j = splitAt(p, to);
  const removed = p.runs.splice(i, j - i);
  if (!p.runs.length && removed.length) p.mark = { ...removed[0].props };
  normalize(p);
}

/** 删除选区。replacement 表示随后输入/粘贴：Word 删除整表两端的文字时保留段落边界，替换时合并。 */
export function deleteRange(doc: Doc, a: Pos, b: Pos, replacement = false): Pos | null {
  // 不同故事即使 path 相同也不是同一段；必须在任何修改之前拒绝。
  if (a.story !== b.story) return null;
  const [s, e] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  if (tracking.on && s.story === e.story) return trackedDelete(doc, s, e);
  if (samePath(s.path, e.path)) {
    deleteInPara(getPara(doc, s), s.offset, e.offset);
    return s;
  }
  if (!sameContainer(s, e)) return null;
  const { list, index: si } = container(doc, s.story, s.path);
  const ei = e.path[e.path.length - 1];
  const first = list[si] as Paragraph;
  const last = list[ei] as Paragraph;
  const crossesTable = list.slice(si + 1, ei).some((b) => b.kind === 'table');
  deleteInPara(first, s.offset, paraLength(first));
  deleteInPara(last, 0, e.offset);
  if (crossesTable && !replacement) {
    // Word 16.0 Range.Delete / Selection.Delete: "before [table] after" -> "be\rter".
    // 不合并两端段落，也不把后一段的段落/分节属性移到前一段。
    list.splice(si + 1, ei - si - 1);
    return s;
  }
  first.runs.push(...last.runs);
  normalize(first);
  list.splice(si + 1, ei - si);
  return s;
}

/** 分节符移到另一段时，它那一节的页眉页脚和页面设置随之移动。 */
function moveSectionHF(from: Paragraph, to: Paragraph) {
  if (from === to) return;
  to.sectionHF = from.sectionHF;
  to.sectionPage = from.sectionPage;
  to.sectionCols = from.sectionCols;
  delete from.sectionHF;
  delete from.sectionPage;
  delete from.sectionCols;
  if (!to.sectionHF) delete to.sectionHF;
  if (!to.sectionPage) delete to.sectionPage;
  if (!to.sectionCols) delete to.sectionCols;
}

export function splitParagraph(doc: Doc, pos: Pos): Pos {
  const { list, index } = container(doc, pos.story, pos.path);
  const p = list[index] as Paragraph;
  const i = splitAt(p, pos.offset);
  const tail = p.runs.splice(i);
  const mark = tail[0]?.props ?? p.runs[p.runs.length - 1]?.props ?? p.mark;
  const np: Paragraph = { kind: 'p', props: { ...p.props, pageBreakBefore: false }, runs: tail, mark: { ...mark } };
  // 分节符属于段落标记，随段落标记留在后一段
  p.props = { ...p.props, sectionBreak: null };
  moveSectionHF(p, np);
  if (!p.runs.length) p.mark = { ...mark };
  normalize(p);
  list.splice(index + 1, 0, np);
  return { story: pos.story, path: withIndex(pos.path, index + 1), offset: 0 };
}

function prevOffset(text: string, off: number): number {
  if (off >= 2 && /[\uDC00-\uDFFF]/.test(text[off - 1]) && /[\uD800-\uDBFF]/.test(text[off - 2])) return off - 2;
  return Math.max(0, off - 1);
}

function nextOffset(text: string, off: number): number {
  if (/[\uD800-\uDBFF]/.test(text[off] ?? '')) return Math.min(text.length, off + 2);
  return Math.min(text.length, off + 1);
}

/** Backspace：删除前一个字符，段首时与上一段合并。 */
export function backspace(doc: Doc, pos: Pos): Pos | null {
  if (tracking.on) return trackedBackspace(doc, pos);
  const p = getPara(doc, pos);
  if (pos.offset > 0) {
    const from = prevOffset(paraText(p), pos.offset);
    deleteInPara(p, from, pos.offset);
    return { ...pos, offset: from };
  }
  // 编号段落段首按 Backspace：先去掉编号（与 Word 一致）
  if (p.props.num) {
    p.props = { ...p.props, num: null, indentLeft: 0, firstLine: 0, firstLineChars: null };
    return pos;
  }
  const { list, index } = container(doc, pos.story, pos.path);
  const prev = list[index - 1];
  if (!prev || prev.kind !== 'p') {
    // 段首有缩进时先去掉缩进（与 Word 一致）
    if (p.props.firstLineChars || p.props.firstLine > 0) {
      p.props = { ...p.props, firstLine: 0, firstLineChars: null };
      return pos;
    }
    return null;
  }
  const len = paraLength(prev);
  prev.props = { ...prev.props, sectionBreak: p.props.sectionBreak };
  moveSectionHF(p, prev);
  prev.runs.push(...p.runs);
  if (!prev.runs.length) prev.mark = p.mark;
  normalize(prev);
  list.splice(index, 1);
  return { story: pos.story, path: withIndex(pos.path, index - 1), offset: len };
}

/** Delete：删除后一个字符，段末时合并下一段。 */
export function deleteForward(doc: Doc, pos: Pos): Pos | null {
  if (tracking.on) return trackedDeleteForward(doc, pos);
  const p = getPara(doc, pos);
  const text = paraText(p);
  if (pos.offset < text.length) {
    deleteInPara(p, pos.offset, nextOffset(text, pos.offset));
    return pos;
  }
  const { list, index } = container(doc, pos.story, pos.path);
  const next = list[index + 1];
  if (!next || next.kind !== 'p') return null;
  p.props = { ...p.props, sectionBreak: next.props.sectionBreak };
  moveSectionHF(next, p);
  p.runs.push(...next.runs);
  normalize(p);
  list.splice(index + 1, 1);
  return pos;
}

/** 选区内的段落及各自的起止偏移。 */
export function paragraphsInRange(doc: Doc, a: Pos, b: Pos): { path: number[]; from: number; to: number }[] {
  const [s, e] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  return paragraphPaths(doc, s.story, s.path)
    .filter((path) => comparePos({ ...s, path, offset: 0 }, { ...s, offset: 0 }) >= 0 && comparePos({ ...s, path, offset: 0 }, { ...e, offset: 0 }) <= 0)
    .map((path) => {
      const len = paraLength(getPara(doc, { story: s.story, path }));
      return { path, from: samePath(path, s.path) ? s.offset : 0, to: samePath(path, e.path) ? e.offset : len };
    });
}

export function applyRunProps(doc: Doc, a: Pos, b: Pos, patch: Partial<RunProps>) {
  for (const { path, from, to } of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path });
    const len = paraLength(p);
    if (from < to) {
      const i = splitAt(p, from);
      const j = splitAt(p, to);
      for (let k = i; k < j; k++) p.runs[k].props = { ...p.runs[k].props, ...patch };
    }
    if ((from === 0 && to === len) || len === 0) p.mark = { ...p.mark, ...patch };
    normalize(p);
  }
}

/** 选区里每个文本段的格式按 fn 重新算（不可拆分的原样对象 / 图片也算；选区内整段时段落标记同样处理） */
export function mapRuns(doc: Doc, a: Pos, b: Pos, fn: (props: RunProps, run: Run, para: Paragraph) => RunProps, markToo = true) {
  for (const { path, from, to } of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path });
    const len = paraLength(p);
    if (from < to) {
      const i = splitAt(p, from);
      const j = splitAt(p, to);
      for (let k = i; k < j; k++) p.runs[k].props = fn(p.runs[k].props, p.runs[k], p);
    }
    if (markToo && ((from === 0 && to === len) || len === 0)) p.mark = fn(p.mark, { text: '', props: p.mark }, p);
    normalize(p);
  }
}

export function setParaProps(doc: Doc, a: Pos, b: Pos, patch: Partial<ParaProps>) {
  for (const { path } of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path });
    p.props = { ...p.props, ...patch };
  }
}

/** 对选中段落应用编号预设；preset 为 null 时取消编号。 */
export function applyNumbering(doc: Doc, a: Pos, b: Pos, preset: string | null) {
  const list = preset ? ensureList(doc, preset) : null;
  for (const { path } of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path });
    if (!list) {
      if (p.props.num) p.props = { ...p.props, num: null, indentLeft: 0, firstLine: 0, firstLineChars: null };
      continue;
    }
    const level = p.props.num?.level ?? 0;
    const l = list.levels[level];
    p.props = { ...p.props, num: { id: list.id, level }, indentLeft: l.indentLeft, firstLine: l.firstLine, firstLineChars: l.firstLineChars };
  }
}

/** 编号段落升级（-1）或降级（+1），缩进随级别改变。返回是否有段落被修改。 */
export function changeNumLevel(doc: Doc, a: Pos, b: Pos, delta: 1 | -1): boolean {
  let changed = false;
  for (const { path } of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path });
    const n = p.props.num;
    const list = n && doc.lists.find((l) => l.id === n.id);
    if (!n || !list) continue;
    const level = Math.min(8, Math.max(0, n.level + delta));
    if (level === n.level) continue;
    const l = list.levels[level];
    p.props = { ...p.props, num: { id: n.id, level }, indentLeft: l.indentLeft, firstLine: l.firstLine, firstLineChars: l.firstLineChars };
    changed = true;
  }
  return changed;
}

/**
 * 增加（+1）/ 减少（-1）缩进量：左缩进移到下一个 / 上一个默认制表位（Word 的“增加缩进量”Ctrl+M，按文档的默认制表位间隔），
 * 不小于 0；编号段落由调用方改级别。返回是否有段落被修改。
 */
export function changeIndent(doc: Doc, a: Pos, b: Pos, dir: 1 | -1): boolean {
  const tab = doc.setup.defaultTab > 0 ? doc.setup.defaultTab : 21;
  let changed = false;
  for (const { path } of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path });
    const cur = p.props.indentLeft;
    const k = cur / tab;
    // 刚好在制表位上（容许误差）时走一整格
    const near = Math.round(k);
    const at = Math.abs(k - near) < 0.01 ? near : dir > 0 ? Math.floor(k) : Math.ceil(k);
    const next = Math.max(0, Math.round((at + dir) * tab * 100) / 100);
    if (Math.abs(next - cur) < 0.001) continue;
    p.props = { ...p.props, indentLeft: next };
    changed = true;
  }
  return changed;
}

/** 在光标处插入分节符：拆分段落，前一段成为一节的最后一段。仅用于正文顶层段落。 */
export function insertSectionBreak(doc: Doc, pos: Pos, type: SectionBreak): Pos | null {
  if (pos.story !== 'body' || pos.path.length !== 1) return null;
  const next = splitParagraph(doc, pos);
  const prev = doc.body[pos.path[0]] as Paragraph;
  prev.props = { ...prev.props, sectionBreak: type };
  // 新的一节与原来所在的节共用页眉页脚（“与上一节相同”），取消链接后才各自独立；页面设置也沿用
  const after = sectionList(doc).find((x) => x.index === null || x.index > pos.path[0])!;
  prev.sectionHF = { ...after.hf };
  if (after.index !== null) {
    const page = (doc.body[after.index] as Paragraph).sectionPage;
    if (page) prev.sectionPage = { ...page };
  }
  if (after.setup.cols) prev.sectionCols = structuredClone(after.setup.cols);
  return next;
}

/**
 * 修改页面设置（纸张与页边距）。wholeDoc 时所有节统一；否则只改 index 所结束的那一节（null 为最后一节）。
 * 与 Doc.setup 相同的节不单独保存。
 */
export function setSectionPage(doc: Doc, index: number | null, geom: PageGeom, wholeDoc: boolean) {
  const breaks = doc.body.filter((b): b is Paragraph => b.kind === 'p' && !!b.props.sectionBreak);
  if (wholeDoc) {
    Object.assign(doc.setup, geom);
    for (const b of breaks) delete b.sectionPage;
    return;
  }
  if (index === null) {
    // 其它节原来沿用最后一节的，先记下原来的设置
    const old = pageGeom(doc.setup);
    for (const b of breaks) b.sectionPage ??= { ...old };
    Object.assign(doc.setup, geom);
  } else {
    (doc.body[index] as Paragraph).sectionPage = { ...geom };
  }
  for (const b of breaks) if (b.sectionPage && sameGeom(b.sectionPage, doc.setup)) delete b.sectionPage;
}

/** 各节都保存自己的页眉页脚（沿用下一节的分节段落补上共用的数组），供修改链接关系前调用。 */
function ensureSectionHF(doc: Doc) {
  let next = sectionList(doc).at(-1)!.hf;
  for (let i = doc.body.length - 1; i >= 0; i--) {
    const b = doc.body[i];
    if (b.kind !== 'p' || !b.props.sectionBreak) continue;
    b.sectionHF ??= { ...next };
    next = b.sectionHF;
  }
}

/** 页眉页脚分区 story 在路径 at 所在的节是否“与上一节相同”；第一节返回 null。 */
export function hfLinked(doc: Doc, story: HFStory, at: number[]): boolean | null {
  const list = sectionList(doc);
  const owner = at.length < 2 ? null : at[0];
  const i = list.findIndex((x) => x.owner === owner);
  if (i <= 0) return null;
  return list[i].hf[story] === list[i - 1].hf[story];
}

/**
 * 设置“与上一节相同”：链接时本节（以及原来与本节相同的后续各节）改用上一节的内容；
 * 取消链接时复制一份，本节和后续相同的节从此独立。返回光标在该分区中的新位置。
 */
export function setHFLinked(doc: Doc, story: HFStory, at: number[], linked: boolean): Pos {
  ensureSectionHF(doc);
  const list = sectionList(doc);
  const owner = at.length < 2 ? null : at[0];
  const i = list.findIndex((x) => x.owner === owner);
  const old = list[i].hf[story];
  if (i > 0 && (list[i - 1].hf[story] === old) !== linked) {
    const arr = linked ? list[i - 1].hf[story] : structuredClone(old);
    for (let j = i; j < list.length && list[j].hf[story] === old; j++) {
      if (list[j].owner === null) doc[story] = arr;
      else list[j].hf[story] = arr;
    }
  }
  return { story, path: hfPath(owner, 0), offset: 0 };
}

/** 插入嵌入式图片（随文字排版），返回图片后面的位置 */
export function insertImage(doc: Doc, pos: Pos, image: RunImage): Pos {
  const p = getPara(doc, pos);
  const props = trackedProps(runPropsAt(doc, pos));
  const i = splitAt(p, pos.offset);
  p.runs.splice(i, 0, { text: FIELD_CHAR, props, image });
  return { ...pos, offset: pos.offset + 1 };
}

export function insertField(doc: Doc, pos: Pos, kind: FieldKind): Pos {
  const p = getPara(doc, pos);
  const props = runPropsAt(doc, pos);
  const i = splitAt(p, pos.offset);
  p.runs.splice(i, 0, { text: FIELD_CHAR, props: { ...props }, field: kind });
  return { ...pos, offset: pos.offset + 1 };
}

// ---------------- 表格 ----------------

export function tableAt(doc: Doc, pos: Pos): { table: Table; bi: number; r: number; c: number } | null {
  if (pos.story !== 'body' || pos.path.length !== 4) return null;
  return { table: doc.body[pos.path[0]] as Table, bi: pos.path[0], r: pos.path[1], c: pos.path[2] };
}

export function insertTable(doc: Doc, pos: Pos, rows: number, cols: number): Pos | null {
  if (pos.story !== 'body' || pos.path.length !== 1) return null;
  const p = getPara(doc, pos);
  const t = newTable(cols, rows, textWidth(doc.setup), { ...runPropsAt(doc, pos) });
  // 空段落或段首：插在该段之前；否则在光标处拆段，表格插在两段之间（段末时后面留一个空段）
  let at = pos.path[0];
  if (paraLength(p) > 0 && pos.offset > 0) at = splitParagraph(doc, pos).path[0];
  doc.body.splice(at, 0, t);
  return { story: 'body', path: [at, 0, 0, 0], offset: 0 };
}

function emptyLike(p: Paragraph): Paragraph {
  return { kind: 'p', props: { ...p.props }, runs: [], mark: { ...(p.runs[0]?.props ?? p.mark) } };
}

export function insertRow(doc: Doc, pos: Pos, below: boolean): Pos | null {
  const t = tableAt(doc, pos);
  if (!t) return null;
  // 按光标所在合并格的上 / 下边插入
  const rect = cellRect(t.table, t.r, t.c);
  const src = below ? rect.r1 : rect.r0;
  const at = below ? rect.r1 + 1 : rect.r0;
  const row = t.table.rows[src];
  const starts = gridStarts(t.table);
  // 插入处正好穿过一个纵向合并格（下面一行是它的 continue）时，新行里那一格也接着合并
  const after = t.table.rows[at];
  const nr = {
    ...row,
    header: false,
    cells: row.cells.map((c, i) => {
      const k = after ? cellAtGrid(t.table, starts, at, starts[src][i]) : -1;
      const through = k >= 0 && after.cells[k].vMerge === 'continue' && after.cells[k].gridSpan === c.gridSpan && starts[at][k] === starts[src][i];
      return { ...c, vMerge: through ? ('continue' as const) : ('none' as const), paragraphs: [emptyLike(c.paragraphs[0])] };
    }),
  };
  t.table.rows.splice(at, 0, nr);
  return { story: 'body', path: [t.bi, at, Math.min(t.c, nr.cells.length - 1), 0], offset: 0 };
}

export function deleteTable(doc: Doc, bi: number): Pos {
  doc.body.splice(bi, 1);
  if (!doc.body[bi] || doc.body[bi].kind !== 'p') doc.body.splice(bi, 0, newParagraph('', undefined, doc.defaults));
  return { story: 'body', path: [bi], offset: 0 };
}

export function deleteRow(doc: Doc, pos: Pos): Pos | null {
  const t = tableAt(doc, pos);
  if (!t) return null;
  if (t.table.rows.length === 1) return deleteTable(doc, t.bi);
  // 删除的是纵向合并的第一行：下一行接手合并（内容一并移下去）
  const starts = gridStarts(t.table);
  t.table.rows[t.r].cells.forEach((c, i) => {
    if (c.vMerge !== 'restart' || t.r + 1 >= t.table.rows.length) return;
    const k = cellAtGrid(t.table, starts, t.r + 1, starts[t.r][i]);
    const next = k >= 0 ? t.table.rows[t.r + 1].cells[k] : null;
    if (!next || next.vMerge !== 'continue') return;
    const k2 = t.r + 2 < t.table.rows.length ? cellAtGrid(t.table, starts, t.r + 2, starts[t.r][i]) : -1;
    next.vMerge = k2 >= 0 && t.table.rows[t.r + 2].cells[k2].vMerge === 'continue' ? 'restart' : 'none';
    next.paragraphs = c.paragraphs;
  });
  t.table.rows.splice(t.r, 1);
  const r = Math.min(t.r, t.table.rows.length - 1);
  return { story: 'body', path: [t.bi, r, Math.min(t.c, t.table.rows[r].cells.length - 1), 0], offset: 0 };
}

/**
 * 在当前格右侧（left = 左侧）插入一列：当前格最右（最左）的网格列一分为二。各行在这条线上结束（开始）的格子后面（前面）
 * 插一个新格（纵向合并随旁边的格子），横跨这条线的格子跨列数加 1。
 */
export function insertColumn(doc: Doc, pos: Pos, left = false): Pos | null {
  const t = tableAt(doc, pos);
  if (!t) return null;
  const tb = t.table;
  const starts = gridStarts(tb);
  const G = left ? starts[t.r][t.c] : starts[t.r][t.c] + tb.rows[t.r].cells[t.c].gridSpan - 1;
  const w = tb.cols[G] / 2;
  tb.cols.splice(G, 1, w, w);
  tb.rows.forEach((row, r) => {
    const i = cellAtGrid(tb, starts, r, G);
    if (i < 0) return;
    const c = row.cells[i];
    const fresh = { ...emptyCellLike(c), vMerge: c.vMerge };
    if (left) {
      if (starts[r][i] < G) c.gridSpan++;
      else row.cells.splice(i, 0, fresh);
    } else if (starts[r][i] + c.gridSpan - 1 > G) c.gridSpan++;
    else row.cells.splice(i + 1, 0, fresh);
  });
  return { story: 'body', path: [t.bi, t.r, left ? t.c : t.c + 1, 0], offset: 0 };
}

/** 删除当前格所占的网格列：完全落在其中的格子删掉，部分重叠的格子减少跨列数；删掉的宽度分给相邻列，保持表格总宽 */
export function deleteColumn(doc: Doc, pos: Pos): Pos | null {
  const t = tableAt(doc, pos);
  if (!t) return null;
  const tb = t.table;
  const starts = gridStarts(tb);
  const g0 = starts[t.r][t.c];
  const g1 = g0 + tb.rows[t.r].cells[t.c].gridSpan - 1;
  if (g1 - g0 + 1 >= tb.cols.length) return deleteTable(doc, t.bi);
  const w = tb.cols.slice(g0, g1 + 1).reduce((a, b) => a + b, 0);
  tb.rows.forEach((row, r) => {
    row.cells = row.cells.filter((c, i) => {
      const s = starts[r][i];
      const e = s + c.gridSpan - 1;
      const overlap = Math.max(0, Math.min(e, g1) - Math.max(s, g0) + 1);
      c.gridSpan -= overlap;
      return c.gridSpan > 0;
    });
  });
  tb.cols.splice(g0, g1 - g0 + 1);
  tb.cols[Math.min(g0, tb.cols.length - 1)] += w;
  const row = tb.rows[t.r].cells;
  return { story: 'body', path: [t.bi, t.r, Math.max(0, Math.min(t.c, row.length - 1)), 0], offset: 0 };
}

// ---------------- 光标移动 ----------------

export function moveHorizontal(doc: Doc, pos: Pos, dir: 1 | -1): Pos {
  const p = getPara(doc, pos);
  const text = paraText(p);
  if (dir < 0 && pos.offset > 0) return { ...pos, offset: prevOffset(text, pos.offset) };
  if (dir > 0 && pos.offset < text.length) return { ...pos, offset: nextOffset(text, pos.offset) };
  const paths = paragraphPaths(doc, pos.story, pos.path);
  const i = paths.findIndex((x) => samePath(x, pos.path));
  const j = i + dir;
  if (j < 0 || j >= paths.length) return pos;
  const np = getPara(doc, { story: pos.story, path: paths[j] });
  return { story: pos.story, path: paths[j], offset: dir > 0 ? 0 : paraLength(np) };
}

/** 故事开头 / 结尾；页眉页脚取路径 at 所在那一节的分区 */
export function storyEdge(doc: Doc, story: Story, end: boolean, at?: number[]): Pos {
  const paths = paragraphPaths(doc, story, at);
  const path = end ? paths[paths.length - 1] : paths[0];
  return { story, path, offset: end ? paraLength(getPara(doc, { story, path })) : 0 };
}

/**
 * 写回 AI 结果（工作台的 AI 动作）：replace = 替换选区，新文字用选区第一个字的格式（与查找替换相同）；
 * insert-after = 在选区末尾另起一段插入，格式随末尾。返回插入后文字末尾的位置；选区跨容器时返回 null。
 */
export function applyAiText(doc: Doc, a: Pos, b: Pos, text: string, mode: 'replace' | 'insert-after'): Pos | null {
  const [start, end] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  if (start.story !== end.story) return null;
  if (mode === 'insert-after') return insertText(doc, end, `\n${text}`, runPropsAt(doc, end));
  const props = { ...runPropsAt(doc, { ...start, offset: start.offset + 1 }) };
  const at = comparePos(start, end) === 0 ? start : deleteRange(doc, start, end, true);
  if (!at) return null;
  return insertText(doc, at, text, props);
}

export function selectedText(doc: Doc, a: Pos, b: Pos): string {
  return paragraphsInRange(doc, a, b)
    .map(({ path, from, to }) => paraText(getPara(doc, { story: a.story, path })).slice(from, to).replace(new RegExp(FIELD_CHAR, 'g'), ''))
    .join('\r\n');
}

/** 字数统计（与 Word 相同口径：中文字符与西文单词数之和）。 */
export function wordCount(doc: Doc): number {
  let n = 0;
  for (const path of paragraphPaths(doc, 'body')) {
    const t = paraText(getPara(doc, { story: 'body', path }));
    for (const ch of t) if (isCJK(ch) && !/[，。、；：？！“”‘’（）《》【】—…·]/.test(ch)) n++;
    n += (t.match(/[A-Za-z0-9À-ɏ]+/g) ?? []).length;
  }
  return n;
}

// ---------------- 撤销 ----------------

export interface Snapshot<S> {
  doc: Doc;
  sel: S;
}

export class History<S> {
  /** 清空（设置 / 停止文档保护后不能再撤销到保护之前） */
  clear() {
    this.undoStack = [];
    this.redoStack = [];
    this.lastKind = '';
  }
  private undoStack: Snapshot<S>[] = [];
  private redoStack: Snapshot<S>[] = [];
  private lastKind = '';
  private lastTime = 0;

  /** 修改前调用。连续输入（1 秒内）合并为一步。 */
  record(doc: Doc, sel: S, kind: string) {
    const now = Date.now();
    if (kind === 'type' && this.lastKind === 'type' && now - this.lastTime < 1000) {
      this.lastTime = now;
      return;
    }
    this.undoStack.push({ doc: structuredClone(doc), sel: structuredClone(sel) });
    if (this.undoStack.length > 200) this.undoStack.shift();
    this.redoStack = [];
    this.lastKind = kind;
    this.lastTime = now;
  }

  /** 撤回刚记下的一步（修改失败时用）：返回那一步的快照，不进重做栈 */
  dropLast(): Snapshot<S> | null {
    this.lastKind = '';
    return this.undoStack.pop() ?? null;
  }

  undo(doc: Doc, sel: S): Snapshot<S> | null {
    const s = this.undoStack.pop();
    if (!s) return null;
    this.redoStack.push({ doc: structuredClone(doc), sel: structuredClone(sel) });
    this.lastKind = '';
    return s;
  }

  redo(doc: Doc, sel: S): Snapshot<S> | null {
    const s = this.redoStack.pop();
    if (!s) return null;
    this.undoStack.push({ doc: structuredClone(doc), sel: structuredClone(sel) });
    this.lastKind = '';
    return s;
  }

  get canUndo() {
    return this.undoStack.length > 0;
  }
  get canRedo() {
    return this.redoStack.length > 0;
  }
}

// ---------------- 复制粘贴（带格式） ----------------

/** 段落里 [from, to) 这一段（深拷贝，段落格式照搬，不带分节符） */
function slicePara(p: Paragraph, from: number, to: number): Paragraph {
  const runs: Paragraph['runs'] = [];
  let acc = 0;
  for (const r of p.runs) {
    const a = Math.max(from, acc);
    const b = Math.min(to, acc + r.text.length);
    if (a < b) runs.push(isAtom(r) ? structuredClone(r) : { ...structuredClone(r), text: r.text.slice(a - acc, b - acc) });
    acc += r.text.length;
  }
  return { kind: 'p', props: { ...p.props, sectionBreak: null }, runs, mark: { ...(runs[0]?.props ?? p.mark) } };
}

/**
 * 选区内容 → 一组块（段落、表格），用于复制。正文里选区跨过整张表格时带上整张表格；
 * 选区一端在表格里时也带上那张表格；表格内部、页眉页脚里只取段落。
 */
export function copyFragment(doc: Doc, a: Pos, b: Pos): Block[] {
  const [s, e] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  const out: Block[] = [];
  if (s.story !== 'body' || (s.path.length > 1 && e.path.length > 1 && s.path[0] === e.path[0])) {
    for (const { path, from, to } of paragraphsInRange(doc, s, e)) out.push(slicePara(getPara(doc, { story: s.story, path }), from, to));
    return out;
  }
  for (let bi = s.path[0]; bi <= e.path[0]; bi++) {
    const blk = doc.body[bi];
    if (blk.kind === 'table') {
      out.push(structuredClone(blk));
      continue;
    }
    const from = bi === s.path[0] && s.path.length === 1 ? s.offset : 0;
    const to = bi === e.path[0] && e.path.length === 1 ? e.offset : paraLength(blk);
    out.push(slicePara(blk, from, to));
  }
  // 复制出去的块不再对应原文件里的元素（粘贴到别处、别的文档时不能当成原来那一块）
  for (const b of out) delete b.src;
  return out;
}

/** 表格 → 段落（每行一段，单元格之间用制表符），用于不能放表格的位置 */
function flattenTable(t: Table): Paragraph[] {
  return t.rows.map((row) => {
    const runs: Paragraph['runs'] = [];
    row.cells.forEach((c, i) => {
      if (i) runs.push({ text: '\t', props: { ...(c.paragraphs[0]?.runs[0]?.props ?? c.paragraphs[0].mark) } });
      c.paragraphs.forEach((p, j) => {
        if (j) runs.push({ text: '\n', props: { ...p.mark } });
        runs.push(...structuredClone(p.runs));
      });
    });
    const first = row.cells[0]?.paragraphs[0];
    return { kind: 'p', props: { ...(first?.props ?? newParagraph().props), sectionBreak: null }, runs, mark: { ...(first?.mark ?? newParagraph().mark) } };
  });
}

/**
 * 在 pos 处插入一组块（粘贴）：第一段接在光标前的文字后面，最后一段接在光标后的文字前面，中间的段落和表格整块插入。
 * 表格只能放在正文顶层；在表格里、页眉页脚里时表格转成制表符分隔的段落。返回粘贴内容末尾的位置。
 */
export function insertFragment(doc: Doc, pos: Pos, blocks: Block[]): Pos {
  const topLevel = pos.story === 'body' && pos.path.length === 1;
  const items: Block[] = blocks.flatMap((b) => (b.kind === 'table' && !topLevel ? flattenTable(b) : [structuredClone(b)]));
  if (!items.length) return pos;
  // 粘贴进来的文字：去掉原来的修订标记；修订模式下标为插入
  const mark = (ps: Paragraph[]) => ps.forEach((p) => p.runs.forEach((r) => r.mark || (r.props = trackedProps(r.props))));
  for (const b of items) mark(b.kind === 'p' ? [b] : b.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs)));
  // 只有一段：插入它的文字（带格式）
  if (items.length === 1 && items[0].kind === 'p') {
    const p = getPara(doc, pos);
    const i = splitAt(p, pos.offset);
    p.runs.splice(i, 0, ...items[0].runs);
    normalize(p);
    return { ...pos, offset: pos.offset + paraLength(items[0]) };
  }
  const tail = splitParagraph(doc, pos);
  const { list } = container(doc, pos.story, pos.path);
  const head = list[pos.path[pos.path.length - 1]] as Paragraph;
  const after = list[tail.path[tail.path.length - 1]] as Paragraph;
  let mid = items;
  if (mid[0].kind === 'p') {
    head.runs.push(...mid[0].runs);
    normalize(head);
    mid = mid.slice(1);
  }
  let endLen = 0;
  const last = mid[mid.length - 1];
  if (last && last.kind === 'p') {
    after.runs.unshift(...last.runs);
    after.props = { ...last.props, sectionBreak: after.props.sectionBreak };
    endLen = paraLength(last);
    normalize(after);
    mid = mid.slice(0, -1);
  }
  const at = tail.path[tail.path.length - 1];
  list.splice(at, 0, ...mid);
  return { story: pos.story, path: withIndex(tail.path, at + mid.length), offset: endLen };
}

// ---------------- 表格：合并、拆分、底纹、列宽（按网格列计算，支持合并单元格） ----------------

/** 每行每个单元格的起始网格列 */
function gridStarts(t: Table): number[][] {
  return t.rows.map((row) => {
    let g = 0;
    return row.cells.map((c) => {
      const s = g;
      g += c.gridSpan;
      return s;
    });
  });
}

/** 第 r 行覆盖网格列 g 的单元格下标 */
function cellAtGrid(t: Table, starts: number[][], r: number, g: number): number {
  return t.rows[r].cells.findIndex((c, i) => starts[r][i] <= g && g < starts[r][i] + c.gridSpan);
}

const emptyCellLike = (c: Cell): Cell => ({
  gridSpan: 1,
  vMerge: 'none',
  vAlign: c.vAlign,
  ...(c.borders !== undefined && { borders: c.borders }),
  ...(c.bdr && { bdr: { ...c.bdr } }),
  ...(c.edges && { edges: { ...c.edges } }),
  ...(c.shading && { shading: c.shading }),
  paragraphs: [emptyLike(c.paragraphs[0])],
});

/** 单元格在网格里占的矩形（含向下合并的行） */
export function cellRect(t: Table, r: number, ci: number): { r0: number; r1: number; g0: number; g1: number } {
  const starts = gridStarts(t);
  const g0 = starts[r][ci];
  const g1 = g0 + t.rows[r].cells[ci].gridSpan - 1;
  let r0 = r;
  while (r0 > 0 && t.rows[r0].cells[ci]?.vMerge === 'continue') {
    const up = cellAtGrid(t, starts, r0 - 1, g0);
    if (up < 0) break;
    r0--;
    ci = up;
  }
  let r1 = r;
  for (;;) {
    const k = r1 + 1 < t.rows.length ? cellAtGrid(t, starts, r1 + 1, g0) : -1;
    if (k < 0 || t.rows[r1 + 1].cells[k].vMerge !== 'continue') break;
    r1++;
  }
  return { r0, r1, g0, g1 };
}

/** Delete on a rectangular cell selection clears contents and preserves table geometry. */
export function clearCellContents(doc: Doc, a: Pos, b: Pos): Pos | null {
  const ta = tableAt(doc, a);
  const tb = tableAt(doc, b);
  if (!ta || !tb || ta.bi !== tb.bi) return null;
  const t = ta.table;
  const ra = cellRect(t, ta.r, ta.c);
  const rb = cellRect(t, tb.r, tb.c);
  const r0 = Math.min(ra.r0, rb.r0), r1 = Math.max(ra.r1, rb.r1);
  const g0 = Math.min(ra.g0, rb.g0), g1 = Math.max(ra.g1, rb.g1);
  const starts = gridStarts(t);
  let first: Pos | null = null;
  // Resolve merged cells to their owner so a selection through a continuation
  // clears the visible contents without changing any gridSpan/vMerge values.
  const owners = new Set<Cell>();
  for (let r = r0; r <= r1; r++) {
    t.rows[r].cells.forEach((c, ci) => {
      if (starts[r][ci] > g1 || starts[r][ci] + c.gridSpan - 1 < g0) return;
      const rect = cellRect(t, r, ci);
      const ownerIndex = cellAtGrid(t, starts, rect.r0, rect.g0);
      const owner = t.rows[rect.r0].cells[ownerIndex];
      if (owners.has(owner)) return;
      owners.add(owner);
      first ??= { story: 'body', path: [ta.bi, rect.r0, ownerIndex, 0], offset: 0 };
      if (tracking.on) {
        owner.paragraphs.forEach((p, pi) => {
          const at: Pos = { story: 'body', path: [ta.bi, rect.r0, ownerIndex, pi], offset: 0 };
          trackedDelete(doc, at, { ...at, offset: paraLength(p) });
        });
      } else owner.paragraphs = [emptyLike(owner.paragraphs[0])];
    });
  }
  return first;
}

/**
 * 合并单元格：两个位置（同一张表里的单元格）围成的矩形；矩形边上不能切开别的合并单元格。
 * 内容按顺序并到左上角的格子里。成功返回左上角的位置。
 */
export function mergeTableCells(doc: Doc, a: Pos, b: Pos): Pos | null {
  const ta = tableAt(doc, a);
  const tb = tableAt(doc, b);
  if (!ta || !tb || ta.bi !== tb.bi) return null;
  const t = ta.table;
  const ra = cellRect(t, ta.r, ta.c);
  const rb = cellRect(t, tb.r, tb.c);
  const R0 = Math.min(ra.r0, rb.r0);
  const R1 = Math.max(ra.r1, rb.r1);
  const G0 = Math.min(ra.g0, rb.g0);
  const G1 = Math.max(ra.g1, rb.g1);
  if (R0 === R1 && G0 === G1) return null;
  const starts = gridStarts(t);
  // 矩形边上的单元格必须完整落在矩形里
  for (let r = R0; r <= R1; r++) {
    for (let i = 0; i < t.rows[r].cells.length; i++) {
      const s = starts[r][i];
      const e = s + t.rows[r].cells[i].gridSpan - 1;
      if (e < G0 || s > G1) continue;
      if (s < G0 || e > G1) return null;
      const rc = cellRect(t, r, i);
      if (rc.r0 < R0 || rc.r1 > R1) return null;
    }
  }
  const paras: Paragraph[] = [];
  for (let r = R0; r <= R1; r++) {
    const keep: Cell[] = [];
    t.rows[r].cells.forEach((c, i) => {
      const s = starts[r][i];
      if (s < G0 || s > G1) return keep.push(c);
      if (c.vMerge !== 'continue') paras.push(...c.paragraphs.filter((p) => paraLength(p) > 0));
      if (s === G0) keep.push({ ...c, gridSpan: G1 - G0 + 1, vMerge: R1 > R0 ? (r === R0 ? 'restart' : 'continue') : 'none', paragraphs: r === R0 ? [] : [emptyLike(c.paragraphs[0])] });
    });
    t.rows[r].cells = keep;
  }
  const top = t.rows[R0].cells[cellAtGrid(t, gridStarts(t), R0, G0)];
  top.paragraphs = paras.length ? paras : [emptyLike(t.rows[R0].cells[0].paragraphs[0] ?? newParagraph())];
  return { story: 'body', path: [ta.bi, R0, cellAtGrid(t, gridStarts(t), R0, G0), 0], offset: 0 };
}

/** 取消合并：光标所在的合并单元格拆回一格一格（内容留在左上角） */
export function unmergeCell(doc: Doc, pos: Pos): Pos | null {
  const at = tableAt(doc, pos);
  if (!at) return null;
  const t = at.table;
  const rc = cellRect(t, at.r, at.c);
  if (rc.r0 === rc.r1 && rc.g0 === rc.g1) return null;
  for (let r = rc.r0; r <= rc.r1; r++) {
    const starts = gridStarts(t);
    const i = cellAtGrid(t, starts, r, rc.g0);
    const c = t.rows[r].cells[i];
    const parts: Cell[] = [];
    for (let g = rc.g0; g <= rc.g1; g++) parts.push(g === rc.g0 && r === rc.r0 ? { ...c, gridSpan: 1, vMerge: 'none' } : emptyCellLike(c));
    t.rows[r].cells.splice(i, 1, ...parts);
  }
  return { story: 'body', path: [at.bi, rc.r0, cellAtGrid(t, gridStarts(t), rc.r0, rc.g0), 0], offset: 0 };
}

/** 拆分单元格为 n 列：所在的网格列平分成 n 份，同一网格列上其它行的格子跨列数随之增加 */
export function splitCell(doc: Doc, pos: Pos, n: number): Pos | null {
  const at = tableAt(doc, pos);
  if (!at || n < 2) return null;
  const t = at.table;
  const starts = gridStarts(t);
  const c = t.rows[at.r].cells[at.c];
  if (c.gridSpan > 1 || c.vMerge !== 'none') return null;
  const g = starts[at.r][at.c];
  const w = t.cols[g] / n;
  t.cols.splice(g, 1, ...Array(n).fill(w));
  t.rows.forEach((row, r) => {
    const i = cellAtGrid(t, starts, r, g);
    if (i < 0) return;
    if (r === at.r) row.cells.splice(i, 1, row.cells[i], ...Array.from({ length: n - 1 }, () => emptyCellLike(row.cells[i])));
    else row.cells[i] = { ...row.cells[i], gridSpan: row.cells[i].gridSpan + n - 1 };
  });
  return pos;
}

/**
 * 单元格底纹（'RRGGBB'，null 去掉）、框线（true / false，null 跟随表格）、垂直对齐与格子里各段的水平对齐：
 * 作用于两个位置围成的矩形里的格子
 */
export function setCellFormat(doc: Doc, a: Pos, b: Pos, patch: { shading?: string | null; borders?: boolean | null; vAlign?: Cell['vAlign']; align?: Align }): boolean {
  const ta = tableAt(doc, a);
  const tb = tableAt(doc, b);
  if (!ta || !tb || ta.bi !== tb.bi) return false;
  const t = ta.table;
  const starts = gridStarts(t);
  const R0 = Math.min(ta.r, tb.r);
  const R1 = Math.max(ta.r, tb.r);
  const G0 = Math.min(starts[ta.r][ta.c], starts[tb.r][tb.c]);
  const G1 = Math.max(starts[ta.r][ta.c], starts[tb.r][tb.c]);
  for (let r = R0; r <= R1; r++) {
    t.rows[r].cells.forEach((c, i) => {
      const s = starts[r][i];
      if (s + c.gridSpan - 1 < G0 || s > G1) return;
      if (patch.shading !== undefined) {
        if (patch.shading) c.shading = patch.shading;
        else delete c.shading;
      }
      if (patch.borders !== undefined) {
        // 改了边框：不再按文件里合成的四条边画
        delete c.edges;
        delete c.bdr;
        if (patch.borders === null) delete c.borders;
        else c.borders = patch.borders;
      }
      if (patch.vAlign) c.vAlign = patch.vAlign;
      if (patch.align) for (const p of c.paragraphs) p.props = { ...p.props, align: patch.align };
    });
  }
  return true;
}

/** 框线应用到选中格子的哪些边（同 Word“边框”下拉：所有框线、外侧框线、内部框线…） */
export type BorderTarget = 'all' | 'outside' | 'inside' | 'insideH' | 'insideV' | 'top' | 'bottom' | 'left' | 'right' | 'none';

type Side = 'top' | 'bottom' | 'left' | 'right';
const OPPOSITE: Record<Side, Side> = { top: 'bottom', bottom: 'top', left: 'right', right: 'left' };

/** 格子现在画出来的四条边（有 edges 按 edges，否则按单元格 / 表格的有无框线） */
function effectiveEdges(t: Table, c: Cell): CellEdges {
  if (c.edges) return { ...c.edges };
  const on = c.borders ?? t.borders;
  const e = on ? { w: t.borderSize ?? 0.5, color: 'auto' } : null;
  return { top: e, bottom: e, left: e, right: e };
}

/** 设一条边：同时记进单元格自己的框线（保存成 tcBorders）和显示用的四条边 */
function setEdge(t: Table, c: Cell, side: Side, e: CellEdge | null) {
  c.edges = { ...effectiveEdges(t, c), [side]: e };
  c.bdr = { ...c.bdr, [side]: e };
}

/**
 * 给两个位置围成的矩形里的格子设框线（line = null 为无框线）。外侧的边同时改相邻格子对着的那条边，
 * 两边一致（同 Word：两格公共边的设置写在两个格子上）。返回是否在表格里。
 */
export function applyCellBorders(doc: Doc, a: Pos, b: Pos, target: BorderTarget, line: CellEdge | null): boolean {
  const ta = tableAt(doc, a);
  const tb = tableAt(doc, b);
  if (!ta || !tb || ta.bi !== tb.bi) return false;
  const t = ta.table;
  const starts = gridStarts(t);
  const R0 = Math.min(ta.r, tb.r);
  const R1 = Math.max(ta.r, tb.r);
  const G0 = Math.min(starts[ta.r][ta.c], starts[tb.r][tb.c]);
  const G1 = Math.max(starts[ta.r][ta.c] + t.rows[ta.r].cells[ta.c].gridSpan - 1, starts[tb.r][tb.c] + t.rows[tb.r].cells[tb.c].gridSpan - 1);
  const value = target === 'none' ? null : line;
  const wants = (side: Side, outer: boolean): boolean => {
    switch (target) {
      case 'all':
      case 'none':
        return true;
      case 'outside':
        return outer;
      case 'inside':
        return !outer;
      case 'insideH':
        return !outer && (side === 'top' || side === 'bottom');
      case 'insideV':
        return !outer && (side === 'left' || side === 'right');
      default:
        return outer && side === target;
    }
  };
  /** 第 r 行与网格列 [g0, g1] 有重叠的格子 */
  const overlapping = (r: number, g0: number, g1: number) =>
    t.rows[r]?.cells.filter((c, i) => starts[r][i] <= g1 && starts[r][i] + c.gridSpan - 1 >= g0) ?? [];
  for (let r = R0; r <= R1; r++) {
    t.rows[r].cells.forEach((c, i) => {
      const g0 = starts[r][i];
      const g1 = g0 + c.gridSpan - 1;
      if (g1 < G0 || g0 > G1) return;
      const outer: Record<Side, boolean> = { top: r === R0, bottom: r === R1, left: g0 <= G0, right: g1 >= G1 };
      for (const side of ['top', 'bottom', 'left', 'right'] as const) {
        if (!wants(side, outer[side])) continue;
        setEdge(t, c, side, value && { ...value });
        // 外侧的边：相邻格子（矩形外面）对着的那条边一起改；里面的边对面的格子本身在矩形里，循环会改到
        if (!outer[side]) continue;
        const nb =
          side === 'top'
            ? overlapping(r - 1, g0, g1)
            : side === 'bottom'
              ? overlapping(r + 1, g0, g1)
              : side === 'left'
                ? t.rows[r].cells.filter((_, k) => starts[r][k] + t.rows[r].cells[k].gridSpan === g0)
                : t.rows[r].cells.filter((_, k) => starts[r][k] === g1 + 1);
        for (const n of nb) setEdge(t, n, OPPOSITE[side], value && { ...value });
      }
    });
  }
  return true;
}

/** 拖动列线：第 g 列的右边线移到 dx 之后；右边还有列时两列此消彼长（表格总宽不变），最后一列直接变宽 */
export function moveColumnLine(t: Table, g: number, dx: number) {
  const min = 12;
  if (g + 1 < t.cols.length) {
    const d = Math.max(min - t.cols[g], Math.min(dx, t.cols[g + 1] - min));
    t.cols[g] += d;
    t.cols[g + 1] -= d;
  } else t.cols[g] = Math.max(min, t.cols[g] + dx);
}

/** 平均分布各列 */
export function distributeColumns(t: Table) {
  const total = t.cols.reduce((a, b) => a + b, 0);
  t.cols = t.cols.map(() => total / t.cols.length);
}

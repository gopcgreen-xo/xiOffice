/**
 * 样式管理：应用段落 / 字符样式、从选中内容新建样式、用选中内容更新样式（并让用到它的段落跟着变）。
 *
 * 模型里段落和文字的格式已经按样式算好（样式 + 直接格式），保存时再与样式求差，所以：
 * - 套用样式 = 把“与旧样式相同的项”换成新样式的值，直接格式（与旧样式不同的项）保留；
 * - 样式的新建 / 修改记在 doc.styleEdits，保存时写进 styles.xml（patch.ts 的 styleUpserts），读 base 时叠在原样式上。
 * 还没有支持的部分：改样式时不改“基于它的其它样式”；不能改样式名、不能删样式。
 */
import type { StyleUpsert } from '../../vendor/docx-engine/src/style-upsert';
import { HF_STORIES, defaultParaProps, sectionList, type Doc, type Paragraph, type ParaProps, type RunProps, type StyleEdit } from './model';
import { comparePos, type Pos } from './layout';
import { getPara, mapRuns, paragraphsInRange, runPropsAt } from './ops';
import { HEADING_FORMAT } from './toc';

export interface StyleInfo {
  id: string;
  /** styles.xml 里的名字（英文内置名如 heading 1，或用户起的名字） */
  name: string;
  type: 'paragraph' | 'character';
  basedOn: string | null;
  /** 在样式库里显示（qFormat） */
  quick: boolean;
  custom: boolean;
}

/** 样式表：原文件的（docx.ts 的 styleSource）或新建文档的（catalogFromDoc） */
export interface StyleCatalog {
  list(): StyleInfo[];
  /** 段落样式 + 字符样式算出的格式（没有的样式 id 按默认样式算） */
  base(pStyle: string | undefined, rStyle?: string): { para: ParaProps; run: RunProps };
  idOf(canonical: string): string | null;
  canonical(id: string): string | null;
  defaultId(): string | null;
}

/** 新建文档（没有原文件）的样式表：只有默认样式，格式取文档默认 */
export function catalogFromDoc(doc: Doc): StyleCatalog {
  const builtin: StyleInfo[] = [
    { id: 'Normal', name: 'Normal', type: 'paragraph', basedOn: null, quick: true, custom: false },
    ...(['Heading1', 'Heading2', 'Heading3'] as const).map((id): StyleInfo => ({ id, name: `heading ${id.slice(-1)}`, type: 'paragraph', basedOn: 'Normal', quick: true, custom: false })),
  ];
  return {
    list: () => builtin,
    base: (p) => {
      const h = p && HEADING_FORMAT[p as keyof typeof HEADING_FORMAT];
      const para = defaultParaProps();
      if (!h) return { para, run: { ...doc.defaults } };
      return {
        para: { ...para, spaceBefore: h.before, spaceAfter: h.after, line: { rule: 'auto', value: h.line }, keepNext: true, keepLines: true, outline: h.outline, firstLine: 0, firstLineChars: null },
        run: { ...doc.defaults, size: h.size, bold: true },
      };
    },
    idOf: (c) => (builtin.some((s) => s.id === c) ? c : null),
    canonical: (id) => (/^(Heading|TOC)[1-9]$/.test(id) ? id : null),
    defaultId: () => 'Normal',
  };
}

const PARA_KEYS: (keyof ParaProps)[] = [
  'align', 'indentLeft', 'indentRight', 'firstLine', 'firstLineChars', 'spaceBefore', 'spaceAfter', 'spaceBeforeLines', 'spaceAfterLines',
  'spaceBeforeAuto', 'spaceAfterAuto', 'line', 'snapToGrid', 'keepNext', 'keepLines', 'widowControl', 'pageBreakBefore', 'num', 'tabs', 'outline', 'borders',
];
const RUN_KEYS: (keyof RunProps)[] = ['fontEA', 'fontLatin', 'size', 'bold', 'italic', 'underline', 'color', 'strike', 'vertAlign', 'spacing', 'highlight'];

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

/** 与旧样式相同的项换成新样式的值，其余（直接格式）保留 */
function rebase<T extends object>(value: T, oldBase: T, newBase: T, keys: (keyof T)[]): T {
  const out = { ...value };
  for (const k of keys) {
    if (!same(value[k], oldBase[k])) continue;
    if (newBase[k] === undefined) delete out[k];
    else out[k] = newBase[k];
  }
  return out;
}

/** 样式的修改叠在原样式上（新建的样式先取 basedOn 的格式，再叠自己的） */
export function withEdits(cat: StyleCatalog, edits: Record<string, StyleEdit> | undefined): StyleCatalog {
  if (!edits || !Object.keys(edits).length) return cat;
  const parentOf = new Map(cat.list().map((s) => [s.id, s.basedOn]));
  const resolve = (id: string | undefined, depth = 0): { para: ParaProps; run: RunProps } => {
    const e = id ? edits[id] : undefined;
    let start: { para: ParaProps; run: RunProps };
    if (e?.basedOn && depth < 20) start = resolve(e.basedOn, depth + 1);
    else {
      start = cat.base(id);
      // 原有的样式：它继承自上一级的那些项（值与上一级原来的相同），跟着上一级被改的值走；自己单独设过的项不动
      const parent = id ? parentOf.get(id) : null;
      if (parent && depth < 20) {
        const oldParent = cat.base(parent);
        const newParent = resolve(parent, depth + 1);
        start = { para: rebase(start.para, oldParent.para, newParent.para, PARA_KEYS), run: rebase(start.run, oldParent.run, newParent.run, RUN_KEYS) };
      }
    }
    if (!e) return start;
    return { para: { ...start.para, ...(e.type === 'paragraph' ? e.para : {}) }, run: { ...start.run, ...e.run } };
  };
  return {
    ...cat,
    list: () => {
      const known = new Set(cat.list().map((s) => s.id));
      const extra = Object.values(edits)
        .filter((e) => !known.has(e.id))
        .map((e): StyleInfo => ({ id: e.id, name: e.name ?? e.id, type: e.type, basedOn: e.basedOn ?? null, quick: true, custom: true }));
      return [...cat.list().map((s) => (edits[s.id]?.name ? { ...s, name: edits[s.id].name! } : s)), ...extra];
    },
    base: (p, r) => {
      const para = resolve(p);
      if (!r) return para;
      const ch = edits[r] ? resolve(r) : cat.base(p, r);
      // 字符样式：取它自己的文字格式，叠在段落样式的文字格式上
      const own = edits[r]?.run;
      const run = own ? { ...para.run, ...own } : ch.run;
      return { para: para.para, run };
    },
  };
}

const ZH: Record<string, string> = {
  normal: '正文', title: '标题', subtitle: '副标题', quote: '引用', 'intense quote': '明显引用', 'list paragraph': '列表段落', 'no spacing': '无间隔',
  header: '页眉', footer: '页脚', caption: '题注', hyperlink: '超链接', strong: '要点', emphasis: '强调', 'body text': '正文文本', 'plain text': '纯文本',
  'default paragraph font': '默认段落字体', 'footnote text': '脚注文本', 'footnote reference': '脚注引用', 'table of contents': '目录标题',
};

/** 样式的显示名：内置样式给中文名（Word 中文版的叫法），用户自己起的名字原样 */
export function styleLabel(s: Pick<StyleInfo, 'name' | 'id'>): string {
  const n = s.name.toLowerCase();
  const h = n.match(/^heading ([1-9])$/);
  if (h) return `标题 ${h[1]}`;
  const t = n.match(/^toc ([1-9])$/);
  if (t) return `目录 ${t[1]}`;
  return ZH[n] ?? s.name;
}

/** 应用于段落的样式 id（读入的有 styleId；只有标题 / 目录名的按样式表找） */
function paraStyleId(cat: StyleCatalog, p: { props: ParaProps }): string | undefined {
  return p.props.styleId ?? (p.props.style ? (cat.idOf(p.props.style) ?? undefined) : undefined);
}

/** 给选区所在的段落套用段落样式：直接格式保留，其余按新样式 */
export function applyParagraphStyle(doc: Doc, base: StyleCatalog, a: Pos, b: Pos, id: string): void {
  const cat = withEdits(base, doc.styleEdits);
  const next = cat.base(id);
  const canonical = cat.canonical(id);
  for (const { path } of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path });
    const oldId = paraStyleId(cat, p);
    const old = cat.base(oldId);
    p.props = rebase(p.props, old.para, next.para, PARA_KEYS);
    p.props.styleId = id;
    if (canonical) {
      p.props.style = canonical;
      const h = canonical.match(/^Heading(\d)$/);
      if (h && p.props.outline === undefined) p.props.outline = Number(h[1]) - 1;
    } else delete p.props.style;
    p.mark = rebase(p.mark, old.run, next.run, RUN_KEYS);
    p.runs = p.runs.map((r) => (r.opaque || r.mark ? r : { ...r, props: rebase(r.props, old.run, next.run, RUN_KEYS) }));
  }
}

/** 给选中的文字套用字符样式（id 为 null = 清除字符样式） */
export function applyCharacterStyle(doc: Doc, base: StyleCatalog, a: Pos, b: Pos, id: string | null): void {
  const cat = withEdits(base, doc.styleEdits);
  mapRuns(
    doc,
    a,
    b,
    (props, run, para) => {
      if (run.opaque || run.mark) return props;
      const pid = paraStyleId(cat, para);
      const old = cat.base(pid, props.rStyle).run;
      const next = cat.base(pid, id ?? undefined).run;
      const out = rebase(props, old, next, RUN_KEYS);
      if (id) out.rStyle = id;
      else delete out.rStyle;
      return out;
    },
    false,
  );
}

/**
 * 清除格式（Word 的“清除所有格式”）：选中文字的格式回到所在段落样式的文字格式（链接保留字符样式，其余去掉字符样式）；
 * paragraphs = true 时，选区涉及的段落也回到默认段落样式（正文）的段落格式、去掉编号，分节符和文本框架保留。
 * 修订标记、链接不动。
 */
export function clearFormatting(doc: Doc, base: StyleCatalog, a: Pos, b: Pos, paragraphs: boolean): void {
  const cat = withEdits(base, doc.styleEdits);
  if (paragraphs) {
    const def = cat.defaultId() ?? undefined;
    const next = cat.base(def);
    for (const { path } of paragraphsInRange(doc, a, b)) {
      const p = getPara(doc, { story: a.story, path });
      const keep = { sectionBreak: p.props.sectionBreak, frame: p.props.frame };
      p.props = { ...next.para, sectionBreak: keep.sectionBreak, ...(keep.frame && { frame: keep.frame }), num: null };
      if (def) p.props.styleId = def;
      delete p.props.style;
    }
  }
  mapRuns(doc, a, b, (props, run, para) => {
    if (run.opaque || run.mark) return props;
    const pid = paraStyleId(cat, para);
    const rStyle = props.link ? props.rStyle : undefined;
    const out: RunProps = { ...cat.base(pid, rStyle).run };
    if (rStyle) out.rStyle = rStyle;
    if (props.link) out.link = props.link;
    if (props.rev) out.rev = props.rev;
    return out;
  });
}

/** 光标所在段落当前的格式里，与默认样式不同的项（新建 / 更新样式用） */
function capture(doc: Doc, cat: StyleCatalog, pos: Pos, type: 'paragraph' | 'character'): Pick<StyleEdit, 'para' | 'run'> {
  const p = getPara(doc, pos);
  const pid = paraStyleId(cat, p);
  const def = cat.base(type === 'paragraph' ? undefined : pid);
  const props = runPropsAt(doc, pos);
  const run: NonNullable<StyleEdit['run']> = {};
  for (const k of ['fontEA', 'fontLatin', 'size', 'bold', 'italic', 'underline', 'color', 'strike'] as const) {
    if (!same(props[k], def.run[k])) (run as Record<string, unknown>)[k] = props[k] ?? false;
  }
  if (type === 'character') return { run };
  const para: NonNullable<StyleEdit['para']> = {};
  for (const k of ['align', 'indentLeft', 'indentRight', 'firstLine', 'spaceBefore', 'spaceAfter', 'line', 'keepNext', 'keepLines', 'pageBreakBefore', 'outline'] as const) {
    if (!same(p.props[k], def.para[k])) (para as Record<string, unknown>)[k] = p.props[k];
  }
  return { para, run };
}

export function checkStyleName(cat: StyleCatalog, doc: Doc, name: string): string | null {
  const n = name.trim();
  if (!n) return '请输入样式名';
  if (n.length > 60 || /[,;<>"&]/.test(n)) return '样式名不能含 , ; < > " & 且不超过 60 个字';
  const lower = n.toLowerCase();
  if (withEdits(cat, doc.styleEdits).list().some((s) => s.name.toLowerCase() === lower || styleLabel(s) === n)) return '已有同名的样式';
  return null;
}

/**
 * 以选区末尾所在段落（段落样式）或选中文字（字符样式）的格式新建样式；返回样式 id。新样式同时套用到选区。
 * from / to 是选区两端（光标时相同）。
 */
export function createStyle(doc: Doc, base: StyleCatalog, from: Pos, name: string, type: 'paragraph' | 'character', to: Pos = from): string {
  const bad = checkStyleName(base, doc, name);
  if (bad) throw new Error(bad);
  const cat = withEdits(base, doc.styleEdits);
  const edits = (doc.styleEdits ??= {});
  const used = new Set(cat.list().map((s) => s.id));
  let n = 1;
  while (used.has(`Wz${n}`)) n++;
  const id = `Wz${n}`;
  const cap = capture(doc, cat, to, type);
  edits[id] = { id, type, name: name.trim(), ...(type === 'paragraph' ? { basedOn: cat.defaultId() ?? undefined } : {}), ...cap };
  if (type === 'paragraph') applyParagraphStyle(doc, base, from, to, id);
  else applyCharacterStyle(doc, base, from, to, id);
  return id;
}

/** 文档里所有段落：正文、表格、各节页眉页脚（共用的只出一次）、文本框 */
export function eachParagraph(doc: Doc, fn: (p: Paragraph) => void): void {
  const seen = new Set<Paragraph>();
  const one = (p: Paragraph) => {
    if (seen.has(p)) return;
    seen.add(p);
    fn(p);
    for (const r of p.runs) if (r.box) r.box.paragraphs.forEach(one);
  };
  for (const b of doc.body) {
    if (b.kind === 'p') one(b);
    else for (const row of b.rows) for (const cell of row.cells) cell.paragraphs.forEach(one);
  }
  const arrays = new Set<Paragraph[]>();
  for (const sec of sectionList(doc)) for (const st of HF_STORIES) arrays.add(sec.hf[st]);
  for (const ps of arrays) for (const p of ps) one(p);
}

/** 样式表从 before 变成 after：每个段落、文字里“与旧样式相同”的项换成新样式的值，直接格式保留 */
function reflow(doc: Doc, before: StyleCatalog, after: StyleCatalog): void {
  eachParagraph(doc, (para) => {
    const pid = paraStyleId(before, para);
    const o = before.base(pid);
    const n = after.base(pid);
    para.props = rebase(para.props, o.para, n.para, PARA_KEYS);
    para.mark = rebase(para.mark, o.run, n.run, RUN_KEYS);
    para.runs = para.runs.map((r) => {
      if (r.opaque || r.mark) return r;
      const ro = before.base(pid, r.props.rStyle).run;
      const rn = after.base(pid, r.props.rStyle).run;
      return { ...r, props: rebase(r.props, ro, rn, RUN_KEYS) };
    });
  });
}

/** 用光标所在段落的格式更新样式，用到这个样式的段落和文字跟着变（各自的直接格式保留） */
export function redefineStyle(doc: Doc, base: StyleCatalog, id: string, pos: Pos): void {
  const before = withEdits(base, doc.styleEdits);
  const info = before.list().find((s) => s.id === id);
  if (!info) throw new Error('找不到这个样式');
  const prev = doc.styleEdits?.[id];
  // 样式的各项都取选中内容当前的格式（结果就是选中内容的样子）
  const edit: StyleEdit = { id, type: info.type, ...(prev?.name ? { name: prev.name } : {}), ...(prev?.basedOn ? { basedOn: prev.basedOn } : {}) };
  const p = getPara(doc, pos);
  const props = runPropsAt(doc, pos);
  edit.run = { fontEA: props.fontEA, fontLatin: props.fontLatin, size: props.size, bold: props.bold, italic: props.italic, underline: props.underline, color: props.color, strike: !!props.strike };
  if (info.type === 'paragraph') {
    const q = p.props;
    edit.para = { align: q.align, indentLeft: q.indentLeft, indentRight: q.indentRight, firstLine: q.firstLine, spaceBefore: q.spaceBefore, spaceAfter: q.spaceAfter, line: q.line, keepNext: q.keepNext, keepLines: q.keepLines, pageBreakBefore: q.pageBreakBefore, ...(q.outline !== undefined ? { outline: q.outline } : {}) };
  }
  const after = withEdits(base, { ...(doc.styleEdits ?? {}), [id]: edit });
  reflow(doc, before, after);
  (doc.styleEdits ??= {})[id] = edit;
}

/** 改样式名（只能改自己建的 / 文档里自定义的样式；内置样式的名字 Word 靠它认，不改） */
export function renameStyle(doc: Doc, base: StyleCatalog, id: string, name: string): void {
  const cat = withEdits(base, doc.styleEdits);
  const info = cat.list().find((s) => s.id === id);
  if (!info) throw new Error('找不到这个样式');
  if (!info.custom) throw new Error('内置样式不能改名');
  const n = name.trim();
  const clash = cat.list().some((s) => s.id !== id && (s.name.toLowerCase() === n.toLowerCase() || styleLabel(s) === n));
  if (!n) throw new Error('请输入样式名');
  if (clash) throw new Error('已有同名的样式');
  if (n.length > 60 || /[,;<>"&]/.test(n)) throw new Error('样式名不能含 , ; < > " & 且不超过 60 个字');
  const edits = (doc.styleEdits ??= {});
  edits[id] = { ...(edits[id] ?? { id, type: info.type }), name: n };
}

/**
 * 删除自定义样式：用到它的段落 / 文字改用它基于的样式（直接格式保留）；基于它的其它样式改成基于它的上一级。
 * 本次新建的直接丢掉；文档里原有的记入 doc.styleDeletes，保存时从 styles.xml 里去掉。
 */
export function deleteStyle(doc: Doc, base: StyleCatalog, id: string): void {
  const before = withEdits(base, doc.styleEdits);
  const info = before.list().find((s) => s.id === id);
  if (!info) throw new Error('找不到这个样式');
  if (!info.custom) throw new Error('内置样式不能删除');
  const parent = info.type === 'paragraph' ? ((doc.styleEdits?.[id]?.basedOn ?? info.basedOn) || before.defaultId()) : null;
  const edits = { ...(doc.styleEdits ?? {}) };
  delete edits[id];
  for (const e of Object.values(edits)) if (e.basedOn === id) edits[e.id] = { ...e, basedOn: parent ?? undefined };
  const after = withEdits(base, edits);
  eachParagraph(doc, (para) => {
    const pid = paraStyleId(before, para);
    const nid = info.type === 'paragraph' && pid === id ? (parent ?? undefined) : pid;
    const o = before.base(pid);
    const n = after.base(nid);
    para.props = rebase(para.props, o.para, n.para, PARA_KEYS);
    para.mark = rebase(para.mark, o.run, n.run, RUN_KEYS);
    if (nid !== pid) {
      if (nid) para.props.styleId = nid;
      else delete para.props.styleId;
      const canonical = nid ? after.canonical(nid) : null;
      if (canonical) para.props.style = canonical;
      else delete para.props.style;
    }
    para.runs = para.runs.map((r) => {
      if (r.opaque || r.mark) return r;
      const rs = r.props.rStyle === id ? undefined : r.props.rStyle;
      const next = rebase(r.props, before.base(pid, r.props.rStyle).run, after.base(nid, rs).run, RUN_KEYS);
      if (rs) next.rStyle = rs;
      else delete next.rStyle;
      return { ...r, props: next };
    });
  });
  doc.styleEdits = edits;
  if (base.list().some((s) => s.id === id)) (doc.styleDeletes ??= []).push({ id, parent });
}

/** styles.xml 里去掉一个样式；别的样式 basedOn / next 指向它的改成它的上一级（没有就去掉） */
export function removeStyleXml(xml: string, id: string, parent: string | null): string {
  const esc = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const el = new RegExp(`<w:style\\b[^>]*\\bw:styleId="${esc}"[^>]*?(?:/>|>[\\s\\S]*?</w:style>)`);
  let out = xml.replace(el, '');
  const ref = (tag: string) => new RegExp(`<w:${tag}\\s+w:val="${esc}"\\s*/>`, 'g');
  out = out.replace(ref('basedOn'), parent ? `<w:basedOn w:val="${parent}"/>` : '');
  out = out.replace(ref('next'), '').replace(ref('link'), '');
  return out;
}

// ---------------- 保存：styleEdits → docx-engine 的 StyleUpsert ----------------

const JC = { left: 'left', center: 'center', right: 'right', both: 'justify', distribute: 'justify' } as const;

export function editToUpsert(e: StyleEdit): StyleUpsert {
  const up: StyleUpsert = { styleId: e.id, type: e.type };
  if (e.name !== undefined) up.name = e.name;
  if (e.basedOn !== undefined) up.basedOn = e.basedOn;
  if (e.type === 'paragraph' && e.name !== undefined) up.next = e.id;
  const r = e.run;
  if (r) {
    up.rPr = {
      ...(r.bold !== undefined ? { bold: r.bold } : {}),
      ...(r.italic !== undefined ? { italic: r.italic } : {}),
      ...(r.underline !== undefined ? { underline: r.underline } : {}),
      ...(r.strike !== undefined ? { strike: r.strike } : {}),
      ...(r.color !== undefined ? { color: r.color === 'auto' ? null : r.color } : {}),
      ...(r.size !== undefined ? { sizeHalfPoints: Math.round(r.size * 2) } : {}),
      ...(r.fontLatin !== undefined ? { font: r.fontLatin } : {}),
      ...(r.fontEA !== undefined ? { eastAsiaFont: r.fontEA } : {}),
    };
  }
  const p = e.para;
  if (p && e.type === 'paragraph') {
    const tw = (pt: number) => Math.round(pt * 20);
    up.pPr = {
      ...(p.align !== undefined ? { align: JC[p.align] } : {}),
      ...(p.indentLeft !== undefined ? { indentLeftTwips: tw(p.indentLeft) } : {}),
      ...(p.indentRight !== undefined ? { indentRightTwips: tw(p.indentRight) } : {}),
      ...(p.firstLine !== undefined ? { firstLineTwips: tw(p.firstLine) } : {}),
      ...(p.spaceBefore !== undefined ? { spaceBeforeTwips: tw(p.spaceBefore) } : {}),
      ...(p.spaceAfter !== undefined ? { spaceAfterTwips: tw(p.spaceAfter) } : {}),
      ...(p.line ? (p.line.rule === 'auto' ? { lineSpacing: p.line.value } : { lineSpacingFixed: { rule: p.line.rule, twips: tw(p.line.value) } }) : {}),
      ...(p.keepNext !== undefined ? { keepNext: p.keepNext } : {}),
      ...(p.keepLines !== undefined ? { keepLines: p.keepLines } : {}),
      ...(p.pageBreakBefore !== undefined ? { pageBreakBefore: p.pageBreakBefore } : {}),
      ...(p.outline !== undefined ? { outlineLevel: p.outline + 1 } : {}),
    };
  }
  return up;
}

/** 比较两个位置是否同段（UI 用） */
export const sameParagraph = (a: Pos, b: Pos) => comparePos({ ...a, offset: 0 }, { ...b, offset: 0 }) === 0;

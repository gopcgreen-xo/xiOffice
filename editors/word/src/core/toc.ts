import { getPara, paragraphsInRange, paraLength } from './ops';
import type { DocLayout, Pos } from './layout';
import { newParagraph, paraText, textWidth, type Doc, type Paragraph, type ParaProps, type RunProps } from './model';

/** 段落样式：正文、标题 1–3（Word 中文版的内置样式默认值） */
export type ParaStyle = 'Normal' | 'Heading1' | 'Heading2' | 'Heading3';

export const PARA_STYLES: { id: ParaStyle; label: string }[] = [
  { id: 'Normal', label: '正文' },
  { id: 'Heading1', label: '标题 1' },
  { id: 'Heading2', label: '标题 2' },
  { id: 'Heading3', label: '标题 3' },
];

/** 内置标题样式的格式：标题 1 二号加粗、段前 17 段后 16.5、2.41 倍行距；标题 2 / 3 三号加粗、段前段后 13、1.73 倍行距 */
export const HEADING_FORMAT: Record<Exclude<ParaStyle, 'Normal'>, { size: number; before: number; after: number; line: number; outline: number }> = {
  Heading1: { size: 22, before: 17, after: 16.5, line: 2.41, outline: 0 },
  Heading2: { size: 16, before: 13, after: 13, line: 1.73, outline: 1 },
  Heading3: { size: 16, before: 13, after: 13, line: 1.73, outline: 2 },
};

/** 选中的段落套用样式：段落格式整体换成样式的，文字改字号和加粗（字体、颜色等不动） */
export function applyParaStyle(doc: Doc, a: Pos, b: Pos, style: ParaStyle) {
  for (const { path } of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path });
    const props: ParaProps = { ...p.props, line: { ...p.props.line } };
    let run: Partial<RunProps>;
    if (style === 'Normal') {
      delete props.style;
      delete props.outline;
      Object.assign(props, { spaceBefore: 0, spaceAfter: 0, spaceBeforeLines: null, spaceAfterLines: null, line: { rule: 'auto', value: 1 }, keepNext: false, keepLines: false });
      run = { size: doc.defaults.size, bold: false };
    } else {
      const h = HEADING_FORMAT[style];
      Object.assign(props, {
        style,
        outline: h.outline,
        spaceBefore: h.before,
        spaceAfter: h.after,
        spaceBeforeLines: null,
        spaceAfterLines: null,
        line: { rule: 'auto', value: h.line },
        keepNext: true,
        keepLines: true,
        firstLine: 0,
        firstLineChars: null,
      });
      run = { size: h.size, bold: true };
    }
    p.props = props;
    p.runs = p.runs.map((r) => ({ ...r, props: { ...r.props, ...run } }));
    p.mark = { ...p.mark, ...run };
  }
}

/** 当前段落的样式（下拉框显示用）；不是这几种时返回 null */
export function paraStyleOf(p: Paragraph): ParaStyle | null {
  if (!p.props.style && p.props.outline === undefined) return 'Normal';
  return PARA_STYLES.find((s) => s.id === p.props.style)?.id ?? null;
}

export interface Heading {
  /** 正文里的块下标 */
  bi: number;
  level: number;
  text: string;
}

/** 正文里有大纲级别的段落（标题），level 从 1 起；只看正文顶层段落，跳过目录本身 */
export function headings(doc: Doc, maxLevel = 3): Heading[] {
  const out: Heading[] = [];
  doc.body.forEach((b, bi) => {
    if (b.kind !== 'p' || b.props.outline === undefined || b.props.outline >= maxLevel || b.props.style?.startsWith('TOC')) return;
    const text = paraText(b).replace(/[-\u0000-\u001F]/g, '').trim();
    if (text) out.push({ bi, level: b.props.outline + 1, text });
  });
  return out;
}

/** 每个正文块从第几页开始（1 起），排版结果里第一次出现的页 */
export function blockPages(L: DocLayout): Map<number, number> {
  const m = new Map<number, number>();
  for (const pg of L.pages) for (const l of pg.body) if (!m.has(l.path[0])) m.set(l.path[0], pg.index + 1);
  return m;
}

/** 目录条目：样式 TOCn，按级别缩进 2 字符，页码右对齐到版心右边，中间点线 */
function tocEntry(doc: Doc, h: Heading, page: number | undefined): Paragraph {
  const size = doc.defaults.size;
  const p = newParagraph(`${h.text}\t${page ?? ''}`, {
    style: `TOC${h.level}`,
    indentLeft: (h.level - 1) * size * 2,
    tabs: [{ pos: textWidth(doc.setup), align: 'right', leader: 'dot' }],
  }, { ...doc.defaults });
  return p;
}

/** 正文里目录的位置：连续的 TOC 样式段落 [start, end) */
export function tocRange(doc: Doc): [number, number] | null {
  const start = doc.body.findIndex((b) => b.kind === 'p' && b.props.style?.startsWith('TOC'));
  if (start < 0) return null;
  let end = start;
  while (end < doc.body.length && doc.body[end].kind === 'p' && (doc.body[end] as Paragraph).props.style?.startsWith('TOC')) end++;
  return [start, end];
}

/**
 * 插入目录（光标所在段落之前；空段落时替换它）。pages 为块 → 页码；插入后页码会变，调用方重新排版后再 updateToc 一次。
 * 返回目录后面的位置；没有标题时返回 null。
 */
export function insertToc(doc: Doc, pos: Pos, pages: Map<number, number>): Pos | null {
  if (pos.story !== 'body' || pos.path.length !== 1) return null;
  const hs = headings(doc);
  if (!hs.length) return null;
  const at = pos.path[0];
  const entries = hs.map((h) => tocEntry(doc, h, pages.get(h.bi)));
  const empty = paraLength(getPara(doc, pos)) === 0 && !(doc.body[at] as Paragraph).props.sectionBreak;
  doc.body.splice(at, empty ? 1 : 0, ...entries, newParagraph('', {}, { ...doc.defaults }));
  return { story: 'body', path: [at + entries.length], offset: 0 };
}

/** 按当前标题和页码重建目录（标题增删改、页码变化）；没有目录时返回 false */
export function updateToc(doc: Doc, pages: Map<number, number>): boolean {
  const r = tocRange(doc);
  if (!r) return false;
  const [start, end] = r;
  // 页码表按更新前的块下标（标题的 bi 也是更新前的），先取条目再替换
  const hs = headings(doc);
  doc.body.splice(start, end - start, ...hs.map((h) => tocEntry(doc, h, pages.get(h.bi))));
  return true;
}

/** 光标是否在目录里 */
export function inToc(doc: Doc, pos: Pos): boolean {
  const r = tocRange(doc);
  return !!r && pos.story === 'body' && pos.path.length === 1 && pos.path[0] >= r[0] && pos.path[0] < r[1];
}

/**
 * 查找替换：正文（含表格）和各节的页眉页脚。匹配不跨段落；区分大小写、全字匹配可选。
 * 替换后的文字用匹配处第一个字的格式；全部替换按倒序进行，前面的位置不受影响。
 */
import { FIELD_CHAR, HF_STORIES, hfPath, paraText, sectionList, type Doc, type Paragraph, type Story } from './model';
import { comparePos, type Pos } from './layout';
import { deleteRange, getPara, insertRaw, paragraphPaths, runPropsAt } from './ops';

export interface FindOpts {
  matchCase?: boolean;
  /** 全字匹配：前后不能是字母数字（中文按字匹配，同 Word） */
  whole?: boolean;
}

export interface Match {
  story: Story;
  path: number[];
  offset: number;
  len: number;
}

/** 文档里所有的段落（正文在前，页眉页脚按节；与上一节相同的分区只算一次） */
function allParagraphs(doc: Doc): { story: Story; path: number[]; p: Paragraph }[] {
  const out: { story: Story; path: number[]; p: Paragraph }[] = [];
  for (const path of paragraphPaths(doc, 'body')) out.push({ story: 'body', path, p: getPara(doc, { story: 'body', path }) });
  const seen = new Set<Paragraph[]>();
  for (const sec of sectionList(doc)) {
    for (const story of HF_STORIES) {
      const list = sec.hf[story];
      if (seen.has(list)) continue;
      seen.add(list);
      list.forEach((p, i) => out.push({ story, path: hfPath(sec.owner, i), p }));
    }
  }
  return out;
}

const WORD = /[A-Za-z0-9_]/;

export function findAll(doc: Doc, query: string, o: FindOpts = {}): Match[] {
  if (!query) return [];
  const q = o.matchCase ? query : query.toLowerCase();
  const out: Match[] = [];
  for (const { story, path, p } of allParagraphs(doc)) {
    const raw = paraText(p);
    const text = o.matchCase ? raw : raw.toLowerCase();
    for (let i = text.indexOf(q); i >= 0; i = text.indexOf(q, i + Math.max(1, q.length))) {
      if (raw.slice(i, i + q.length).includes(FIELD_CHAR)) continue;
      if (o.whole && (WORD.test(raw[i - 1] ?? '') || WORD.test(raw[i + q.length] ?? ''))) continue;
      out.push({ story, path, offset: i, len: q.length });
    }
  }
  return out;
}

/** 从 from 之后（不含）找下一个，到末尾从头再来；back = 往前找 */
export function findNext(matches: Match[], from: Pos, back = false): Match | null {
  if (!matches.length) return null;
  const storyRank = (s: Story) => (s === 'body' ? 0 : 1 + HF_STORIES.indexOf(s));
  const cmp = (m: Match) => {
    const d = storyRank(m.story) - storyRank(from.story);
    return d || comparePos({ story: m.story, path: m.path, offset: m.offset }, from);
  };
  if (back) return [...matches].reverse().find((m) => cmp(m) < 0) ?? matches[matches.length - 1];
  return matches.find((m) => cmp(m) > 0) ?? matches[0];
}

/** 替换一处：文字用匹配处第一个字的格式。返回替换后文字末尾的位置 */
export function replaceMatch(doc: Doc, m: Match, repl: string): Pos {
  const start: Pos = { story: m.story, path: m.path, offset: m.offset };
  const props = { ...runPropsAt(doc, { ...start, offset: m.offset + 1 }) };
  deleteRange(doc, start, { ...start, offset: m.offset + m.len });
  return repl ? insertRaw(doc, start, repl, props) : start;
}

/** 全部替换，返回替换的处数 */
export function replaceAll(doc: Doc, query: string, repl: string, o: FindOpts = {}): number {
  const matches = findAll(doc, query, o);
  for (let i = matches.length - 1; i >= 0; i--) replaceMatch(doc, matches[i], repl);
  return matches.length;
}

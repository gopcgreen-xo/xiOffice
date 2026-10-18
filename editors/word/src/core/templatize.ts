/**
 * 合同 → 模版：把文档里的具体数据（公司名、合同号、金额…）换成 {{字段}}。
 * 建议（哪段、哪段文字、换成哪个字段）由 AI 给出，这里只负责在文档里找到并修改；
 * 修改走 deleteRange / insertRaw，修订模式打开时就记为修订（删除原文字 + 插入占位符），等用户审阅后接受或拒绝。
 * 匹配不跨段落；已标为删除的文字不算；已有的 {{…}} 里面不动。
 */
import { FIELD_CHAR, HF_STORIES, hfPath, sectionList, type Doc, type Paragraph, type RunProps, type Story } from './model';
import { deleteRange, getPara, insertRaw, paragraphPaths } from './ops';

/** 一条建议：p = 段落序号（templateParagraphs 的下标）；text = 要换掉的原文；after = 在这段文字后面插入（原文没有值、只有“合同编号：”这种标签时） */
export interface TemplateHit {
  p?: number;
  text?: string;
  after?: string;
  field: string;
}

export interface ParaRef {
  story: Story;
  path: number[];
}

/** 文档里有文字的段落（正文和表格在前，页眉页脚按节，与上一节相同的只算一次）和给 AI 看的文字（去掉已标为删除的字） */
export function templateParagraphs(doc: Doc): { refs: ParaRef[]; texts: string[] } {
  const refs: ParaRef[] = [];
  const texts: string[] = [];
  // 空段落（含没有内容的页眉页脚）没有可换的数据，不列出
  const add = (story: Story, path: number[], p: Paragraph) => {
    const t = liveText(p).text;
    if (!t.trim()) return;
    refs.push({ story, path });
    texts.push(t);
  };
  for (const path of paragraphPaths(doc, 'body')) add('body', path, getPara(doc, { story: 'body', path }));
  const seen = new Set<Paragraph[]>();
  for (const sec of sectionList(doc)) {
    for (const story of HF_STORIES) {
      const list = sec.hf[story];
      if (seen.has(list)) continue;
      seen.add(list);
      list.forEach((p, i) => add(story, hfPath(sec.owner, i), p));
    }
  }
  return { refs, texts };
}

/** 段落里没被删除的文字，以及每个字在原段落里的位置 */
function liveText(p: Paragraph): { text: string; raw: number[] } {
  let text = '';
  const raw: number[] = [];
  let acc = 0;
  for (const r of p.runs) {
    if (r.props.rev?.kind !== 'del') {
      text += r.text;
      for (let k = 0; k < r.text.length; k++) raw.push(acc + k);
    }
    acc += r.text.length;
  }
  return { text, raw };
}

/** 只给出段落、没有给出段落时可以全文替换的文字：至少 2 个字，且不是纯数字 / 符号（免得把“1”“13%”换遍全文） */
const globalOk = (t: string) => t.trim().length >= 2 && !/^[\d\s.,，。:：%％+\-±()（）/]+$/.test(t);

interface Edit {
  ref: ParaRef;
  /** 段落的标识（同一段落的修改才可能重叠） */
  key: string;
  /** 原段落里的范围 [from, to)；插入时 from = to */
  from: number;
  to: number;
  /** 在 live 文字里的范围（判断重叠用） */
  a: number;
  b: number;
  field: string;
}

export interface TemplateResult {
  /** 改了几处 */
  applied: number;
  /** 没有处理的建议及原因 */
  skipped: { hit: TemplateHit; reason: string }[];
}

/**
 * 按建议修改文档（调用方先记撤销、设置好修订开关）。fields = 允许的字段名（不在里面的建议跳过）。
 * 先在原文上找齐所有位置、去掉重叠（长的优先），再从后往前改，前面的位置不受影响。
 */
export function applyTemplateHits(doc: Doc, refs: ParaRef[], hits: TemplateHit[], fields: Set<string>): TemplateResult {
  const skipped: TemplateResult['skipped'] = [];
  const cache = new Map<number, { text: string; raw: number[]; holes: [number, number][]; len: number }>();
  const info = (i: number) => {
    let c = cache.get(i);
    if (!c) {
      const p = getPara(doc, refs[i]);
      const live = liveText(p);
      // 已有的占位符不动
      const holes: [number, number][] = [...live.text.matchAll(/\{\{[^{}]*\}\}/g)].map((m) => [m.index!, m.index! + m[0].length]);
      c = { ...live, holes, len: p.runs.reduce((n, r) => n + r.text.length, 0) };
      cache.set(i, c);
    }
    return c;
  };
  const inHole = (c: { holes: [number, number][] }, a: number, b: number) => c.holes.some(([x, y]) => a < y && b > x && !(a === b && (a === x || a === y)));

  const cands: Edit[] = [];
  const keyOf = (r: ParaRef) => `${r.story}:${r.path.join('.')}`;
  const push = (i: number, e: Omit<Edit, 'ref' | 'key'>) => cands.push({ ...e, ref: refs[i], key: keyOf(refs[i]) });
  for (const hit of hits) {
    const field = hit.field?.trim();
    if (!field || !fields.has(field)) {
      skipped.push({ hit, reason: `没有这个字段：${hit.field ?? ''}` });
      continue;
    }
    const insert = !hit.text;
    const needle = (insert ? hit.after : hit.text) ?? '';
    if (!needle.trim()) {
      skipped.push({ hit, reason: '没有给出原文' });
      continue;
    }
    let scope: number[];
    if (hit.p !== undefined && Number.isInteger(hit.p) && hit.p >= 0 && hit.p < refs.length) scope = [hit.p];
    else if (hit.p === undefined && !insert && globalOk(needle)) scope = refs.map((_, i) => i);
    else {
      skipped.push({ hit, reason: hit.p === undefined ? '没有指明段落' : `段落序号不对：${hit.p}` });
      continue;
    }
    let found = 0;
    for (const i of scope) {
      const c = info(i);
      const tries = [needle, needle.trim()].filter((t, k, arr) => t && arr.indexOf(t) === k);
      for (const t of tries) {
        let n = 0;
        for (let a = c.text.indexOf(t); a >= 0; a = c.text.indexOf(t, a + t.length)) {
          const b = a + t.length;
          if (c.text.slice(a, b).includes(FIELD_CHAR)) continue;
          if (insert) {
            // 插入点：标签后面；后面已经是占位符就不再插
            if (/^\s*\{\{/.test(c.text.slice(b))) continue;
            if (inHole(c, b, b)) continue;
            // 标签后面其实已经有值（AI 用错了 after）：换掉这个值，免得留下“{{船号}}H1702”
            const v = labelValue(c.text.slice(b));
            if (v) {
              const va = b + v.at;
              const vb = va + v.len;
              if (inHole(c, va, vb)) continue;
              push(i, { from: c.raw[va], to: c.raw[vb - 1] + 1, a: va, b: vb, field });
              n++;
              break;
            }
            // 表格里“标签一格、值一格”：标签在单元格末尾、右边那格是空的 → 填在右边那格（像填表一样，不挤在标签格里换行）
            const cell = !c.text.slice(b).trim() ? emptyRightCell(doc, refs[i]) : null;
            if (cell) {
              const len = getPara(doc, cell).runs.reduce((k, r) => k + r.text.length, 0);
              cands.push({ ref: cell, key: keyOf(cell), from: len, to: len, a: 0, b: 0, field });
              n++;
              break;
            }
            const at = b < c.raw.length ? c.raw[b] : c.len;
            push(i, { from: at, to: at, a: b, b, field });
            n++;
            break; // 标签只在第一次出现的地方插
          }
          if (inHole(c, a, b)) continue;
          push(i, { from: c.raw[a], to: c.raw[b - 1] + 1, a, b, field });
          n++;
        }
        if (n) {
          found += n;
          break;
        }
      }
    }
    if (!found) skipped.push({ hit, reason: needle.includes(FIELD_CHAR) ? '这段文字里有域或对象（￼），不能整段替换，请手工处理' : insert ? '没找到标签文字' : '原文里没找到这段文字' });
  }

  // 去重叠：替换按长度优先；插入点落在某个替换范围里面时不要
  const keep: Edit[] = [];
  const overlaps = (x: Edit, y: Edit) => {
    if (x.key !== y.key) return false;
    if (x.a === x.b && y.a === y.b) return x.a === y.a;
    if (x.a === x.b) return x.a > y.a && x.a < y.b;
    if (y.a === y.b) return y.a > x.a && y.a < x.b;
    return x.a < y.b && y.a < x.b;
  };
  for (const e of [...cands].sort((x, y) => y.b - y.a - (x.b - x.a))) if (!keep.some((k) => overlaps(k, e))) keep.push(e);

  // 从后往前改
  // （不同段落互不影响；同一段落里从后往前，前面的位置不变）
  keep.sort((x, y) => (x.key < y.key ? -1 : x.key > y.key ? 1 : y.from - x.from || y.to - x.to));
  for (const e of keep) {
    const ref = e.ref;
    const start = { story: ref.story, path: ref.path, offset: e.from };
    const props = propsAt(getPara(doc, ref), e.to > e.from ? e.from : Math.max(0, e.from - 1));
    if (e.to > e.from) deleteRange(doc, start, { ...start, offset: e.to });
    insertRaw(doc, start, `{{${e.field}}}`, props);
  }
  return { applied: keep.length, skipped };
}

/**
 * 正文表格里，段落 ref 是某格的最后一段、同一行右边紧挨着的那格没有文字（已标为删除的不算）时，返回那格的第一段；否则 null。
 * 页眉页脚里的表格只读，不处理。
 */
function emptyRightCell(doc: Doc, ref: ParaRef): ParaRef | null {
  if (ref.story !== 'body' || ref.path.length !== 4) return null;
  const [bi, r, ci, pi] = ref.path;
  const t = doc.body[bi];
  if (t?.kind !== 'table') return null;
  const row = t.rows[r];
  if (pi !== row.cells[ci].paragraphs.length - 1) return null;
  const right = row.cells[ci + 1];
  // 纵向合并的下半格不能放内容（显示的是上面那格）
  if (!right?.paragraphs.length || right.vMerge === 'continue' || right.paragraphs.some((p) => liveText(p).text.trim())) return null;
  return { story: 'body', path: [bi, r, ci + 1, 0] };
}

/**
 * 标签后面的值：跳过开头的空格，到连续两个空格 / 制表符 / 段尾为止（中间单个空格算值的一部分），去掉末尾的句号分号等。
 * 没有值、或者紧跟着的是下一个标签（含冒号，如“电话：    传真：”）时返回 null。
 */
function labelValue(rest: string): { at: number; len: number } | null {
  const m = /^([ 　]*)([^\s　](?:[^\s　]| (?! ))*)/.exec(rest);
  if (!m) return null;
  const value = m[2].replace(/[。．.；;，,、]+$/, '');
  if (!value || /[:：]/.test(value) || value.includes(FIELD_CHAR)) return null;
  return { at: m[1].length, len: value.length };
}

/** 原段落位置 off 处那个字的格式（不带修订标记），用作占位符的格式 */
function propsAt(p: Paragraph, off: number): RunProps {
  let acc = 0;
  for (const r of p.runs) {
    if (off >= acc && off < acc + r.text.length) {
      const out = { ...r.props };
      delete out.rev;
      return out;
    }
    acc += r.text.length;
  }
  const out = { ...(p.runs.at(-1)?.props ?? p.mark) };
  delete out.rev;
  return out;
}

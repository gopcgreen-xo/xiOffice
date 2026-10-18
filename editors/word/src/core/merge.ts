/**
 * 模板与批量生成（邮件合并）：文档里写 {{字段名}}，
 * 每条数据生成一份填好的文档，或把所有数据合并成一个文档。占位符可以被 Word 拆到多个文本段里，替换后沿用占位符开头那个字的格式。
 * 也认 Word 自带的邮件合并域（MERGEFIELD，复杂域和简单域）：显示为 «字段名»，\b 前缀 / \f 后缀 / \# 数字 / \@ 日期 / \* 大小写开关有效；
 * 规则域 IF（可嵌套 MERGEFIELD）、NEXT、NEXTIF、SKIPIF 也会求值（NEXT 让一份文档里放多条数据，如标签、名录）。
 * 数据来自表格 / csv / json（见下面的解析函数）；只做文字替换，不改文档的其它部分。
 * 脚注 / 尾注、页眉页脚里的表格：显示用的模型里填，保存时对原始 XML 做同样的替换（Doc.xmlFill，见 patch.ts）。
 * 还没有支持的部分：IF 的结果里再嵌套别的规则域、ASK / FILLIN / SET 等交互域、\# 的 x 截断位和 \@ 的中文日期专用图片、
 * 合并成一个文档时只支持只有一节的模板。
 */
import { FIELD_CHAR, HF_STORIES, isAtom, sectionList, type Block, type Doc, type Opaque, type Paragraph, type Run, type RunProps } from './model';
import type { Pos } from './layout';
import { getPara, runPropsAt } from './ops';
import { evalRule, mergeText, parseRule, ruleLabel, type MergeSpec } from './mergeInstr';

export { parseMergeInstr, parseRule } from './mergeInstr';

/** 邮件合并域在编辑器里的样子：«字段名»（按 EQ 域的显示方式排版） */
export function mergeFieldView(merge: NonNullable<Opaque['merge']>, props: RunProps): Pick<Opaque, 'merge' | 'text' | 'eq'> {
  const text = `«${merge.name}»`;
  return { merge, text, eq: { overstrike: false, align: 'left', parts: [[{ text, props }]] } };
}

/** 规则域的显示 */
export function ruleFieldView(full: string, props: RunProps): Pick<Opaque, 'rule' | 'text' | 'eq'> | null {
  const rule = parseRule(full);
  if (!rule) return null;
  const text = ruleLabel(rule);
  return { rule: { full }, text, eq: { overstrike: false, align: 'left', parts: [[{ text, props }]] } };
}

/** 插入用的邮件合并域 XML（Word 的简单域写法），保存后 Word 里就是邮件合并域 */
export function mergeFieldXml(name: string, props: RunProps): string {
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const rpr =
    `<w:rPr><w:rFonts w:ascii="${esc(props.fontLatin)}" w:hAnsi="${esc(props.fontLatin)}" w:eastAsia="${esc(props.fontEA)}"/>` +
    `${props.bold ? '<w:b/>' : ''}${props.italic ? '<w:i/>' : ''}<w:noProof/><w:sz w:val="${Math.round(props.size * 2)}"/></w:rPr>`;
  const field = /\s/.test(name) ? `&quot;${esc(name)}&quot;` : esc(name);
  return `<w:fldSimple w:instr=" MERGEFIELD ${field} \\* MERGEFORMAT "><w:r>${rpr}<w:t>«${esc(name)}»</w:t></w:r></w:fldSimple>`;
}

/** 在位置处插入一个原样的域对象（返回新位置） */
function insertFieldRun(doc: Doc, pos: Pos, opaque: Opaque): Pos {
  const props = runPropsAt(doc, pos);
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
      break;
    }
    acc += r.text.length;
  }
  p.runs.splice(i, 0, { text: FIELD_CHAR, props: { ...props }, opaque });
  return { ...pos, offset: pos.offset + 1 };
}

/** 在位置处插入邮件合并域 */
export function insertMergeField(doc: Doc, pos: Pos, name: string): Pos {
  const n = name.trim();
  if (!n || /["\n]/.test(n)) throw new Error('字段名不能为空，也不能含引号或换行');
  const props = runPropsAt(doc, pos);
  return insertFieldRun(doc, pos, { xml: mergeFieldXml(n, props), level: 'para', kind: 'other', w: 0, h: 0, ...mergeFieldView({ name: n }, props) });
}

/** 在位置处插入“下一条记录”（NEXT）规则域：一份文档里放多条数据时用 */
export function insertNextRecord(doc: Doc, pos: Pos): Pos {
  const props = runPropsAt(doc, pos);
  const view = ruleFieldView(' NEXT ', props)!;
  const rpr = `<w:rPr><w:noProof/><w:sz w:val="${Math.round(props.size * 2)}"/></w:rPr>`;
  const xml = `<w:fldSimple w:instr=" NEXT "><w:r>${rpr}<w:t>«下一条记录»</w:t></w:r></w:fldSimple>`;
  return insertFieldRun(doc, pos, { xml, level: 'para', kind: 'other', w: 0, h: 0, ...view });
}

const PLACEHOLDER = /\{\{\s*([^{}\n]+?)\s*\}\}/g;

/** 文档里所有段落（正文、表格、文本框、各节页眉页脚及其中的表格、脚注），共用的页眉页脚数组只出一次 */
function allParagraphs(doc: Doc): Paragraph[] {
  const out: Paragraph[] = [];
  const seen = new Set<Paragraph>();
  const add = (p: Paragraph) => {
    if (seen.has(p)) return;
    seen.add(p);
    if (p.hfTable) {
      for (const row of p.hfTable.table.rows) for (const c of row.cells) c.paragraphs.forEach(add);
      return;
    }
    out.push(p);
    for (const r of p.runs) if (r.box) r.box.paragraphs.forEach(add);
  };
  for (const b of doc.body) {
    if (b.kind === 'p') add(b);
    else for (const row of b.rows) for (const c of row.cells) c.paragraphs.forEach(add);
  }
  const lists = new Set<Paragraph[]>();
  for (const s of sectionList(doc)) for (const st of HF_STORIES) lists.add(s.hf[st]);
  for (const ps of lists) ps.forEach(add);
  for (const ps of Object.values(doc.footnotes ?? {})) ps.forEach(add);
  return out;
}

const textOf = (p: Paragraph) => p.runs.map((r) => r.text).join('');

export interface PlaceholderInfo {
  key: string;
  count: number;
}

/** 文档里用到的占位符（按第一次出现的顺序）和出现次数（含 MERGEFIELD 和规则域里嵌套的字段） */
export function placeholders(doc: Doc): PlaceholderInfo[] {
  const map = new Map<string, number>();
  for (const p of allParagraphs(doc)) {
    // 文字占位符和 Word 域，按在段落里的位置排
    const found: { at: number; key: string }[] = [];
    for (const m of textOf(p).matchAll(PLACEHOLDER)) found.push({ at: m.index!, key: m[1].trim() });
    let acc = 0;
    for (const r of p.runs) {
      if (r.opaque?.merge) found.push({ at: acc, key: r.opaque.merge.name });
      if (r.opaque?.rule) {
        const rule = parseRule(r.opaque.rule.full);
        for (const t of rule?.toks ?? []) if (t.t === 'field') found.push({ at: acc, key: t.m.name });
      }
      acc += r.text.length;
    }
    for (const f of found.sort((a, b) => a.at - b.at)) map.set(f.key, (map.get(f.key) ?? 0) + 1);
  }
  return [...map].map(([key, count]) => ({ key, count }));
}

/** 文档里有没有 NEXT 类规则域（有的话一份文档会用掉不止一条数据） */
export function hasNextRule(doc: Doc): boolean {
  return allParagraphs(doc).some((p) =>
    p.runs.some((r) => {
      const k = r.opaque?.rule && parseRule(r.opaque.rule.full)?.kind;
      return k === 'next' || k === 'nextif' || k === 'skipif';
    }),
  );
}

/** 把段落里 [s, e) 换成 text（沿用 s 处文字的格式；范围里的图片、书签等原样对象一并去掉） */
function replaceRange(p: Paragraph, s: number, e: number, text: string) {
  const out: Run[] = [];
  let acc = 0;
  let inserted = false;
  let like: Run | undefined;
  const put = () => {
    if (inserted) return;
    inserted = true;
    out.push({ text, props: stripRev((like ?? p.runs[p.runs.length - 1])?.props ?? p.mark) });
  };
  for (const r of p.runs) {
    const a = acc;
    const b = acc + r.text.length;
    acc = b;
    if (b <= s) out.push(r);
    else if (a >= e) {
      put();
      out.push(r);
    } else {
      like ??= r;
      const atom = isAtom(r);
      if (!atom && a < s) out.push({ ...r, text: r.text.slice(0, s - a) });
      put();
      if (!atom && b > e) out.push({ ...r, props: { ...r.props }, text: r.text.slice(e - a) });
    }
  }
  put();
  p.runs = out.filter((r) => r.text.length > 0);
}

const stripRev = <T extends { rev?: unknown }>(props: T): T => {
  const { rev: _rev, ...rest } = props;
  return { ...rest } as T;
};

/** 值里的换行变成文档里的手动换行；去掉不能写进 XML 的控制字符；占位符标记符号不允许出现在值里（避免再次被替换） */
function cleanValue(v: string): string {
  // eslint-disable-next-line no-control-regex
  return v.replace(/\r\n?/g, '\n').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').replaceAll(FIELD_CHAR, '').replace(/\{\{/g, '{ {');
}

export interface FillResult {
  /** 模板里有、但数据里没有的字段 */
  missing: string[];
  /** 实际替换了多少处 */
  replaced: number;
  /** SKIPIF 条件成立：这一份文档不要 */
  skipped: boolean;
  /** 这份文档用掉了几条数据（有 NEXT 时大于 1） */
  consumed: number;
}

/** 填值时当前数据的游标：NEXT 往后走，走过头后字段都填成空 */
interface Cursor {
  index: number;
  rec(): Record<string, string>;
  exhausted(): boolean;
  next(): void;
}

function fillWith(doc: Doc, cur: Cursor, start: number, startRec: Record<string, string>, opts: { blankMissing?: boolean }): FillResult {
  const missing = new Set<string>();
  let replaced = 0;
  let skipped = false;
  const lookup = (name: string): string | undefined => {
    if (cur.exhausted()) return '';
    const rec = cur.rec();
    if (Object.prototype.hasOwnProperty.call(rec, name)) return rec[name] ?? '';
    missing.add(name);
    return opts.blankMissing ? '' : undefined;
  };
  for (const p of allParagraphs(doc)) {
    // 段落里的 {{字段}}、邮件合并域、规则域，按先后顺序求值（NEXT 会让它后面的字段取下一条数据），再从后往前换
    interface Hit {
      s: number;
      e: number;
      kind: 'text' | 'merge' | 'rule';
      key?: string;
      merge?: MergeSpec;
      rule?: NonNullable<ReturnType<typeof parseRule>>;
    }
    const hits: Hit[] = [];
    let acc = 0;
    for (const r of p.runs) {
      const rl = r.opaque?.rule && parseRule(r.opaque.rule.full);
      if (r.opaque?.merge) hits.push({ s: acc, e: acc + r.text.length, kind: 'merge', merge: r.opaque.merge });
      else if (rl) hits.push({ s: acc, e: acc + r.text.length, kind: 'rule', rule: rl });
      acc += r.text.length;
    }
    const text = textOf(p);
    if (text.includes('{{')) for (const m of text.matchAll(PLACEHOLDER)) hits.push({ s: m.index!, e: m.index! + m[0].length, kind: 'text', key: m[1].trim() });
    if (!hits.length) continue;
    hits.sort((x, y) => x.s - y.s);
    const edits: { s: number; e: number; text: string }[] = [];
    for (const h of hits) {
      if (h.kind === 'text') {
        const v = lookup(h.key!);
        if (v === undefined) continue;
        replaced++;
        edits.push({ s: h.s, e: h.e, text: cleanValue(v) });
      } else if (h.kind === 'merge') {
        const v = lookup(h.merge!.name);
        if (v === undefined) continue;
        replaced++;
        edits.push({ s: h.s, e: h.e, text: mergeText(h.merge!, cleanValue(v)) });
      } else {
        const res = evalRule(h.rule!, (spec: MergeSpec) => mergeText(spec, cleanValue(lookup(spec.name) ?? '')));
        replaced++;
        if (res.kind === 'next') cur.next();
        else if (res.kind === 'skip') {
          skipped = true;
          cur.next();
        }
        edits.push({ s: h.s, e: h.e, text: res.kind === 'text' ? cleanValue(res.text) : '' });
      }
    }
    for (const e of edits.reverse()) replaceRange(p, e.s, e.e, e.text);
  }
  // 页眉页脚里的表格、脚注 / 尾注：保存时对原始 XML 做同样的替换，值取开始时的那条数据
  const lookupStart = (k: string): string | undefined => (Object.prototype.hasOwnProperty.call(startRec, k) ? startRec[k] : opts.blankMissing ? '' : undefined);
  for (const p of hfParagraphs(doc)) if (p.hfTable) p.hfTable.xml = fillPartXml(p.hfTable.xml, lookupStart, new Set());
  doc.xmlFill = { values: startRec, blank: !!opts.blankMissing };
  return { missing: [...missing], replaced, skipped, consumed: Math.max(1, cur.index - start + 1) };
}

function hfParagraphs(doc: Doc): Paragraph[] {
  const out: Paragraph[] = [];
  const seen = new Set<Paragraph[]>();
  for (const s of sectionList(doc)) {
    for (const st of HF_STORIES) {
      const ps = s.hf[st];
      if (seen.has(ps)) continue;
      seen.add(ps);
      out.push(...ps);
    }
  }
  return out;
}

/**
 * 按一条数据把文档里的 {{字段}}、邮件合并域换成值（原地修改，调用方先复制文档）。
 * 数据里没有的字段：blankMissing 为真时换成空，否则原样保留（并记入 missing）。NEXT 类规则域在单条数据时被去掉。
 */
export function fillDoc(doc: Doc, values: Record<string, string>, opts: { blankMissing?: boolean } = {}): FillResult {
  const cur: Cursor = { index: 0, rec: () => values, exhausted: () => false, next: () => {} };
  return fillWith(doc, cur, 0, values, opts);
}

/**
 * 从第 start 条数据开始填一份文档：文档里的 NEXT / NEXTIF / SKIPIF 会往后取数据，返回用掉了几条（consumed）。
 * 数据用完后剩下的字段填成空。
 */
export function fillRecords(doc: Doc, records: Record<string, string>[], start: number, opts: { blankMissing?: boolean } = {}): FillResult {
  const cur: Cursor = {
    index: start,
    rec: () => records[cur.index] ?? {},
    exhausted: () => cur.index >= records.length,
    next: () => void cur.index++,
  };
  return fillWith(doc, cur, start, records[start] ?? {}, opts);
}

// ---------------- 合并成一个文档 ----------------

/** 复制出来的那份正文：去掉原文件块号（保存时重新生成）、书签（名字不能重复）和批注标记 */
function detach(blocks: Block[]) {
  const para = (p: Paragraph) => {
    delete p.src;
    p.runs = p.runs.filter((r) => !(r.mark || (r.opaque?.kind === 'bookmark' && /^<w:bookmark/.test(r.opaque.xml))));
  };
  for (const b of blocks) {
    if (b.kind === 'p') para(b);
    else {
      delete b.src;
      for (const row of b.rows) for (const c of row.cells) c.paragraphs.forEach(para);
    }
  }
}

export interface CombineResult {
  doc: Doc;
  /** 合并进去的份数 */
  count: number;
  skipped: number;
  missing: string[];
}

/**
 * 所有数据合并成一个文档：模板正文按数据重复，每份从新的一页开始；页眉页脚、页面设置用模板的（页眉页脚里的字段按第一条数据填）。
 * 模板只能有一节。
 */
export function combineRecords(template: Doc, records: Record<string, string>[], opts: { blankMissing?: boolean } = {}): CombineResult {
  if (sectionList(template).length > 1) throw new Error('合并成一个文档只支持只有一节的模板；请改成“每条数据一个文件”，或先去掉分节符');
  const out = structuredClone(template);
  out.body = [];
  const missing = new Set<string>();
  let count = 0;
  let skipped = 0;
  let first = true;
  for (let i = 0; i < records.length; ) {
    const copy = structuredClone(template);
    const r = fillRecords(copy, records, i, opts);
    r.missing.forEach((m) => missing.add(m));
    i += r.consumed;
    if (r.skipped) {
      skipped++;
      continue;
    }
    if (first) {
      // 页眉页脚、脚注、数据来源按第一份
      out.header = copy.header;
      out.footer = copy.footer;
      out.headerFirst = copy.headerFirst;
      out.footerFirst = copy.footerFirst;
      out.headerEven = copy.headerEven;
      out.footerEven = copy.footerEven;
      if (copy.footnotes) out.footnotes = copy.footnotes;
      if (copy.xmlFill) out.xmlFill = copy.xmlFill;
    } else detach(copy.body);
    const head = copy.body[0];
    if (!first && head) {
      if (head.kind === 'p') head.props = { ...head.props, pageBreakBefore: true };
      else {
        const sep = structuredClone(head.rows[0].cells[0].paragraphs[0]);
        sep.runs = [];
        sep.props = { ...sep.props, pageBreakBefore: true };
        delete sep.src;
        copy.body.unshift(sep);
      }
    }
    out.body.push(...copy.body);
    count++;
    first = false;
  }
  if (!out.body.length) out.body = [structuredClone(template.body[0])];
  return { doc: out, count, skipped, missing: [...missing] };
}

// ---------------- 部件 XML 里的占位符（脚注 / 尾注、页眉页脚表格） ----------------

const xmlUnescape = (s: string) =>
  s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, e: string) => {
    const k = e.toLowerCase();
    if (k === 'amp') return '&';
    if (k === 'lt') return '<';
    if (k === 'gt') return '>';
    if (k === 'quot') return '"';
    if (k === 'apos') return "'";
    return String.fromCodePoint(k.startsWith('#x') ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10));
  });
const xmlEscape = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** <w:t> 的内容：换行转为 <w:br/> */
function textNode(s: string): string {
  return s
    .split('\n')
    .map((line) => `<w:t xml:space="preserve">${xmlEscape(line)}</w:t>`)
    .join('<w:br/>');
}

/**
 * 替换一段 XML 里的 {{占位符}}（可以跨多个 <w:t>，不跨段落）：值写进占位符开头所在的 <w:t>，占位符在其它 <w:t> 里的部分删除。
 * lookup 返回 undefined 的占位符原样保留（记入 missing）。
 */
export function fillPartXml(xml: string, lookup: (key: string) => string | undefined, missing: Set<string>): string {
  const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:t(?:\s[^>]*)?\/>/g;
  interface Node {
    start: number;
    end: number;
    text: string;
    offset: number;
  }
  const nodes: Node[] = [];
  let full = '';
  let last = 0;
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    if (nodes.length && xml.slice(last, m.index).includes('</w:p>')) full += '\u0000';
    const text = xmlUnescape(m[1] ?? '');
    nodes.push({ start: m.index, end: m.index + m[0].length, text, offset: full.length });
    full += text;
    last = m.index + m[0].length;
  }
  if (!full.includes('{{')) return xml;
  const matches: { s: number; e: number; value: string }[] = [];
  for (const m of full.matchAll(/\{\{\s*([^{}\u0000]+?)\s*\}\}/g)) {
    const key = m[1].trim();
    const value = lookup(key);
    if (value === undefined) {
      missing.add(key);
      continue;
    }
    matches.push({ s: m.index!, e: m.index! + m[0].length, value: cleanValue(value) });
  }
  if (!matches.length) return xml;
  let out = '';
  let pos = 0;
  for (const n of nodes) {
    const ns = n.offset;
    const ne = n.offset + n.text.length;
    const hit = matches.filter((m) => m.s < ne && m.e > ns);
    if (!hit.length) continue;
    let text = '';
    let i = ns;
    for (const m of hit) {
      if (m.s > i) text += n.text.slice(i - ns, m.s - ns);
      if (m.s >= ns) text += m.value;
      i = Math.max(i, Math.min(m.e, ne));
    }
    if (i < ne) text += n.text.slice(i - ns);
    out += xml.slice(pos, n.start) + textNode(text);
    pos = n.end;
  }
  return out + xml.slice(pos);
}

// ---------------- 数据 ----------------

export interface DataTable {
  /** 列名（第一行） */
  columns: string[];
  /** 每条数据：列名 → 文字 */
  records: Record<string, string>[];
}

/** csv：支持引号、引号里的逗号 / 换行、双引号转义；分隔符自动识别逗号 / 制表符 / 分号 */
export function parseCsv(input: string): string[][] {
  const text = input.replace(/^﻿/, '');
  const first = text.split(/\r?\n/, 1)[0] ?? '';
  const delim = [',', '\t', ';'].map((d) => [d, first.split(d).length] as const).sort((a, b) => b[1] - a[1])[0][0];
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"') {
        if (text[i + 1] === '"') (cell += '"', i++);
        else quoted = false;
      } else cell += c;
    } else if (c === '"' && !cell) quoted = true;
    else if (c === delim) (row.push(cell), (cell = ''));
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell);
      cell = '';
      rows.push(row);
      row = [];
    } else cell += c;
  }
  if (cell || row.length) (row.push(cell), rows.push(row));
  return rows;
}

/** 第一行是列名，其余每行一条数据；空行跳过；列名重复时后面的加序号 */
export function tableFromRows(rows: (string | number | null | undefined)[][]): DataTable {
  const str = (v: string | number | null | undefined) => (v === null || v === undefined ? '' : typeof v === 'number' ? String(Math.round(v * 1e9) / 1e9) : v);
  const head = (rows[0] ?? []).map((h) => str(h).trim());
  const seen = new Map<string, number>();
  const columns = head.map((h, i) => {
    const base = h || `列${i + 1}`;
    const n = (seen.get(base) ?? 0) + 1;
    seen.set(base, n);
    return n === 1 ? base : `${base}${n}`;
  });
  const records = rows
    .slice(1)
    .filter((r) => r.some((v) => str(v).trim() !== ''))
    .map((r) => Object.fromEntries(columns.map((c, i) => [c, str(r[i])])));
  return { columns, records };
}

/** json：对象数组（列名取各对象的键，按首次出现的顺序） */
export function tableFromJson(input: string): DataTable {
  const data: unknown = JSON.parse(input);
  const list = Array.isArray(data) ? data : [data];
  if (!list.every((x) => x && typeof x === 'object' && !Array.isArray(x))) throw new Error('JSON 应当是对象数组，如 [{"名称":"甲"}]');
  const columns: string[] = [];
  const records = (list as Record<string, unknown>[]).map((o) => {
    const rec: Record<string, string> = {};
    for (const [k, v] of Object.entries(o)) {
      if (!columns.includes(k)) columns.push(k);
      rec[k] = v === null || v === undefined ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v);
    }
    return rec;
  });
  return { columns, records };
}

/** 输出文件名：模板里的 {{列名}} 换成这条数据的值，{#} 是序号（从 1 起）；去掉不能做文件名的字符；批次内重名的加序号 */
export function outputName(template: string, values: Record<string, string>, index: number, taken: Set<string>): string {
  let name = template
    .replace(/\{#\}/g, String(index + 1))
    .replace(PLACEHOLDER, (_, k: string) => values[k.trim()] ?? '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '');
  if (!name) name = `文档${index + 1}`;
  if (name.length > 80) name = name.slice(0, 80);
  let out = name;
  for (let n = 2; taken.has(out.toLowerCase()); n++) out = `${name} (${n})`;
  taken.add(out.toLowerCase());
  return out;
}

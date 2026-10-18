/**
 * 一般的域（Run.fc 标记的 fldChar 域）：解析域代码、计算结果、更新域、改域代码。
 * 规则按 Word：域代码 = 域名 + 参数 + 开关（\@ 日期图片、\# 数字图片、\* 格式，可有多个 \*）；嵌套的域先算；
 * 算不出来的（文档属性、目录、索引等本编辑器没有的数据）保留原结果。
 */
import { applyCase, formatDate, formatNumber, parseNumber } from './mergeFormat';
import { bookmarkText } from './links';
import { eachParagraph } from './styles';
import { getPara, paragraphsInRange } from './ops';
import type { Pos } from './layout';
import type { Doc, Paragraph, Run } from './model';
import { fieldSpans, type FieldSpan } from './fieldView';

// ---------- 域代码 ----------

export interface Instr {
  /** 域名（原样大小写）；公式域为 '=' */
  name: string;
  args: string[];
  /** 开关，按出现顺序：key 不含反斜杠（\@ → '@'，\p → 'p'） */
  sw: { key: string; value?: string }[];
}

/** 带值的开关：通用的 \@ \# \*，SEQ 的 \r \s，REF 的 \d */
const VALUE_SW = new Set(['@', '#', '*', 'r', 's', 'd']);

interface Tok {
  text: string;
  quoted: boolean;
}

function tokens(s: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (c === '"') {
      // 引号里的文字：\" 和 \\ 是转义
      let v = '';
      i++;
      while (i < s.length && s[i] !== '"') {
        if (s[i] === '\\' && (s[i + 1] === '"' || s[i + 1] === '\\')) i++;
        v += s[i++];
      }
      i++;
      out.push({ text: v, quoted: true });
      continue;
    }
    if (c === '{') {
      // 嵌套的域整个当一个参数
      let depth = 0;
      let v = '';
      while (i < s.length) {
        const ch = s[i++];
        v += ch;
        if (ch === '{') depth++;
        else if (ch === '}' && --depth === 0) break;
      }
      out.push({ text: v, quoted: false });
      continue;
    }
    let v = '';
    while (i < s.length && !/\s/.test(s[i])) v += s[i++];
    out.push({ text: v, quoted: false });
  }
  return out;
}

export function parseInstr(instr: string): Instr {
  const s = instr.trim();
  // 公式域：= 后面到第一个开关之前都是表达式
  if (s.startsWith('=')) {
    const k = s.search(/\s\\[@#*]/);
    const expr = (k < 0 ? s.slice(1) : s.slice(1, k)).trim();
    return { name: '=', args: expr ? [expr] : [], sw: k < 0 ? [] : parseInstr(`X ${s.slice(k)}`).sw };
  }
  const ts = tokens(s);
  const out: Instr = { name: ts[0]?.quoted ? '' : (ts[0]?.text ?? ''), args: [], sw: [] };
  for (let i = 1; i < ts.length; i++) {
    const t = ts[i];
    if (!t.quoted && t.text.startsWith('\\') && t.text.length >= 2) {
      const key = t.text.length === 2 ? t.text[1] : t.text.slice(1);
      const k = key === '@' || key === '#' || key === '*' ? key : key.toLowerCase();
      const next = ts[i + 1];
      if (VALUE_SW.has(k) && next && !(next.text.startsWith('\\') && !next.quoted)) {
        out.sw.push({ key: k, value: next.text });
        i++;
      } else out.sw.push({ key: k });
    } else out.args.push(t.text);
  }
  return out;
}

const quote = (v: string) => (v === '' || /[\s"\\]/.test(v) ? `"${v.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"` : v);

/** 拼回域代码（对话框改了选项后用）；日期 / 数字图片总加引号（Word 也这样写） */
export function formatInstr(i: Instr): string {
  const parts = [i.name];
  if (i.name === '=') parts.push(...i.args);
  else for (const a of i.args) parts.push(a.startsWith('{') ? a : quote(a));
  for (const s of i.sw) {
    parts.push(`\\${s.key}`);
    if (s.value !== undefined) parts.push(s.key === '@' || s.key === '#' ? `"${s.value.replace(/"/g, '\\"')}"` : quote(s.value));
  }
  return parts.join(' ');
}

/** 改一个开关（value = null 去掉；true 表示不带值的开关）；\* 只改大小写类那一个，MERGEFORMAT 等其它 \* 保留 */
export function withSwitch(i: Instr, key: string, value: string | true | null): Instr {
  const sw = i.sw.filter((s) => s.key !== key || (key === '*' && !CASE_FMT.has((s.value ?? '').toLowerCase()) && !NUM_STYLE.has((s.value ?? '').toLowerCase())));
  if (value !== null) sw.push(value === true ? { key } : { key, value });
  return { ...i, sw };
}

/** 域名（大写）；公式域为 '=' */
export const fieldName = (instr: string): string => parseInstr(instr).name.toUpperCase();

// ---------- 计算 ----------

export interface FieldEnv {
  now?: Date;
  /** 文件名（含扩展名）和完整路径 */
  fileName?: string;
  filePath?: string;
  numPages?: number;
  /** 书签覆盖的文字（REF）；没有这个书签返回 null */
  bookmark?: (name: string) => string | null;
}

/** 域名 → 说明（编辑域对话框的列表，按这个顺序） */
export const FIELD_INFO: { name: string; label: string; desc: string }[] = [
  { name: 'DATE', label: '日期', desc: '当前日期' },
  { name: 'TIME', label: '时间', desc: '当前时间' },
  { name: 'FILENAME', label: '文件名', desc: '文档的文件名' },
  { name: 'NUMPAGES', label: '总页数', desc: '文档的页数' },
  { name: 'REF', label: '书签引用', desc: '插入书签处的文字' },
  { name: 'SEQ', label: '题注编号', desc: '按顺序编号（图 1、表 1……）' },
  { name: '=', label: '公式', desc: '计算表达式的值' },
  { name: 'QUOTE', label: '文字', desc: '插入指定的文字' },
];

/** 能在编辑器里算出结果的域 */
export const COMPUTABLE = new Set(FIELD_INFO.map((f) => f.name));

const CASE_FMT = new Set(['upper', 'lower', 'firstcap', 'caps']);
const NUM_STYLE = new Set(['arabic', 'roman', 'alphabetic', 'chinesenum3', 'chinesenum2']);

const CN = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
const CN_UP = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];

/** 整数的中文写法（0–9999）：CHINESENUM3 一二三、CHINESENUM2 壹贰叁 */
function chineseNum(n: number, digits: string[], units: string[]): string {
  if (n === 0) return digits[0];
  if (n < 0 || n > 9999) return String(n);
  let s = '';
  let zero = false;
  for (let k = 3; k >= 0; k--) {
    const d = Math.floor(n / 10 ** k) % 10;
    if (d === 0) {
      if (s) zero = true;
      continue;
    }
    if (zero) s += digits[0];
    zero = false;
    s += digits[d] + units[k];
  }
  // 十几：一十 → 十（小写）
  return digits === CN && s.startsWith('一十') ? s.slice(1) : s;
}

function roman(n: number): string {
  if (n <= 0 || n >= 4000) return String(n);
  const t: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, r] of t) while (n >= v) (s += r), (n -= v);
  return s;
}

function alpha(n: number): string {
  if (n <= 0) return String(n);
  // Word：a…z，然后 aa…zz（同一字母重复）
  return String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
}

/** \* 的数字样式（值是整数时）和大小写 */
function applyStar(v: string, fmt: string): string {
  const f = fmt.toLowerCase();
  const n = /^-?\d+$/.test(v.trim()) ? Number(v.trim()) : null;
  if (n !== null) {
    if (f === 'arabic') return String(n);
    if (f === 'roman') return fmt === 'ROMAN' ? roman(n).toUpperCase() : roman(n);
    if (f === 'alphabetic') return fmt === 'ALPHABETIC' ? alpha(n).toUpperCase() : alpha(n);
    if (f === 'chinesenum3') return chineseNum(n, CN, ['', '十', '百', '千']);
    if (f === 'chinesenum2') return chineseNum(n, CN_UP, ['', '拾', '佰', '仟']);
  }
  return CASE_FMT.has(f) ? applyCase(v, fmt) : v;
}

/** 数字的默认显示：最多 10 位小数，去掉末尾的 0 */
const plainNum = (n: number) => String(Math.round(n * 1e10) / 1e10);

/** 公式域的表达式：+ - * / ^ %、括号、数字、书签名，函数 ABS INT ROUND MAX MIN SUM AVERAGE PRODUCT MOD SIGN */
export function evalFormula(expr: string, bookmark: (name: string) => string | null = () => null): number | null {
  const s = expr.replace(/\s+/g, '');
  let i = 0;
  const fail = (): never => {
    throw new Error('formula');
  };
  const peek = () => s[i];
  const atom = (): number => {
    if (peek() === '(') {
      i++;
      const v = add();
      if (s[i++] !== ')') fail();
      return v;
    }
    const num = /^(\d+(?:\.\d*)?|\.\d+)/.exec(s.slice(i));
    if (num) {
      i += num[0].length;
      return Number(num[0]);
    }
    const id = /^[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*/.exec(s.slice(i));
    if (!id) return fail();
    i += id[0].length;
    if (peek() === '(') {
      i++;
      const args: number[] = [];
      if (peek() !== ')') {
        args.push(add());
        while (peek() === ',') {
          i++;
          args.push(add());
        }
      }
      if (s[i++] !== ')') fail();
      const f = id[0].toUpperCase();
      const a0 = args[0] ?? fail();
      switch (f) {
        case 'ABS': return Math.abs(a0);
        case 'INT': return Math.trunc(a0);
        case 'SIGN': return Math.sign(a0);
        case 'ROUND': return Math.round(a0 * 10 ** (args[1] ?? 0)) / 10 ** (args[1] ?? 0);
        case 'MOD': return a0 % (args[1] ?? fail());
        case 'MAX': return Math.max(...args);
        case 'MIN': return Math.min(...args);
        case 'SUM': return args.reduce((x, y) => x + y, 0);
        case 'AVERAGE': return args.reduce((x, y) => x + y, 0) / args.length;
        case 'PRODUCT': return args.reduce((x, y) => x * y, 1);
        default: return fail();
      }
    }
    // 书签名：书签处的文字当数字
    const v = parseNumber(bookmark(id[0]) ?? '');
    return v === null ? fail() : v;
  };
  const unary = (): number => {
    if (peek() === '-') {
      i++;
      return -unary();
    }
    if (peek() === '+') {
      i++;
      return unary();
    }
    let v = atom();
    while (peek() === '%') {
      i++;
      v /= 100;
    }
    return v;
  };
  const pow = (): number => {
    const b = unary();
    if (peek() === '^') {
      i++;
      return b ** pow();
    }
    return b;
  };
  const mul = (): number => {
    let v = pow();
    while (peek() === '*' || peek() === '/') {
      const op = s[i++];
      const r = pow();
      v = op === '*' ? v * r : v / r;
    }
    return v;
  };
  const add = (): number => {
    let v = mul();
    while (peek() === '+' || peek() === '-') {
      const op = s[i++];
      const r = mul();
      v = op === '+' ? v + r : v - r;
    }
    return v;
  };
  try {
    const v = add();
    return i === s.length && Number.isFinite(v) ? v : null;
  } catch {
    return null;
  }
}

/**
 * 算一个域的结果；算不出来返回 null（保留原结果）。seq 是 SEQ 域的计数（按文档顺序算，调用方传同一个 Map）。
 * 嵌套的域（{…}）先算，结果代入。
 */
export function evalField(instr: string, env: FieldEnv, seq: Map<string, number> = new Map(), inner?: string[]): string | null {
  const p = parseInstr(resolveNested(instr, env, seq, inner));
  const name = p.name.toUpperCase();
  const at = p.sw.find((s) => s.key === '@')?.value;
  const num = p.sw.find((s) => s.key === '#')?.value;
  let v: string | null;
  let date = false;
  switch (name) {
    case 'DATE':
    case 'TIME': {
      const d = env.now ?? new Date();
      const two = (n: number) => String(n).padStart(2, '0');
      v = `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())} ${two(d.getHours())}:${two(d.getMinutes())}:${two(d.getSeconds())}`;
      v = formatDate(v, at ?? (name === 'DATE' ? 'yyyy/M/d' : 'H:mm'));
      date = true;
      break;
    }
    case 'FILENAME':
      v = (p.sw.some((s) => s.key === 'p') ? env.filePath : env.fileName) ?? null;
      break;
    case 'NUMPAGES':
      v = env.numPages === undefined ? null : String(env.numPages);
      break;
    case 'QUOTE':
      v = p.args.join(' ');
      break;
    case 'SEQ': {
      const id = p.args[0];
      if (!id) return null;
      const reset = p.sw.find((s) => s.key === 'r')?.value;
      const cur = seq.get(id) ?? 0;
      const n = reset !== undefined && /^\d+$/.test(reset) ? Number(reset) : p.sw.some((s) => s.key === 'c') ? cur : cur + 1;
      seq.set(id, n);
      if (p.sw.some((s) => s.key === 'h')) return '';
      v = String(n);
      break;
    }
    case '=': {
      const r = evalFormula(p.args[0] ?? '', env.bookmark);
      if (r === null) return null;
      v = plainNum(r);
      break;
    }
    case 'REF':
      v = p.args[0] ? (env.bookmark?.(p.args[0]) ?? null) : null;
      break;
    default:
      // 只写书签名也是引用（Word：{ 书签名 }）
      v = p.name && !p.name.startsWith('\\') && env.bookmark ? env.bookmark(p.name) : null;
      break;
  }
  if (v === null) return null;
  if (num !== undefined && !date) v = formatNumber(v, num);
  for (const s of p.sw) if (s.key === '*' && s.value) v = applyStar(v, s.value);
  return v;
}

/** 把域代码里嵌套的 {…} 换成它们的结果（先算里层）；inner 收集各嵌套域的结果，按左花括号出现的顺序 */
function resolveNested(code: string, env: FieldEnv, seq: Map<string, number>, inner?: string[]): string {
  let out = '';
  let i = 0;
  while (i < code.length) {
    if (code[i] !== '{') {
      out += code[i++];
      continue;
    }
    let depth = 0;
    let j = i;
    for (; j < code.length; j++) {
      if (code[j] === '{') depth++;
      else if (code[j] === '}' && --depth === 0) break;
    }
    if (j >= code.length) return out + code.slice(i);
    const slot = inner ? inner.push('') - 1 : -1;
    const v = evalField(code.slice(i + 1, j), env, seq, inner) ?? '';
    if (inner) inner[slot] = v;
    out += v;
    i = j + 1;
  }
  return out;
}

// ---------- 更新、改代码 ----------

/** 把域的结果换成 text：用原结果第一段文字的格式（没有就用开始标记的格式） */
function replaceResult(p: Paragraph, f: FieldSpan, text: string) {
  const inner = p.runs.slice(f.begin + 1, f.end);
  const like = inner.find((r) => !r.fc && !r.opaque && !r.image && !r.box && !r.mark && !r.field)?.props ?? p.runs[f.begin].props;
  const run: Run = { text, props: { ...like } };
  p.runs.splice(f.begin + 1, f.end - f.begin - 1, ...(text ? [run] : []));
}

export interface UpdateResult {
  /** 结果换了的域 */
  updated: number;
  /** 算不出来、保留原结果的域名 */
  kept: string[];
  /** 锁定、没更新的域 */
  locked: number;
}

/**
 * 更新域（F9）：ids 给了只更新这些，否则更新所有域。按文档顺序计算（SEQ 计数要从头数，不更新的 SEQ 也算进去）。
 * 只改域结果；结果没变的不动（不算修改）。
 */
export function updateFields(doc: Doc, env: FieldEnv, ids?: Set<string>): UpdateResult {
  const res: UpdateResult = { updated: 0, kept: [], locked: 0 };
  const seq = new Map<string, number>();
  const e: FieldEnv = { ...env, bookmark: env.bookmark ?? ((n) => bookmarkText(doc, n)) };
  eachParagraph(doc, (p) => {
    const changes: { f: FieldSpan; text: string; inner: string[] }[] = [];
    for (const f of fieldSpans(p)) {
      const target = !ids || ids.has(f.id);
      const name = fieldName(f.instr);
      if (!target && name !== 'SEQ') continue;
      const inner: string[] = [];
      const v = evalField(f.instr, e, seq, inner);
      if (!target) continue;
      if (f.locked) {
        res.locked++;
        continue;
      }
      if (v === null) {
        res.kept.push(name || '（空）');
        continue;
      }
      // 套在别的域结果里的不单独换（外层换了结果它也就没了）
      if (f.depth > 0) continue;
      const old = p.runs.slice(f.begin + 1, f.end).map((r) => r.text).join('');
      // 按代码生成的域（没有原始 XML）还要记下嵌套域的结果
      const fc = p.runs[f.begin].fc!;
      const innerChanged = !fc.raw && inner.length > 0 && JSON.stringify(inner) !== JSON.stringify(fc.inner ?? []);
      if (old !== v || innerChanged) changes.push({ f, text: v, inner });
    }
    for (const c of changes.reverse()) {
      const fc = p.runs[c.f.begin].fc!;
      if (!fc.raw && c.inner.length) p.runs[c.f.begin] = { ...p.runs[c.f.begin], fc: { ...fc, inner: c.inner } };
      replaceResult(p, c.f, c.text);
      res.updated++;
    }
  });
  return res;
}

/** 改域代码（编辑域）：原始 XML 作废，之后按新代码写；不更新结果（调用方接着更新） */
export function setFieldInstr(p: Paragraph, id: string, instr: string): boolean {
  const i = p.runs.findIndex((r) => r.fc?.kind === 'begin' && r.fc.id === id);
  if (i < 0) return false;
  const fc = p.runs[i].fc!;
  p.runs[i] = { ...p.runs[i], fc: { kind: 'begin', id: fc.id, instr: instr.trim(), ...(fc.locked ? { locked: true } : {}) } };
  return true;
}

/** 选区碰到的域（选区收拢时是光标所在的域） */
export function fieldIdsInRange(doc: Doc, a: Pos, b: Pos): string[] {
  const out: string[] = [];
  for (const r of paragraphsInRange(doc, a, b)) {
    const p = getPara(doc, { story: a.story, path: r.path });
    for (const f of fieldSpans(p)) {
      const hit = r.from === r.to ? f.from < r.from && r.from <= f.to : f.from < r.to && r.from < f.to;
      if (hit) out.push(f.id);
    }
  }
  return out;
}

/** 文档里所有域 */
export function allFieldIds(doc: Doc): string[] {
  const out: string[] = [];
  eachParagraph(doc, (p) => {
    for (const f of fieldSpans(p)) out.push(f.id);
  });
  return out;
}

/** 按 id 找域所在的段落 */
export function findField(doc: Doc, id: string): { p: Paragraph; f: FieldSpan } | null {
  let hit: { p: Paragraph; f: FieldSpan } | null = null;
  eachParagraph(doc, (p) => {
    if (hit) return;
    const f = fieldSpans(p).find((x) => x.id === id);
    if (f) hit = { p, f };
  });
  return hit;
}

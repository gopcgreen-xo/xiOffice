/**
 * 公式里的单元格引用：找出引用并改写（不求值）。
 * 用于共享公式展开、填充 / 复制粘贴时相对引用平移、插入 / 删除行列时引用跟着移动（被删掉的变 #REF!）。
 *
 * 识别的引用：A1、$A$1、A1:B5、A:A、$1:$3，可带工作表前缀 Sheet1!、'表 名'!（'' 表示一个单引号）。
 * 字符串 "..."、结构化引用 Table1[...]、外部工作簿 [1] 里的内容原样保留。
 */

export const MAX_ROW = 1048575;
export const MAX_COL = 16383;

export interface RefPart {
  /** 行 / 列（从 0 开始）；整列引用没有 r，整行引用没有 c */
  r?: number;
  c?: number;
  absR?: boolean;
  absC?: boolean;
}

export interface Ref {
  /** 工作表名（已去掉引号）；没有前缀时为 undefined */
  sheet?: string;
  /** 单格只有 a；区域、整行、整列有 b */
  a: RefPart;
  b?: RefPart;
  /** 原文里的表名前缀（含 !），表名没变时原样写回 */
  prefix?: string;
}

const SHEET = String.raw`(?:'((?:[^']|'')+)'|([\p{L}_][\p{L}\p{N}_.]*))!`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})`;
const COL = String.raw`(\$?)([A-Za-z]{1,3})`;
const ROW = String.raw`(\$?)(\d{1,7})`;
/** 依次尝试：区域 / 单格、整列、整行；后面不能紧跟标识符字符、( 或 !（函数名 LOG10( 之类） */
const REF_RE = new RegExp(
  String.raw`^(?:${SHEET})?(?:${CELL}(?::${CELL})?|${COL}:${COL}|${ROW}:${ROW})(?![\p{L}\p{N}_.(!\[])`,
  'u',
);
const WORD_RE = /^[\p{L}\p{N}_.\\]+/u;

function colNum(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function colStr(i: number): string {
  let s = '';
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function parseMatch(m: RegExpMatchArray): Ref | null {
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const r = parseParts(m, sheet);
  if (r && sheet !== undefined) r.prefix = m[0].slice(0, m[0].indexOf(m[1] !== undefined ? "'!" : '!') + (m[1] !== undefined ? 2 : 1));
  return r;
}

function parseParts(m: RegExpMatchArray, sheet: string | undefined): Ref | null {
  if (m[4] !== undefined) {
    const a: RefPart = { absC: !!m[3], c: colNum(m[4]), absR: !!m[5], r: Number(m[6]) - 1 };
    if (a.c! > MAX_COL || a.r! < 0 || a.r! > MAX_ROW) return null;
    if (m[8] === undefined) return { sheet, a };
    const b: RefPart = { absC: !!m[7], c: colNum(m[8]), absR: !!m[9], r: Number(m[10]) - 1 };
    if (b.c! > MAX_COL || b.r! < 0 || b.r! > MAX_ROW) return null;
    return { sheet, a, b };
  }
  if (m[12] !== undefined) {
    const a: RefPart = { absC: !!m[11], c: colNum(m[12]) };
    const b: RefPart = { absC: !!m[13], c: colNum(m[14]) };
    if (a.c! > MAX_COL || b.c! > MAX_COL) return null;
    return { sheet, a, b };
  }
  const a: RefPart = { absR: !!m[15], r: Number(m[16]) - 1 };
  const b: RefPart = { absR: !!m[17], r: Number(m[18]) - 1 };
  if (a.r! < 0 || b.r! < 0 || a.r! > MAX_ROW || b.r! > MAX_ROW) return null;
  return { sheet, a, b };
}

/** 工作表名写进公式：只有字母数字下划线（且不像单元格地址）时不加引号 */
export function quoteSheet(name: string): string {
  if (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(TRUE|FALSE)$/i.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

function partStr(p: RefPart): string {
  return `${p.c !== undefined ? `${p.absC ? '$' : ''}${colStr(p.c)}` : ''}${p.r !== undefined ? `${p.absR ? '$' : ''}${p.r + 1}` : ''}`;
}

export function refStr(ref: Ref): string {
  const prefix = ref.sheet === undefined ? '' : ref.prefix ?? `${quoteSheet(ref.sheet)}!`;
  return prefix + partStr(ref.a) + (ref.b ? `:${partStr(ref.b)}` : '');
}

/**
 * 逐个引用改写公式：fn 返回新的引用、null（变成 #REF!）或 undefined（不变）。
 * 公式里其它内容（函数名、字符串、数字、运算符、空格）原样保留。
 */
export function mapRefs(formula: string, fn: (ref: Ref, start: number, end: number) => Ref | null | undefined): string {
  let out = '';
  let i = 0;
  const n = formula.length;
  while (i < n) {
    const ch = formula[i];
    if (ch === '"') {
      // 字符串："" 是转义的引号
      let j = i + 1;
      while (j < n) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') j += 2;
          else break;
        } else j++;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === '[') {
      // 结构化引用 / 外部工作簿编号：到配对的 ] 为止
      let depth = 0;
      let j = i;
      for (; j < n; j++) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']' && --depth === 0) break;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "'" || ch === '$' || /[\p{L}\p{N}_]/u.test(ch)) {
      const m = formula.slice(i).match(REF_RE);
      const ref = m ? parseMatch(m) : null;
      if (m && ref) {
        const next = fn(ref, i, i + m[0].length);
        const keep = next && next.sheet === ref.sheet ? { ...next, prefix: ref.prefix } : next ? { ...next, prefix: undefined } : next;
        out += keep === undefined ? m[0] : keep === null ? `${ref.prefix ?? ''}#REF!` : refStr(keep);
        i += m[0].length;
        continue;
      }
      if (ch === "'") {
        // 引号里的工作表名（后面不是合法引用）：原样跳过
        let j = i + 1;
        while (j < n && !(formula[j] === "'" && formula[j + 1] !== "'")) j += formula[j] === "'" ? 2 : 1;
        out += formula.slice(i, j + 1);
        i = j + 1;
        continue;
      }
      // 函数名、名称、数字：整个词跳过，不会从词中间开始匹配
      const w = formula.slice(i).match(WORD_RE);
      const len = w ? w[0].length : 1;
      out += formula.slice(i, i + len);
      i += len;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

/** 从 src[i] 开始的一个引用（调用方保证 i 是一个词的开头） */
export function matchRef(src: string, i: number): { ref: Ref; len: number } | null {
  const m = src.slice(i).match(REF_RE);
  const ref = m ? parseMatch(m) : null;
  return m && ref ? { ref, len: m[0].length } : null;
}

/** 公式里的所有引用（求值、引用高亮用） */
export function listRefs(formula: string): Ref[] {
  return refSpans(formula).map((x) => x.ref);
}

/** 公式里的引用及其在文字中的位置 [start, end) */
export function refSpans(formula: string): { ref: Ref; start: number; end: number }[] {
  const out: { ref: Ref; start: number; end: number }[] = [];
  mapRefs(formula, (ref, start, end) => {
    out.push({ ref, start, end });
    return undefined;
  });
  return out;
}

/** F4：引用在 A1 → $A$1 → A$1 → $A1 → A1 之间切换（区域两端一起） */
export function cycleAbsolute(ref: Ref): Ref {
  const p = ref.a;
  const state = p.absC && p.absR ? 1 : !p.absC && p.absR ? 2 : p.absC && !p.absR ? 3 : 0;
  const next = (state + 1) % 4;
  const set = (x: RefPart): RefPart => ({ ...x, absC: next === 1 || next === 3, absR: next === 1 || next === 2 });
  return { ...ref, a: set(ref.a), b: ref.b ? set(ref.b) : undefined };
}

/** 光标位置能插入引用：在 = ( , 运算符之后（中间可以有空格） */
export function canInsertRef(text: string, caret: number): boolean {
  if (!text.startsWith('=')) return false;
  const before = text.slice(0, caret).replace(/\s+$/, '');
  return before.length > 0 && /[=(,+\-*/^&<>:;{]$/.test(before);
}

/**
 * 复制 / 填充：相对引用平移 dr 行 dc 列，绝对引用（$）不动；移出表格范围的引用变 #REF!。
 */
export function moveFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  const mv = (p: RefPart): RefPart | null => {
    const q = { ...p };
    if (q.r !== undefined && !q.absR) q.r += dr;
    if (q.c !== undefined && !q.absC) q.c += dc;
    if ((q.r !== undefined && (q.r < 0 || q.r > MAX_ROW)) || (q.c !== undefined && (q.c < 0 || q.c > MAX_COL))) return null;
    return q;
  };
  return mapRefs(formula, (ref) => {
    const a = mv(ref.a);
    const b = ref.b ? mv(ref.b) : undefined;
    if (!a || b === null) return null;
    return { sheet: ref.sheet, a, b };
  });
}

export interface ShiftSpec {
  /** 被插入 / 删除行列的工作表名 */
  sheet: string;
  axis: 'row' | 'col';
  /** 从第 at 行 / 列开始 */
  at: number;
  /** 正数插入，负数删除 */
  n: number;
  /**
   * 插入 / 删除单元格：只移动另一方向上落在 [band[0], band[1]] 里的格子（下移时为列范围，右移时为行范围）。
   * 引用的另一方向范围完全在 band 里才跟着移动，否则不变（同 Excel）。
   */
  band?: [number, number];
}

/** 区域（另一方向的 [lo, hi]，undefined = 整行 / 整列）是否完全在 band 里 */
export function withinBand(lo: number | undefined, hi: number | undefined, band: [number, number] | undefined): boolean {
  if (!band) return true;
  if (lo === undefined || hi === undefined) return false;
  return Math.min(lo, hi) >= band[0] && Math.max(lo, hi) <= band[1];
}

const sameSheet = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

/** 一维区间 [lo, hi] 在插入 / 删除后的位置；整段被删掉返回 null */
export function shiftSpan(lo: number, hi: number, at: number, n: number): [number, number] | null {
  if (n > 0) return [lo >= at ? lo + n : lo, hi >= at ? hi + n : hi];
  const end = at - n; // 删除 [at, end)
  const map = (x: number) => (x < at ? x : x >= end ? x + n : at);
  if (lo >= at && hi < end) return null;
  const nlo = lo >= at && lo < end ? at : map(lo);
  const nhi = hi >= at && hi < end ? at - 1 : map(hi);
  return nhi < nlo ? null : [nlo, nhi];
}

/**
 * 插入 / 删除行列后改写公式：指向该表的引用（带同名前缀，或公式本身就在该表里且不带前缀）跟着移动；
 * 绝对、相对引用都移动（引用的是那个格子本身）；整个被删除的引用变 #REF!，区域被删掉一部分时缩小。
 * owner 为公式所在工作表名（定义名称等不属于任何表时传 undefined）。
 */
export function shiftFormula(formula: string, owner: string | undefined, spec: ShiftSpec): string {
  const key = spec.axis === 'row' ? 'r' : 'c';
  return mapRefs(formula, (ref) => {
    const target = ref.sheet ?? owner;
    if (target === undefined || !sameSheet(target, spec.sheet)) return undefined;
    const lo = ref.a[key];
    if (lo === undefined) return undefined; // 整行引用不受插入列影响，反之亦然
    const hi = ref.b ? ref.b[key]! : lo;
    const other = spec.axis === 'row' ? 'c' : 'r';
    if (!withinBand(ref.a[other], ref.b ? ref.b[other] : ref.a[other], spec.band)) return undefined;
    const span = shiftSpan(Math.min(lo, hi), Math.max(lo, hi), spec.at, spec.n);
    if (!span) return null;
    if (span[0] === Math.min(lo, hi) && span[1] === Math.max(lo, hi)) return undefined;
    const a = { ...ref.a, [key]: span[0] };
    const b = ref.b ? { ...ref.b, [key]: span[1] } : undefined;
    if (span[1] > (key === 'r' ? MAX_ROW : MAX_COL)) return null;
    return { sheet: ref.sheet, a, b };
  });
}

/** 工作表改名后，带旧表名前缀的引用改成新表名 */
export function renameSheetRefs(formula: string, from: string, to: string): string {
  return mapRefs(formula, (ref) => (ref.sheet !== undefined && sameSheet(ref.sheet, from) ? { ...ref, sheet: to } : undefined));
}

/** sqref（空格分隔的区域列表，如 "A1:B5 D3"）插入 / 删除行列后的结果；全被删掉返回空串 */
export function shiftSqref(sqref: string, spec: Omit<ShiftSpec, 'sheet'>): string {
  return sqref
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => {
      const out = shiftFormula(part, '_', { ...spec, sheet: '_' });
      return out.includes('#REF!') ? '' : out;
    })
    .filter(Boolean)
    .join(' ');
}

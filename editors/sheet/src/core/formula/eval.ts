/**
 * 公式求值与自动重算。
 *
 * 值：数字、文字、布尔、错误（FErr）、空（null）；区域引用是惰性的 RangeV，数组常量 / 数组运算结果是 ArrV。
 * 类型转换同 Excel：算术里文字按输入识别转数字（转不了是 #VALUE!），空当 0；比较时 数字 < 文字 < 布尔，文字不区分大小写；
 * SUM 等函数直接写的参数会转换，区域里的文字 / 布尔忽略。
 *
 * 重算（Engine）：改了哪些格子 → 找出（间接）引用它们的公式 → 按依赖顺序（拓扑排序，不递归）逐个重算；循环引用显示 #CIRC!。
 * 打开文件时不重算，直接用文件里的缓存值。含不支持的函数 / 语法的公式保留原值，记在 unsupported 里。
 */
import { formatGeneral, formatNumber } from '../numfmt';
import { parseInput, dateSerial } from '../input';
import { colIndex, colName } from '@wz/editor-shared/xlsx';
import { refreshText, type SCell, type Workbook } from '../workbook';
import { parseFormula, type Node } from './parse';
import type { Ref } from './refs';
import { moreFunctions } from './funcs2';

export class FErr {
  constructor(readonly code: string) {}
}
export type Scalar = number | string | boolean | null | FErr;
export interface RangeV {
  kind: 'range';
  si: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}
export interface ArrV {
  kind: 'arr';
  data: Scalar[][];
}
export type V = Scalar | RangeV | ArrV;

const E = {
  div0: () => new FErr('#DIV/0!'),
  value: () => new FErr('#VALUE!'),
  ref: () => new FErr('#REF!'),
  name: () => new FErr('#NAME?'),
  num: () => new FErr('#NUM!'),
  na: () => new FErr('#N/A'),
};
/** 能写进 xlsx 的错误值（#CIRC! 只在编辑器里显示） */
export const FILE_ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);

const isErr = (v: unknown): v is FErr => v instanceof FErr;
const isRange = (v: V): v is RangeV => typeof v === 'object' && v !== null && !isErr(v) && v.kind === 'range';
const isArr = (v: V): v is ArrV => typeof v === 'object' && v !== null && !isErr(v) && v.kind === 'arr';

export interface Ctx {
  book: Workbook;
  si: number;
  r: number;
  c: number;
  /** 读一个格子的值 */
  get(si: number, r: number, c: number): Scalar;
  /** 定义名称 → 语法树（找不到返回 null） */
  name(name: string): Node | null;
  /** 表名 → 序号（找不到 -1） */
  sheetIndex(name: string): number;
}

// ---------- 转换 ----------

function rangeDims(v: RangeV, ctx: Ctx): { r1: number; c1: number } {
  // 整列 / 整行引用只取到数据范围，避免遍历一百万行
  const s = ctx.book.sheets[v.si];
  return { r1: Math.min(v.r1, Math.max(v.r0, (s?.rows ?? 0) - 1)), c1: Math.min(v.c1, Math.max(v.c0, (s?.cols ?? 0) - 1)) };
}

function toArr(v: V, ctx: Ctx): Scalar[][] {
  if (isArr(v)) return v.data;
  if (isRange(v)) {
    const { r1, c1 } = rangeDims(v, ctx);
    const out: Scalar[][] = [];
    for (let r = v.r0; r <= r1; r++) {
      const row: Scalar[] = [];
      for (let c = v.c0; c <= c1; c++) row.push(ctx.get(v.si, r, c));
      out.push(row);
    }
    return out;
  }
  return [[v]];
}

/** 区域 → 单值（隐式交集：与公式同行 / 同列的那一格） */
export function scalar(v: V, ctx: Ctx): Scalar {
  if (isArr(v)) return v.data[0]?.[0] ?? null;
  if (!isRange(v)) return v;
  if (v.r0 === v.r1 && v.c0 === v.c1) return ctx.get(v.si, v.r0, v.c0);
  if (v.c0 === v.c1 && ctx.r >= v.r0 && ctx.r <= v.r1) return ctx.get(v.si, ctx.r, v.c0);
  if (v.r0 === v.r1 && ctx.c >= v.c0 && ctx.c <= v.c1) return ctx.get(v.si, v.r0, ctx.c);
  return E.value();
}

function textToNum(s: string, ctx: Ctx): number | FErr {
  if (s.trim() === '') return E.value();
  const p = parseInput(s.trim(), 'General', ctx.book.date1904);
  return p.kind === 'value' && typeof p.v === 'number' ? p.v : E.value();
}

export function toNum(v: Scalar, ctx: Ctx): number | FErr {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) return v;
  return textToNum(v, ctx);
}

export function toStr(v: Scalar): string | FErr {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isErr(v)) return v;
  return formatGeneral(v);
}

function toBool(v: Scalar): boolean | FErr {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isErr(v)) return v;
  const u = v.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return E.value();
}

/** 比较：数字 < 文字 < 布尔；空按对方类型当 0 / "" / FALSE；文字不区分大小写 */
export function compare(a: Scalar, b: Scalar): number {
  const rank = (x: Scalar) => (typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : typeof x === 'boolean' ? 2 : 3);
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number') return a - (b as number);
  if (typeof a === 'string') {
    const x = a.toLowerCase();
    const y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return Number(a) - Number(b);
}

// ---------- 求值 ----------

function refToRange(ref: Ref, ctx: Ctx): RangeV | FErr {
  const si = ref.sheet === undefined ? ctx.si : ctx.sheetIndex(ref.sheet);
  if (si < 0) return E.ref();
  const a = ref.a;
  const b = ref.b ?? ref.a;
  const r0 = a.r ?? 0;
  const r1 = b.r ?? 1048575;
  const c0 = a.c ?? 0;
  const c1 = b.c ?? 16383;
  return { kind: 'range', si, r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
}

/** 两个值逐元素运算（任一边是区域 / 数组时按 Excel 的数组规则广播） */
function elementwise(a: V, b: V, ctx: Ctx, f: (x: Scalar, y: Scalar) => Scalar): V {
  if (!isRange(a) && !isArr(a) && !isRange(b) && !isArr(b)) return f(a, b);
  const A = toArr(a, ctx);
  const B = toArr(b, ctx);
  const rows = Math.max(A.length, B.length);
  const cols = Math.max(A[0]?.length ?? 0, B[0]?.length ?? 0);
  const at = (M: Scalar[][], r: number, c: number): Scalar => {
    const rr = M.length === 1 ? 0 : r;
    const cc = (M[0]?.length ?? 0) === 1 ? 0 : c;
    return rr < M.length && cc < (M[rr]?.length ?? 0) ? M[rr][cc] : E.na();
  };
  const data: Scalar[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) row.push(f(at(A, r, c), at(B, r, c)));
    data.push(row);
  }
  return { kind: 'arr', data };
}

function arith(op: string, x: Scalar, y: Scalar, ctx: Ctx): Scalar {
  if (op === '&') {
    const a = toStr(x);
    const b = toStr(y);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    return a + b;
  }
  if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    const d = compare(x, y);
    return op === '=' ? d === 0 : op === '<>' ? d !== 0 : op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0;
  }
  const a = toNum(x, ctx);
  if (isErr(a)) return a;
  const b = toNum(y, ctx);
  if (isErr(b)) return b;
  let out: number;
  switch (op) {
    case '+':
      out = a + b;
      break;
    case '-':
      out = a - b;
      break;
    case '*':
      out = a * b;
      break;
    case '/':
      if (b === 0) return E.div0();
      out = a / b;
      break;
    case '^':
      if (a === 0 && b === 0) return E.num();
      out = Math.pow(a, b);
      break;
    default:
      return E.value();
  }
  if (!Number.isFinite(out)) return E.num();
  // 去掉浮点误差（0.1 + 0.2 = 0.3，同 Excel 的 15 位有效数字）
  return Number(out.toPrecision(15));
}

export function evaluate(node: Node, ctx: Ctx): V {
  switch (node.t) {
    case 'num':
      return node.v;
    case 'str':
      return node.v;
    case 'bool':
      return node.v;
    case 'err':
      return new FErr(node.v);
    case 'missing':
      return null;
    case 'ref':
      return refToRange(node.ref, ctx);
    case 'name': {
      const n = ctx.name(node.name);
      return n ? evaluate(n, ctx) : E.name();
    }
    case 'arr':
      return { kind: 'arr', data: node.rows.map((row) => row.map((x) => scalar(evaluate(x, ctx), ctx))) };
    case 'un': {
      const a = evaluate(node.a, ctx);
      if (node.op === '+') return a;
      return elementwise(a, null, ctx, (x) => {
        const n = toNum(x, ctx);
        if (isErr(n)) return n;
        return node.op === '-' ? -n : Number((n / 100).toPrecision(15));
      });
    }
    case 'bin': {
      if (node.op === ':') {
        const a = evaluate(node.a, ctx);
        const b = evaluate(node.b, ctx);
        if (isErr(a)) return a;
        if (isErr(b)) return b;
        if (!isRange(a) || !isRange(b) || a.si !== b.si) return E.ref();
        return { kind: 'range', si: a.si, r0: Math.min(a.r0, b.r0), c0: Math.min(a.c0, b.c0), r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1) };
      }
      const a = evaluate(node.a, ctx);
      const b = evaluate(node.b, ctx);
      // 单格区域按单值算（=A1+1）
      const sa = isRange(a) && a.r0 === a.r1 && a.c0 === a.c1 ? ctx.get(a.si, a.r0, a.c0) : a;
      const sb = isRange(b) && b.r0 === b.r1 && b.c0 === b.c1 ? ctx.get(b.si, b.r0, b.c0) : b;
      return elementwise(sa, sb, ctx, (x, y) => arith(node.op, x, y, ctx));
    }
    case 'fn': {
      const f = FUNCS[node.name];
      if (!f) return E.name();
      if (f.lazy) return f.lazy(node.args, ctx);
      return f.fn!(node.args.map((x) => evaluate(x, ctx)), ctx);
    }
  }
}

// ---------- 函数 ----------

interface Fn {
  fn?: (args: V[], ctx: Ctx) => V;
  /** 需要按条件只算部分参数的函数（IF、IFERROR 等） */
  lazy?: (args: Node[], ctx: Ctx) => V;
}

/** 区域 / 数组 / 单值里的所有值；fromRef 表示来自引用或数组（文字、布尔不参与求和） */
function* values(args: V[], ctx: Ctx): Generator<{ v: Scalar; fromRef: boolean }> {
  for (const a of args) {
    if (isRange(a) || isArr(a)) for (const row of toArr(a, ctx)) for (const v of row) yield { v, fromRef: true };
    else yield { v: a, fromRef: false };
  }
}

/** 参与数值统计的数字（Excel：引用里的文字 / 布尔 / 空忽略；直接写的参数转换） */
function numbers(args: V[], ctx: Ctx): number[] | FErr {
  const out: number[] = [];
  for (const { v, fromRef } of values(args, ctx)) {
    if (isErr(v)) return v;
    if (typeof v === 'number') out.push(v);
    else if (!fromRef && v !== null) {
      const n = toNum(v, ctx);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}

const sum = (xs: number[]) => Number(xs.reduce((a, b) => a + b, 0).toPrecision(15));

function num1(v: V, ctx: Ctx): number | FErr {
  return toNum(scalar(v, ctx), ctx);
}
function str1(v: V, ctx: Ctx): string | FErr {
  return toStr(scalar(v, ctx));
}

/** 通配符 * ? 与 ~ 转义 → 正则 */
function wildcard(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'i');
}

/** COUNTIF 一类的条件：数字、"文字"、">=5"、"<>"、带通配符的文字 */
function criteria(c: Scalar, ctx: Ctx): (v: Scalar) => boolean {
  if (typeof c === 'number') return (v) => typeof v === 'number' ? v === c : typeof v === 'string' && v.trim() !== '' && toNum(v, ctx) === c;
  if (typeof c === 'boolean') return (v) => v === c;
  if (c === null) return (v) => v === null || v === '';
  if (isErr(c)) return (v) => isErr(v) && v.code === c.code;
  const m = c.match(/^(<=|>=|<>|<|>|=)?([\s\S]*)$/)!;
  const op = m[1] ?? '=';
  const rest = m[2];
  const n = rest.trim() === '' ? NaN : Number(rest.replace(/,/g, ''));
  const asNum = Number.isFinite(n) ? n : (() => {
    const p = parseInput(rest, 'General', ctx.book.date1904);
    return p.kind === 'value' && typeof p.v === 'number' ? p.v : NaN;
  })();
  if (rest === '') {
    if (op === '=') return (v) => v === null || v === '';
    if (op === '<>') return (v) => v !== null && v !== '';
  }
  if (!Number.isNaN(asNum)) {
    return (v) => {
      if (typeof v !== 'number') return op === '<>' ? true : op === '=' ? typeof v === 'string' && v.trim() !== '' && toNum(v, ctx) === asNum : false;
      return op === '=' ? v === asNum : op === '<>' ? v !== asNum : op === '<' ? v < asNum : op === '>' ? v > asNum : op === '<=' ? v <= asNum : v >= asNum;
    };
  }
  const upper = rest.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE';
    return (v) => (op === '<>' ? v !== b : v === b);
  }
  if (op === '=' || op === '<>') {
    const re = /[*?~]/.test(rest) ? wildcard(rest) : null;
    const low = rest.toLowerCase();
    const eq = (v: Scalar) => typeof v === 'string' && (re ? re.test(v) : v.toLowerCase() === low);
    return op === '=' ? eq : (v) => !eq(v);
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    const d = compare(v, rest);
    return op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0;
  };
}

/** SUMIFS / COUNTIFS 的 (区域, 条件) 对 → 满足全部条件的位置 */
function ifsMask(pairs: V[], ctx: Ctx, shape?: Scalar[][]): boolean[][] | FErr {
  let mask: boolean[][] | null = null;
  for (let i = 0; i + 1 < pairs.length; i += 2) {
    const range = toArr(pairs[i], ctx);
    const cond = scalar(pairs[i + 1], ctx);
    const test = criteria(cond, ctx);
    if (shape && (range.length !== shape.length || (range[0]?.length ?? 0) !== (shape[0]?.length ?? 0))) return E.value();
    if (mask && (range.length !== mask.length || (range[0]?.length ?? 0) !== (mask[0]?.length ?? 0))) return E.value();
    const m = range.map((row) => row.map((v) => test(v)));
    mask = mask ? mask.map((row, r) => row.map((x, c) => x && m[r][c])) : m;
  }
  return mask ?? [];
}

function roundTo(x: number, d: number, mode: 'round' | 'up' | 'down'): number {
  const sign = x < 0 ? -1 : 1;
  const abs = Math.abs(x);
  const shifted = Number(`${abs}e${d}`);
  const r = mode === 'round' ? Math.round(Number(shifted.toPrecision(15))) : mode === 'up' ? Math.ceil(Number(shifted.toPrecision(15))) : Math.floor(Number(shifted.toPrecision(15)));
  return sign * Number(`${r}e${-d}`);
}

/** 日期序列号 → 年月日（1900 日期系统，含 Excel 把 1900 当闰年的规则） */
function ymd(serial: number, date1904 = false): { y: number; m: number; d: number } {
  let s = Math.floor(serial);
  if (!date1904) {
    if (s === 60) return { y: 1900, m: 2, d: 29 };
    if (s < 60) s += 1;
  }
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const dt = new Date(base + s * 86400000);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}

function lookupIndex(needle: Scalar, list: Scalar[], approx: boolean): number {
  if (!approx) {
    const re = typeof needle === 'string' && /[*?~]/.test(needle) ? wildcard(needle) : null;
    return list.findIndex((v) => (re ? typeof v === 'string' && re.test(v) : v !== null && compare(v, needle) === 0 && typeof v === typeof needle));
  }
  // 近似匹配：已排序的列表里最后一个 ≤ needle 的（同类型）
  let best = -1;
  for (let i = 0; i < list.length; i++) {
    const v = list[i];
    if (v === null || typeof v !== typeof needle) continue;
    if (compare(v, needle) <= 0) best = i;
    else break;
  }
  return best;
}

const flat = (m: Scalar[][]) => m.flat();

function err1(...xs: unknown[]): FErr | null {
  for (const x of xs) if (isErr(x)) return x;
  return null;
}

const FUNCS: Record<string, Fn> = {
  // ---- 数学 ----
  SUM: { fn: (a, ctx) => {
    const xs = numbers(a, ctx);
    return isErr(xs) ? xs : sum(xs);
  } },
  PRODUCT: { fn: (a, ctx) => {
    const xs = numbers(a, ctx);
    return isErr(xs) ? xs : xs.length ? Number(xs.reduce((x, y) => x * y, 1).toPrecision(15)) : 0;
  } },
  AVERAGE: { fn: (a, ctx) => {
    const xs = numbers(a, ctx);
    if (isErr(xs)) return xs;
    return xs.length ? Number((sum(xs) / xs.length).toPrecision(15)) : E.div0();
  } },
  MAX: { fn: (a, ctx) => {
    const xs = numbers(a, ctx);
    return isErr(xs) ? xs : xs.length ? Math.max(...xs) : 0;
  } },
  MIN: { fn: (a, ctx) => {
    const xs = numbers(a, ctx);
    return isErr(xs) ? xs : xs.length ? Math.min(...xs) : 0;
  } },
  COUNT: { fn: (a, ctx) => {
    let n = 0;
    for (const { v, fromRef } of values(a, ctx)) if (typeof v === 'number' || (!fromRef && v !== null && !isErr(v) && !isErr(toNum(v, ctx)))) n++;
    return n;
  } },
  COUNTA: { fn: (a, ctx) => {
    let n = 0;
    for (const { v, fromRef } of values(a, ctx)) if (v !== null || !fromRef) n++;
    return n;
  } },
  COUNTBLANK: { fn: (a, ctx) => {
    let n = 0;
    for (const { v } of values(a, ctx)) if (v === null || v === '') n++;
    return n;
  } },
  COUNTIF: { fn: (a, ctx) => {
    const test = criteria(scalar(a[1] ?? null, ctx), ctx);
    return flat(toArr(a[0], ctx)).filter(test).length;
  } },
  COUNTIFS: { fn: (a, ctx) => {
    const m = ifsMask(a, ctx);
    return isErr(m) ? m : m.flat().filter(Boolean).length;
  } },
  SUMIF: { fn: (a, ctx) => {
    const range = toArr(a[0], ctx);
    const test = criteria(scalar(a[1] ?? null, ctx), ctx);
    const target = a[2] !== undefined && a[2] !== null ? sumRangeLike(a[2], range, ctx) : range;
    let s = 0;
    range.forEach((row, r) => row.forEach((v, c) => {
      const t = target[r]?.[c];
      if (test(v) && typeof t === 'number') s += t;
    }));
    return Number(s.toPrecision(15));
  } },
  SUMIFS: { fn: (a, ctx) => {
    const target = toArr(a[0], ctx);
    const m = ifsMask(a.slice(1), ctx, target);
    if (isErr(m)) return m;
    let s = 0;
    target.forEach((row, r) => row.forEach((v, c) => m[r]?.[c] && typeof v === 'number' && (s += v)));
    return Number(s.toPrecision(15));
  } },
  AVERAGEIF: { fn: (a, ctx) => {
    const range = toArr(a[0], ctx);
    const test = criteria(scalar(a[1] ?? null, ctx), ctx);
    const target = a[2] !== undefined && a[2] !== null ? sumRangeLike(a[2], range, ctx) : range;
    const xs: number[] = [];
    range.forEach((row, r) => row.forEach((v, c) => {
      const t = target[r]?.[c];
      if (test(v) && typeof t === 'number') xs.push(t);
    }));
    return xs.length ? Number((sum(xs) / xs.length).toPrecision(15)) : E.div0();
  } },
  AVERAGEIFS: { fn: (a, ctx) => {
    const target = toArr(a[0], ctx);
    const m = ifsMask(a.slice(1), ctx, target);
    if (isErr(m)) return m;
    const xs: number[] = [];
    target.forEach((row, r) => row.forEach((v, c) => m[r]?.[c] && typeof v === 'number' && xs.push(v)));
    return xs.length ? Number((sum(xs) / xs.length).toPrecision(15)) : E.div0();
  } },
  MAXIFS: { fn: (a, ctx) => {
    const target = toArr(a[0], ctx);
    const m = ifsMask(a.slice(1), ctx, target);
    if (isErr(m)) return m;
    const xs: number[] = [];
    target.forEach((row, r) => row.forEach((v, c) => m[r]?.[c] && typeof v === 'number' && xs.push(v)));
    return xs.length ? Math.max(...xs) : 0;
  } },
  MINIFS: { fn: (a, ctx) => {
    const target = toArr(a[0], ctx);
    const m = ifsMask(a.slice(1), ctx, target);
    if (isErr(m)) return m;
    const xs: number[] = [];
    target.forEach((row, r) => row.forEach((v, c) => m[r]?.[c] && typeof v === 'number' && xs.push(v)));
    return xs.length ? Math.min(...xs) : 0;
  } },
  SUMPRODUCT: { fn: (a, ctx) => {
    const ms = a.map((x) => toArr(x, ctx));
    const rows = ms[0]?.length ?? 0;
    const cols = ms[0]?.[0]?.length ?? 0;
    if (ms.some((m) => m.length !== rows || (m[0]?.length ?? 0) !== cols)) return E.value();
    let s = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let p = 1;
        for (const m of ms) {
          const v = m[r][c];
          if (isErr(v)) return v;
          p *= typeof v === 'number' ? v : 0;
        }
        s += p;
      }
    }
    return Number(s.toPrecision(15));
  } },
  ROUND: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    const d = num1(a[1] ?? 0, ctx);
    return err1(x, d) ?? roundTo(x as number, Math.trunc(d as number), 'round');
  } },
  ROUNDUP: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    const d = num1(a[1] ?? 0, ctx);
    return err1(x, d) ?? roundTo(x as number, Math.trunc(d as number), 'up');
  } },
  ROUNDDOWN: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    const d = num1(a[1] ?? 0, ctx);
    return err1(x, d) ?? roundTo(x as number, Math.trunc(d as number), 'down');
  } },
  INT: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    return isErr(x) ? x : Math.floor(x);
  } },
  ABS: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    return isErr(x) ? x : Math.abs(x);
  } },
  MOD: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    const d = num1(a[1], ctx);
    if (isErr(x)) return x;
    if (isErr(d)) return d;
    if (d === 0) return E.div0();
    return Number((x - d * Math.floor(x / d)).toPrecision(15));
  } },
  // ---- 逻辑 ----
  IF: { lazy: (args, ctx) => {
    const c = toBool(scalar(evaluate(args[0], ctx), ctx));
    if (isErr(c)) return c;
    const pick = c ? args[1] : args[2];
    if (!pick) return c ? true : false;
    if (pick.t === 'missing') return 0;
    return evaluate(pick, ctx);
  } },
  IFS: { lazy: (args, ctx) => {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const c = toBool(scalar(evaluate(args[i], ctx), ctx));
      if (isErr(c)) return c;
      if (c) return evaluate(args[i + 1], ctx);
    }
    return E.na();
  } },
  IFERROR: { lazy: (args, ctx) => {
    const v = evaluate(args[0], ctx);
    const s = isRange(v) ? scalar(v, ctx) : v;
    return isErr(s) ? (args[1] ? evaluate(args[1], ctx) : 0) : v;
  } },
  IFNA: { lazy: (args, ctx) => {
    const v = evaluate(args[0], ctx);
    const s = isRange(v) ? scalar(v, ctx) : v;
    return isErr(s) && s.code === '#N/A' ? (args[1] ? evaluate(args[1], ctx) : 0) : v;
  } },
  AND: { fn: (a, ctx) => {
    let seen = false;
    for (const { v, fromRef } of values(a, ctx)) {
      if (isErr(v)) return v;
      if (v === null || (fromRef && typeof v === 'string')) continue;
      const b = toBool(v);
      if (isErr(b)) return b;
      seen = true;
      if (!b) return false;
    }
    return seen ? true : E.value();
  } },
  OR: { fn: (a, ctx) => {
    let seen = false;
    let any = false;
    for (const { v, fromRef } of values(a, ctx)) {
      if (isErr(v)) return v;
      if (v === null || (fromRef && typeof v === 'string')) continue;
      const b = toBool(v);
      if (isErr(b)) return b;
      seen = true;
      any ||= b;
    }
    return seen ? any : E.value();
  } },
  NOT: { fn: (a, ctx) => {
    const b = toBool(scalar(a[0], ctx));
    return isErr(b) ? b : !b;
  } },
  ISBLANK: { fn: (a, ctx) => scalar(a[0], ctx) === null },
  ISNUMBER: { fn: (a, ctx) => typeof scalar(a[0], ctx) === 'number' },
  ISTEXT: { fn: (a, ctx) => typeof scalar(a[0], ctx) === 'string' },
  ISERROR: { fn: (a, ctx) => isErr(scalar(a[0], ctx)) },
  ISNA: { fn: (a, ctx) => {
    const v = scalar(a[0], ctx);
    return isErr(v) && v.code === '#N/A';
  } },
  // ---- 查找 ----
  VLOOKUP: { fn: (a, ctx) => lookup(a, ctx, true) },
  HLOOKUP: { fn: (a, ctx) => lookup(a, ctx, false) },
  INDEX: { fn: (a, ctx) => {
    const src = a[0];
    const r = a[1] === undefined || a[1] === null ? 0 : num1(a[1], ctx);
    const c = a[2] === undefined || a[2] === null ? 0 : num1(a[2], ctx);
    if (isErr(r)) return r;
    if (isErr(c)) return c;
    if (isRange(src)) {
      const rows = src.r1 - src.r0 + 1;
      const cols = src.c1 - src.c0 + 1;
      // 一行 / 一列的区域只给一个序号时，按它的方向取
      let ri = Math.trunc(r);
      let ci = Math.trunc(c);
      if (rows === 1 && ci === 0 && a[2] === undefined) [ri, ci] = [1, ri];
      if (ri < 0 || ci < 0 || ri > rows || ci > cols) return E.ref();
      if (ri === 0 && ci === 0) return src;
      if (ri === 0) return { ...src, c0: src.c0 + ci - 1, c1: src.c0 + ci - 1 };
      if (ci === 0) return cols === 1 ? ctx.get(src.si, src.r0 + ri - 1, src.c0) : { ...src, r0: src.r0 + ri - 1, r1: src.r0 + ri - 1 };
      return ctx.get(src.si, src.r0 + ri - 1, src.c0 + ci - 1);
    }
    const m = toArr(src, ctx);
    const ri = Math.max(1, Math.trunc(r)) - 1;
    const ci = Math.max(1, Math.trunc(c)) - 1;
    return m[ri]?.[ci] ?? E.ref();
  } },
  MATCH: { fn: (a, ctx) => {
    const needle = scalar(a[0], ctx);
    if (isErr(needle)) return needle;
    const list = flat(toArr(a[1], ctx));
    const type = a[2] === undefined || a[2] === null ? 1 : num1(a[2], ctx);
    if (isErr(type)) return type;
    if (type === 0) {
      const i = lookupIndex(needle, list, false);
      return i < 0 ? E.na() : i + 1;
    }
    if (type > 0) {
      const i = lookupIndex(needle, list, true);
      return i < 0 ? E.na() : i + 1;
    }
    // -1：降序列表里最后一个 ≥ needle 的
    let best = -1;
    for (let i = 0; i < list.length; i++) {
      if (list[i] === null) continue;
      if (compare(list[i], needle) >= 0) best = i;
      else break;
    }
    return best < 0 ? E.na() : best + 1;
  } },
  XLOOKUP: { fn: (a, ctx) => {
    const needle = scalar(a[0], ctx);
    if (isErr(needle)) return needle;
    const keys = toArr(a[1], ctx);
    const vertical = keys.length > 1 || (keys[0]?.length ?? 0) === 1;
    const list = flat(keys);
    const mode = a[4] === undefined || a[4] === null ? 0 : num1(a[4], ctx);
    if (isErr(mode)) return mode;
    let i = -1;
    if (mode === 2) i = lookupIndex(needle, list, false);
    else {
      i = list.findIndex((v) => v !== null && typeof v === typeof needle && compare(v, needle) === 0);
      if (i < 0 && (mode === -1 || mode === 1)) {
        // -1：精确或下一个较小的；1：精确或下一个较大的
        let best = -1;
        list.forEach((v, j) => {
          if (v === null || typeof v !== typeof needle) return;
          const d = compare(v, needle);
          if (mode === -1 && d < 0 && (best < 0 || compare(v, list[best]) > 0)) best = j;
          if (mode === 1 && d > 0 && (best < 0 || compare(v, list[best]) < 0)) best = j;
        });
        i = best;
      }
    }
    if (i < 0) return a[3] !== undefined && a[3] !== null ? scalar(a[3], ctx) : E.na();
    const ret = a[2];
    if (isRange(ret)) {
      return vertical
        ? ret.c0 === ret.c1 ? ctx.get(ret.si, ret.r0 + i, ret.c0) : { ...ret, r0: ret.r0 + i, r1: ret.r0 + i }
        : ret.r0 === ret.r1 ? ctx.get(ret.si, ret.r0, ret.c0 + i) : { ...ret, c0: ret.c0 + i, c1: ret.c0 + i };
    }
    const m = toArr(ret, ctx);
    return (vertical ? m[i]?.[0] : m[0]?.[i]) ?? E.na();
  } },
  ROW: { fn: (a, ctx) => (a[0] && isRange(a[0]) ? a[0].r0 + 1 : ctx.r + 1) },
  COLUMN: { fn: (a, ctx) => (a[0] && isRange(a[0]) ? a[0].c0 + 1 : ctx.c + 1) },
  CHOOSE: { lazy: (args, ctx) => {
    const i = toNum(scalar(evaluate(args[0], ctx), ctx), ctx);
    if (isErr(i)) return i;
    const k = Math.trunc(i);
    if (k < 1 || k >= args.length) return E.value();
    return evaluate(args[k], ctx);
  } },
  // ---- 文本 ----
  LEN: { fn: (a, ctx) => {
    const s = str1(a[0], ctx);
    return isErr(s) ? s : [...s].length;
  } },
  LEFT: { fn: (a, ctx) => {
    const s = str1(a[0], ctx);
    const n = a[1] === undefined ? 1 : num1(a[1], ctx);
    if (isErr(s)) return s;
    if (isErr(n)) return n;
    if (n < 0) return E.value();
    return [...s].slice(0, Math.trunc(n)).join('');
  } },
  RIGHT: { fn: (a, ctx) => {
    const s = str1(a[0], ctx);
    const n = a[1] === undefined ? 1 : num1(a[1], ctx);
    if (isErr(s)) return s;
    if (isErr(n)) return n;
    if (n < 0) return E.value();
    const ch = [...s];
    return Math.trunc(n) === 0 ? '' : ch.slice(-Math.trunc(n)).join('');
  } },
  MID: { fn: (a, ctx) => {
    const s = str1(a[0], ctx);
    const start = num1(a[1], ctx);
    const n = num1(a[2], ctx);
    const e = err1(s, start, n);
    if (e) return e;
    if ((start as number) < 1 || (n as number) < 0) return E.value();
    return [...(s as string)].slice(Math.trunc(start as number) - 1, Math.trunc(start as number) - 1 + Math.trunc(n as number)).join('');
  } },
  TRIM: { fn: (a, ctx) => {
    const s = str1(a[0], ctx);
    return isErr(s) ? s : s.replace(/ +/g, ' ').trim();
  } },
  UPPER: { fn: (a, ctx) => {
    const s = str1(a[0], ctx);
    return isErr(s) ? s : s.toUpperCase();
  } },
  LOWER: { fn: (a, ctx) => {
    const s = str1(a[0], ctx);
    return isErr(s) ? s : s.toLowerCase();
  } },
  CONCATENATE: { fn: (a, ctx) => {
    let out = '';
    for (const x of a) {
      const s = str1(x, ctx);
      if (isErr(s)) return s;
      out += s;
    }
    return out;
  } },
  CONCAT: { fn: (a, ctx) => {
    let out = '';
    for (const { v } of values(a, ctx)) {
      const s = toStr(v);
      if (isErr(s)) return s;
      out += s;
    }
    return out;
  } },
  TEXTJOIN: { fn: (a, ctx) => {
    const sep = str1(a[0], ctx);
    const skip = toBool(scalar(a[1], ctx));
    const e = err1(sep, skip);
    if (e) return e;
    const parts: string[] = [];
    for (const { v } of values(a.slice(2), ctx)) {
      const s = toStr(v);
      if (isErr(s)) return s;
      if (s === '' && skip) continue;
      parts.push(s);
    }
    return parts.join(sep as string);
  } },
  SUBSTITUTE: { fn: (a, ctx) => {
    const s = str1(a[0], ctx);
    const from = str1(a[1], ctx);
    const to = str1(a[2], ctx);
    const e = err1(s, from, to);
    if (e) return e;
    if (from === '') return s;
    if (a[3] === undefined) return (s as string).split(from as string).join(to as string);
    const k = num1(a[3], ctx);
    if (isErr(k)) return k;
    if (k < 1) return E.value();
    let idx = -1;
    for (let i = 0; i < k; i++) {
      idx = (s as string).indexOf(from as string, idx + 1);
      if (idx < 0) return s;
    }
    return (s as string).slice(0, idx) + to + (s as string).slice(idx + (from as string).length);
  } },
  FIND: { fn: (a, ctx) => {
    const needle = str1(a[0], ctx);
    const hay = str1(a[1], ctx);
    const start = a[2] === undefined ? 1 : num1(a[2], ctx);
    const e = err1(needle, hay, start);
    if (e) return e;
    const i = (hay as string).indexOf(needle as string, Math.trunc(start as number) - 1);
    return i < 0 || (start as number) < 1 ? E.value() : i + 1;
  } },
  SEARCH: { fn: (a, ctx) => {
    const needle = str1(a[0], ctx);
    const hay = str1(a[1], ctx);
    const start = a[2] === undefined ? 1 : num1(a[2], ctx);
    const e = err1(needle, hay, start);
    if (e) return e;
    const re = new RegExp(wildcard(needle as string).source.replace(/^\^|\$$/g, ''), 'i');
    const sub = (hay as string).slice(Math.trunc(start as number) - 1);
    const m = sub.match(re);
    return !m || m.index === undefined ? E.value() : m.index + Math.trunc(start as number);
  } },
  TEXT: { fn: (a, ctx) => {
    const v = scalar(a[0], ctx);
    const f = str1(a[1], ctx);
    if (isErr(v)) return v;
    if (isErr(f)) return f;
    const n = typeof v === 'number' ? v : typeof v === 'string' ? toNum(v, ctx) : v === null ? 0 : null;
    if (n === null || isErr(n)) return toStr(v);
    try {
      return formatNumber(n, f, ctx.book.date1904);
    } catch {
      return E.value();
    }
  } },
  VALUE: { fn: (a, ctx) => {
    const v = scalar(a[0], ctx);
    if (typeof v === 'number') return v;
    if (isErr(v)) return v;
    return textToNum(toStr(v) as string, ctx);
  } },
  // ---- 日期 ----
  TODAY: { fn: (_a, ctx) => {
    const d = new Date();
    return dateSerial(d.getFullYear(), d.getMonth() + 1, d.getDate(), ctx.book.date1904);
  } },
  NOW: { fn: (_a, ctx) => {
    const d = new Date();
    const s = dateSerial(d.getFullYear(), d.getMonth() + 1, d.getDate(), ctx.book.date1904);
    return s + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
  } },
  DATE: { fn: (a, ctx) => {
    const y = num1(a[0], ctx);
    const m = num1(a[1], ctx);
    const d = num1(a[2], ctx);
    const e = err1(y, m, d);
    if (e) return e;
    let yy = Math.trunc(y as number);
    if (yy < 1900) yy += 1900;
    // 月、日超出范围时进位（DATE(2026,13,1) = 2027-1-1）
    const t = new Date(Date.UTC(yy, Math.trunc(m as number) - 1, Math.trunc(d as number)));
    const s = dateSerial(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate(), ctx.book.date1904);
    return s < 0 ? E.num() : s;
  } },
  YEAR: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    return isErr(x) ? x : x < 0 ? E.num() : ymd(x, ctx.book.date1904).y;
  } },
  MONTH: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    return isErr(x) ? x : x < 0 ? E.num() : ymd(x, ctx.book.date1904).m;
  } },
  DAY: { fn: (a, ctx) => {
    const x = num1(a[0], ctx);
    return isErr(x) ? x : x < 0 ? E.num() : ymd(x, ctx.book.date1904).d;
  } },
  DATEDIF: { fn: (a, ctx) => {
    const s = num1(a[0], ctx);
    const t = num1(a[1], ctx);
    const u = str1(a[2], ctx);
    const e = err1(s, t, u);
    if (e) return e;
    if ((s as number) > (t as number)) return E.num();
    const A = ymd(s as number, ctx.book.date1904);
    const B = ymd(t as number, ctx.book.date1904);
    const months = (B.y - A.y) * 12 + (B.m - A.m) - (B.d < A.d ? 1 : 0);
    switch ((u as string).toUpperCase()) {
      case 'D':
        return Math.floor(t as number) - Math.floor(s as number);
      case 'M':
        return months;
      case 'Y':
        return Math.floor(months / 12);
      case 'YM':
        return months % 12;
      case 'MD': {
        if (B.d >= A.d) return B.d - A.d;
        const prev = new Date(Date.UTC(B.y, B.m - 1, 0)).getUTCDate();
        return prev - A.d + B.d;
      }
      case 'YD': {
        const ys = B.y - (B.m < A.m || (B.m === A.m && B.d < A.d) ? 1 : 0);
        const start = dateSerial(ys, A.m, A.d, ctx.book.date1904);
        return Math.floor(t as number) - start;
      }
      default:
        return E.num();
    }
  } },
  EDATE: { fn: (a, ctx) => {
    const s = num1(a[0], ctx);
    const m = num1(a[1], ctx);
    const e = err1(s, m);
    if (e) return e;
    return addMonths(s as number, Math.trunc(m as number), false, ctx);
  } },
  EOMONTH: { fn: (a, ctx) => {
    const s = num1(a[0], ctx);
    const m = num1(a[1], ctx);
    const e = err1(s, m);
    if (e) return e;
    return addMonths(s as number, Math.trunc(m as number), true, ctx);
  } },
  WEEKDAY: { fn: (a, ctx) => {
    const s = num1(a[0], ctx);
    const t = a[1] === undefined || a[1] === null ? 1 : num1(a[1], ctx);
    const e = err1(s, t);
    if (e) return e;
    if ((s as number) < 0) return E.num();
    const w = weekday(s as number, ctx); // 0 = 星期日
    switch (Math.trunc(t as number)) {
      case 1:
      case 17:
        return w + 1;
      case 2:
      case 11:
        return ((w + 6) % 7) + 1;
      case 3:
        return (w + 6) % 7;
      case 12:
      case 13:
      case 14:
      case 15:
      case 16: {
        // 12 = 星期二为 1 … 16 = 星期六为 1
        const start = Math.trunc(t as number) - 10;
        return ((w - start + 7) % 7) + 1;
      }
      default:
        return E.num();
    }
  } },
  NETWORKDAYS: { fn: (a, ctx) => {
    const s = num1(a[0], ctx);
    const t = num1(a[1], ctx);
    const e = err1(s, t);
    if (e) return e;
    const hol = holidays(a[2], ctx);
    if (isErr(hol)) return hol;
    let from = Math.floor(s as number);
    let to = Math.floor(t as number);
    const sign = from > to ? -1 : 1;
    if (sign < 0) [from, to] = [to, from];
    let n = 0;
    for (let d = from; d <= to; d++) if (isWorkday(d, hol, ctx)) n++;
    return n * sign;
  } },
  WORKDAY: { fn: (a, ctx) => {
    const s = num1(a[0], ctx);
    const k = num1(a[1], ctx);
    const e = err1(s, k);
    if (e) return e;
    const hol = holidays(a[2], ctx);
    if (isErr(hol)) return hol;
    let d = Math.floor(s as number);
    let left = Math.trunc(k as number);
    const step = left < 0 ? -1 : 1;
    while (left !== 0) {
      d += step;
      if (isWorkday(d, hol, ctx)) left -= step;
    }
    return d;
  } },
  RANK: { fn: (a, ctx) => rank(a, ctx) },
  LARGE: { fn: (a, ctx) => kth(a, ctx, true) },
  SMALL: { fn: (a, ctx) => kth(a, ctx, false) },
  MEDIAN: { fn: (a, ctx) => {
    const xs = numbers(a, ctx);
    if (isErr(xs)) return xs;
    if (!xs.length) return E.num();
    const s = [...xs].sort((x, y) => x - y);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  } },
  STDEV: { fn: (a, ctx) => spread(a, ctx, true, true) },
  STDEVP: { fn: (a, ctx) => spread(a, ctx, true, false) },
  VAR: { fn: (a, ctx) => spread(a, ctx, false, true) },
  VARP: { fn: (a, ctx) => spread(a, ctx, false, false) },
  SUBTOTAL: { fn: (a, ctx) => {
    const k = num1(a[0], ctx);
    if (isErr(k)) return k;
    const code = Math.trunc(k);
    const name = SUBTOTAL_FUNCS[code % 100];
    if (!name || (code > 11 && code < 101) || code > 111) return E.value();
    // 区域里的 SUBTOTAL 不算（避免小计重复计入）；101–111 忽略所有隐藏行，1–11 只忽略被筛选掉的行
    const parts: V[] = [];
    for (const x of a.slice(1)) {
      if (!isRange(x)) return E.value();
      const s = ctx.book.sheets[x.si];
      const { r1, c1 } = rangeDims(x, ctx);
      const af = s?.autoFilter?.ref;
      const data: Scalar[][] = [];
      for (let r = x.r0; r <= r1; r++) {
        const hidden = !!s?.rowDefs?.[r]?.hidden;
        if (hidden && (code > 100 || (af && r > af.r0 && r <= af.r1))) continue;
        const row: Scalar[] = [];
        for (let c = x.c0; c <= c1; c++) row.push(/SUBTOTAL\s*\(/i.test(s?.cells[r]?.[c]?.f ?? '') ? null : ctx.get(x.si, r, c));
        data.push(row);
      }
      parts.push({ kind: 'arr', data });
    }
    return (FUNCS[name] as { fn: (a: V[], ctx: Ctx) => V }).fn(parts, ctx);
  } },
  OFFSET: { fn: (a, ctx) => {
    const base = a[0];
    if (!isRange(base)) return E.value();
    const dr = num1(a[1], ctx);
    const dc = num1(a[2], ctx);
    const h = a[3] === undefined || a[3] === null ? base.r1 - base.r0 + 1 : num1(a[3], ctx);
    const w = a[4] === undefined || a[4] === null ? base.c1 - base.c0 + 1 : num1(a[4], ctx);
    const e = err1(dr, dc, h, w);
    if (e) return e;
    const r0 = base.r0 + Math.trunc(dr as number);
    const c0 = base.c0 + Math.trunc(dc as number);
    const hh = Math.trunc(h as number);
    const ww = Math.trunc(w as number);
    if (hh === 0 || ww === 0) return E.ref();
    const rr = [r0, r0 + hh + (hh > 0 ? -1 : 1)].sort((x, y) => x - y);
    const cc = [c0, c0 + ww + (ww > 0 ? -1 : 1)].sort((x, y) => x - y);
    if (rr[0] < 0 || cc[0] < 0 || rr[1] > 1048575 || cc[1] > 16383) return E.ref();
    return { kind: 'range', si: base.si, r0: rr[0], r1: rr[1], c0: cc[0], c1: cc[1] };
  } },
  INDIRECT: { fn: (a, ctx) => {
    const t = str1(a[0], ctx);
    if (isErr(t)) return t;
    if (a[1] !== undefined && a[1] !== null) {
      const a1 = toBool(scalar(a[1], ctx));
      if (isErr(a1)) return a1;
      if (!a1) return E.ref(); // R1C1 样式暂不支持
    }
    try {
      const node = parseFormula(t.trim());
      if (node.t !== 'ref' && node.t !== 'name') return E.ref();
      const v = evaluate(node, ctx);
      return isRange(v) ? v : E.ref();
    } catch {
      return E.ref();
    }
  } },
};

const SUBTOTAL_FUNCS: Record<number, string> = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };

/** 星期几：0 = 星期日。1900 日期系统里序列号 1（1900-1-1）是星期日，Excel 虚构的 1900-2-29（60）之后的星期也都按连续序号算 */
function weekday(serial: number, ctx: Ctx): number {
  const s = Math.floor(serial) + (ctx.book.date1904 ? 1462 : 0);
  return (((s - 1) % 7) + 7) % 7;
}

function isWorkday(d: number, hol: Set<number>, ctx: Ctx): boolean {
  const w = weekday(d, ctx);
  return w !== 0 && w !== 6 && !hol.has(d);
}

function holidays(v: V | undefined, ctx: Ctx): Set<number> | FErr {
  const out = new Set<number>();
  if (v === undefined || v === null) return out;
  for (const row of toArr(v, ctx))
    for (const x of row) {
      if (x === null) continue;
      const n = toNum(x, ctx);
      if (isErr(n)) return n;
      out.add(Math.floor(n));
    }
  return out;
}

/** 加几个月；那个月没有这一天时取最后一天（EDATE(1-31, 1) = 2-28）；end = 取结果那个月的最后一天 */
function addMonths(serial: number, months: number, end: boolean, ctx: Ctx): number | FErr {
  if (serial < 0) return E.num();
  const { y, m, d } = ymd(serial, ctx.book.date1904);
  const total = y * 12 + (m - 1) + months;
  const ny = Math.floor(total / 12);
  const nm = (total % 12) + 1;
  const last = new Date(Date.UTC(ny, nm, 0)).getUTCDate();
  const out = dateSerial(ny, nm, end ? last : Math.min(d, last), ctx.book.date1904);
  return out < 0 ? E.num() : out;
}

function rank(a: V[], ctx: Ctx): Scalar {
  const x = num1(a[0], ctx);
  if (isErr(x)) return x;
  const asc = a[2] === undefined || a[2] === null ? 0 : num1(a[2], ctx);
  if (isErr(asc)) return asc;
  const xs: number[] = [];
  for (const row of toArr(a[1], ctx)) for (const v of row) if (typeof v === 'number') xs.push(v);
  if (!xs.includes(x)) return E.na();
  return 1 + xs.filter((v) => (asc ? v < x : v > x)).length;
}

function kth(a: V[], ctx: Ctx, largest: boolean): Scalar {
  const xs = numbers([a[0]], ctx);
  const k = num1(a[1], ctx);
  if (isErr(xs)) return xs;
  if (isErr(k)) return k;
  const i = Math.ceil(k);
  if (i < 1 || i > xs.length) return E.num();
  const s = [...xs].sort((p, q) => (largest ? q - p : p - q));
  return s[i - 1];
}

/** 标准差 / 方差；sample = 样本（除以 n-1） */
function spread(a: V[], ctx: Ctx, sqrt: boolean, sample: boolean): Scalar {
  const xs = numbers(a, ctx);
  if (isErr(xs)) return xs;
  const n = xs.length;
  if (n < (sample ? 2 : 1)) return E.div0();
  const mean = xs.reduce((p, q) => p + q, 0) / n;
  const v = xs.reduce((p, q) => p + (q - mean) ** 2, 0) / (sample ? n - 1 : n);
  return Number((sqrt ? Math.sqrt(v) : v).toPrecision(15));
}

/** SUMIF 的求和区域：与条件区域同样大小、从求和区域左上角开始 */
function sumRangeLike(v: V, shape: Scalar[][], ctx: Ctx): Scalar[][] {
  if (!isRange(v)) return toArr(v, ctx);
  const rows = shape.length;
  const cols = shape[0]?.length ?? 0;
  return toArr({ ...v, r1: v.r0 + rows - 1, c1: v.c0 + cols - 1 }, ctx);
}

function lookup(a: V[], ctx: Ctx, vertical: boolean): V {
  const needle = scalar(a[0], ctx);
  if (isErr(needle)) return needle;
  const table = toArr(a[1], ctx);
  const k = num1(a[2], ctx);
  if (isErr(k)) return k;
  const approx = a[3] === undefined || a[3] === null ? true : toBool(scalar(a[3], ctx));
  if (isErr(approx)) return approx;
  const idx = Math.trunc(k);
  const width = vertical ? (table[0]?.length ?? 0) : table.length;
  if (idx < 1) return E.value();
  if (idx > width) return E.ref();
  const keys = vertical ? table.map((row) => row[0]) : (table[0] ?? []);
  const i = lookupIndex(needle, keys, approx);
  if (i < 0) return E.na();
  const v = vertical ? table[i][idx - 1] : table[idx - 1][i];
  return v ?? 0;
}

/** 日期文字 → 序列号（只认带 - / 年月日 分隔的日期，纯数字不算） */
function parseDate(text: string, ctx: Ctx): number | null {
  if (!/[-/年]/.test(text) || !/\d/.test(text)) return null;
  const p = parseInput(text.trim(), 'General', ctx.book.date1904);
  return p.kind === 'value' && typeof p.v === 'number' ? p.v : null;
}

Object.assign(
  FUNCS,
  moreFunctions({
    E, mkErr: (code) => new FErr(code), isErr, isRange, isArr, toArr, scalar, toNum, toStr, toBool, compare, numbers, values, num1, str1, err1,
    wildcard, ymd, dateSerial, weekday, roundTo, parseDate,
  }),
);
// Excel 2010 起带点号的同义名
for (const [alias, name] of [['STDEV.S', 'STDEV'], ['STDEV.P', 'STDEVP'], ['VAR.S', 'VAR'], ['VAR.P', 'VARP'], ['RANK.EQ', 'RANK']] as const) FUNCS[alias] = FUNCS[name];

export const SUPPORTED = new Set(Object.keys(FUNCS));

// ---------- 重算 ----------

interface Dep {
  si: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

interface FInfo {
  si: number;
  r: number;
  c: number;
  key: string;
  ast: Node | null;
  deps: Dep[];
  /** 含 OFFSET / INDIRECT：引用哪里要算了才知道，每次重算都算（同 Excel 的易失函数） */
  volatile?: boolean;
}

const VOLATILE = /\b(OFFSET|INDIRECT|RAND|RANDBETWEEN)\s*\(/i;

const keyOf = (si: number, r: number, c: number) => `${si}:${r}:${c}`;

/** 公式能否由编辑器计算：函数都认识、没有数组公式 */
function supported(node: Node, names: (n: string) => Node | null, seen = new Set<string>()): boolean {
  switch (node.t) {
    case 'fn':
      return SUPPORTED.has(node.name) && node.args.every((a) => supported(a, names, seen));
    case 'un':
      return supported(node.a, names, seen);
    case 'bin':
      return supported(node.a, names, seen) && supported(node.b, names, seen);
    case 'arr':
      return node.rows.every((row) => row.every((x) => supported(x, names, seen)));
    case 'name': {
      const key = node.name.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      const n = names(node.name);
      return !!n && supported(n, names, seen);
    }
    default:
      return true;
  }
}

interface Rect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** "D1:D5" / "D1" → 区域 */
function parseRect(ref: string): Rect | null {
  const m = ref.match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i);
  if (!m) return null;
  const c0 = colIndex(m[1]);
  const r0 = Number(m[2]) - 1;
  return { r0, c0, r1: m[4] ? Number(m[4]) - 1 : r0, c1: m[3] ? colIndex(m[3]) : c0 };
}

export class Engine {
  private index: FInfo[] | null = null;
  private asts = new Map<string, Node | null>();
  private nameAsts = new Map<string, Node | null>();
  /** 不能计算的公式（键 si:r:c） */
  readonly unsupported = new Set<string>();
  /** 上次重算遇到的循环引用格子数 */
  circular = 0;
  /** 还没有结果的公式（文件里没存缓存值、或新输入的），下次重算时一起算 */
  private pending: FInfo[] = [];

  constructor(readonly book: Workbook) {}

  /** 工作表结构或名称变了：下次重算前重建公式索引 */
  invalidate(): void {
    this.index = null;
    this.nameAsts.clear();
  }

  private parse(f: string): Node | null {
    if (!this.asts.has(f)) {
      try {
        this.asts.set(f, parseFormula(f));
      } catch {
        this.asts.set(f, null);
      }
    }
    return this.asts.get(f)!;
  }

  private nameNode = (name: string, si = -1): Node | null => {
    const names = this.book.names ?? [];
    const low = name.toLowerCase();
    const hit = names.find((n) => n.name.toLowerCase() === low && n.local === si) ?? names.find((n) => n.name.toLowerCase() === low && n.local === undefined);
    if (!hit) return null;
    const key = `${hit.local ?? ''}|${low}`;
    if (!this.nameAsts.has(key)) {
      try {
        this.nameAsts.set(key, parseFormula(hit.f));
      } catch {
        this.nameAsts.set(key, null);
      }
    }
    return this.nameAsts.get(key)!;
  };

  sheetIndex(name: string): number {
    const low = name.toLowerCase();
    return this.book.sheets.findIndex((s) => s.name.toLowerCase() === low);
  }

  private depsOf(node: Node, si: number, out: Dep[], seen = new Set<string>()): void {
    switch (node.t) {
      case 'ref': {
        const s = node.ref.sheet === undefined ? si : this.sheetIndex(node.ref.sheet);
        if (s < 0) return;
        const a = node.ref.a;
        const b = node.ref.b ?? a;
        out.push({ si: s, r0: Math.min(a.r ?? 0, b.r ?? 0), c0: Math.min(a.c ?? 0, b.c ?? 0), r1: Math.max(a.r ?? 1048575, b.r ?? 1048575), c1: Math.max(a.c ?? 16383, b.c ?? 16383) });
        return;
      }
      case 'fn':
        node.args.forEach((x) => this.depsOf(x, si, out, seen));
        return;
      case 'un':
        this.depsOf(node.a, si, out, seen);
        return;
      case 'bin':
        this.depsOf(node.a, si, out, seen);
        this.depsOf(node.b, si, out, seen);
        return;
      case 'arr':
        node.rows.forEach((row) => row.forEach((x) => this.depsOf(x, si, out, seen)));
        return;
      case 'name': {
        const key = node.name.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        const n = this.nameNode(node.name, si);
        if (n) this.depsOf(n, si, out, seen);
        return;
      }
    }
  }

  /** 打开文件后：动态数组区域里非公式的格子是溢出结果，标上 sp（重算时才知道哪些可以覆盖） */
  private markLoadedSpills(): void {
    for (const s of this.book.sheets)
      s.cells.forEach((row, r) =>
        row?.forEach((cell, c) => {
          const g = cell?.f && cell.fa && cell.dyn ? parseRect(cell.fa) : null;
          if (!g) return;
          for (let rr = g.r0; rr <= g.r1; rr++)
            for (let cc = g.c0; cc <= g.c1; cc++) {
              const x = s.cells[rr]?.[cc];
              if (x && !x.f && !x.fraw && (rr !== r || cc !== c)) x.sp = true;
            }
        }),
      );
  }

  /** 清掉不属于任何动态数组公式的溢出格子（公式被删除 / 撤销 / 改成了别的） */
  private purgeOrphanSpills(): void {
    for (const s of this.book.sheets) {
      const rects: Rect[] = [];
      s.cells.forEach((row) =>
        row?.forEach((cell) => {
          const g = cell?.f && cell.fa && cell.dyn ? parseRect(cell.fa) : null;
          if (g) rects.push(g);
        }),
      );
      s.cells.forEach((row, r) =>
        row?.forEach((cell, c) => {
          if (!cell?.sp) return;
          if (rects.some((g) => r >= g.r0 && r <= g.r1 && c >= g.c0 && c <= g.c1)) return;
          row[c] = cell.s ? ({ v: '', t: '', s: cell.s } as SCell) : undefined;
          s.dirty = true;
        }),
      );
    }
  }

  private build(): FInfo[] {
    this.purgeOrphanSpills();
    const list: FInfo[] = [];
    this.pending = [];
    this.unsupported.clear();
    this.book.sheets.forEach((s, si) => {
      s.cells.forEach((row, r) => {
        row?.forEach((cell, c) => {
          if (!cell?.f && !cell?.fraw) return;
          const key = keyOf(si, r, c);
          const ast = cell.f && (!cell.fa || cell.dyn) ? this.parse(cell.f) : null;
          const ok = !!ast && supported(ast, (n) => this.nameNode(n, si));
          const deps: Dep[] = [];
          if (ast) this.depsOf(ast, si, deps);
          if (!ok) this.unsupported.add(key);
          const info: FInfo = { si, r, c, key, ast: ok ? ast : null, deps, ...(ok && VOLATILE.test(cell.f!) && { volatile: true }) };
          list.push(info);
          if (ok && cell.v === '' && !cell.k) this.pending.push(info);
        });
      });
    });
    return list;
  }

  /** 分析全部公式（打开文件后调用）：标出不能计算的公式，算出没有缓存值的公式。返回算了几个 */
  analyze(): number {
    this.markLoadedSpills();
    this.index = this.build();
    return this.pending.length ? this.recalc([]) : 0;
  }

  isUnsupported(si: number, r: number, c: number): boolean {
    return this.unsupported.has(keyOf(si, r, c));
  }

  private ctx(si: number, r: number, c: number): Ctx {
    const book = this.book;
    return {
      book,
      si,
      r,
      c,
      get: (s, rr, cc) => cellValue(book.sheets[s]?.cells[rr]?.[cc]),
      name: (n) => this.nameNode(n, si),
      sheetIndex: (n) => this.sheetIndex(n),
    };
  }

  /** 单独计算一个公式（不写入格子） */
  evalFormula(f: string, si: number, r: number, c: number): Scalar {
    const ast = this.parse(f);
    if (!ast) return E.name();
    return finalValue(evaluate(ast, this.ctx(si, r, c)), this.ctx(si, r, c));
  }

  /**
   * 改了 changed 这些格子之后重算。rebuild = 公式本身变了（输入 / 删除公式），先重建索引；
   * all = 结构改动（插删行列、改表名）后全部重算。返回重算的公式个数。
   */
  recalc(changed: { si: number; r: number; c: number }[], opts: { rebuild?: boolean; all?: boolean; spillPass?: boolean } = {}): number {
    const rebuild = opts.rebuild || opts.all;
    if (rebuild || !this.index) {
      this.index = this.build();
      this.nameAsts.clear();
    }
    const index = this.index;
    if (!index.length) return 0;
    const byKey = new Map(index.map((f) => [f.key, f]));
    const dirty = new Map<string, FInfo>();
    let wave = new Map<string, { si: number; r: number; c: number }>();
    for (const p of changed) {
      const k = keyOf(p.si, p.r, p.c);
      wave.set(k, p);
      const f = byKey.get(k);
      if (f?.ast) dirty.set(k, f);
    }
    for (const f of this.pending) {
      dirty.set(f.key, f);
      wave.set(f.key, f);
    }
    this.pending = [];
    // 动态数组：溢出区域里的格子被用户改了、或上次是 #SPILL!（挡路的格子可能已经清掉）→ 重算
    for (const f of index) {
      const cell = f.ast ? this.cellOf(f) : undefined;
      if (!cell?.dyn) continue;
      const g = cell.fa ? parseRect(cell.fa) : null;
      const blocked = cell.k === 'e' && cell.v === '#SPILL!';
      const hit = !opts.spillPass && g && changed.some((p) => p.si === f.si && p.r >= g.r0 && p.r <= g.r1 && p.c >= g.c0 && p.c <= g.c1 && (p.r !== f.r || p.c !== f.c));
      if (blocked || hit) {
        dirty.set(f.key, f);
        wave.set(f.key, f);
      }
    }
    // 溢出的格子也算“被改了”：只引用其中某一格的公式要跟着重算
    const addSpill = (f: FInfo, into: Map<string, { si: number; r: number; c: number }>) => {
      const cell = this.cellOf(f);
      const g = cell?.dyn && cell.fa ? parseRect(cell.fa) : null;
      if (g) for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) into.set(keyOf(f.si, r, c), { si: f.si, r, c });
    };
    for (const f of [...wave.keys()].map((k) => byKey.get(k)).filter((f): f is FInfo => !!f)) addSpill(f, wave);
    // 易失公式：有任何改动就重算（它们引用的格子由公式算出来，依赖表里没有）
    let hasVolatile = false;
    if (wave.size || opts.all)
      for (const f of index)
        if (f.volatile) {
          hasVolatile = true;
          dirty.set(f.key, f);
          wave.set(f.key, f);
        }
    if (opts.all) {
      // 结构改动后，公式文字变了的格子都要重算：全部重算最稳妥
      for (const f of index) if (f.ast) dirty.set(f.key, f);
      wave = new Map();
    }
    const hits = (d: Dep, cells: Map<string, { si: number; r: number; c: number }>) => {
      const area = (d.r1 - d.r0 + 1) * (d.c1 - d.c0 + 1);
      if (area <= cells.size) {
        for (let r = d.r0; r <= d.r1; r++) for (let c = d.c0; c <= d.c1; c++) if (cells.has(keyOf(d.si, r, c))) return true;
        return false;
      }
      for (const p of cells.values()) if (p.si === d.si && p.r >= d.r0 && p.r <= d.r1 && p.c >= d.c0 && p.c <= d.c1) return true;
      return false;
    };
    while (wave.size) {
      const next = new Map<string, { si: number; r: number; c: number }>();
      for (const f of index) {
        if (!f.ast || dirty.has(f.key)) continue;
        if (f.deps.some((d) => hits(d, wave))) {
          dirty.set(f.key, f);
          next.set(f.key, f);
          addSpill(f, next);
        }
      }
      wave = next;
    }
    if (!dirty.size) return 0;

    // 拓扑排序：先算被引用的
    const list = [...dirty.values()];
    const asCells = new Map(list.map((f) => [f.key, f]));
    const indeg = new Map<string, number>();
    const users = new Map<string, string[]>();
    for (const f of list) {
      let n = 0;
      for (const d of f.deps) {
        const area = (d.r1 - d.r0 + 1) * (d.c1 - d.c0 + 1);
        const pre: FInfo[] = [];
        if (area <= asCells.size) {
          for (let r = d.r0; r <= d.r1; r++) for (let c = d.c0; c <= d.c1; c++) {
            const g = asCells.get(keyOf(d.si, r, c));
            if (g) pre.push(g);
          }
        } else for (const g of list) if (g.si === d.si && g.r >= d.r0 && g.r <= d.r1 && g.c >= d.c0 && g.c <= d.c1) pre.push(g);
        for (const g of list) {
          if (pre.includes(g) || g === f) continue;
          const gc = this.cellOf(g);
          const rect = gc?.dyn && gc.fa ? parseRect(gc.fa) : null;
          if (rect && g.si === d.si && rect.r0 <= d.r1 && rect.r1 >= d.r0 && rect.c0 <= d.c1 && rect.c1 >= d.c0) pre.push(g);
        }
        for (const g of pre) {
          n++;
          (users.get(g.key) ?? users.set(g.key, []).get(g.key)!).push(f.key);
        }
      }
      indeg.set(f.key, n);
    }
    const queue = list.filter((f) => indeg.get(f.key) === 0);
    const order: FInfo[] = [];
    while (queue.length) {
      const f = queue.shift()!;
      order.push(f);
      for (const u of users.get(f.key) ?? []) {
        const n = indeg.get(u)! - 1;
        indeg.set(u, n);
        if (n === 0) queue.push(asCells.get(u)!);
      }
    }
    const done = new Set(order.map((f) => f.key));
    this.circular = 0;
    this.grown = [];
    for (const f of order) this.compute(f);
    // 易失公式引用的格子可能排在它后面才算好：再算一遍（OFFSET / INDIRECT 套 OFFSET 这类多层的不保证）
    if (hasVolatile) for (const f of order) this.compute(f);
    for (const f of list) {
      if (done.has(f.key)) continue;
      this.circular++;
      this.clearSpill(f);
      this.write(f, new FErr('#CIRC!'));
    }
    // 溢出区域变大：新占用的格子里如果有别的公式引用，补算一遍
    const grown = this.grown;
    if (grown.length && !opts.spillPass) this.recalc(grown, { spillPass: true });
    return list.length;
  }

  /** 上次重算里新占用的溢出格子 */
  private grown: { si: number; r: number; c: number }[] = [];

  private cellOf(f: FInfo): SCell | undefined {
    return this.book.sheets[f.si]?.cells[f.r]?.[f.c];
  }

  /** 算一个公式；动态数组公式的结果是多格数组时溢出到相邻格子 */
  private compute(f: FInfo): void {
    const ctx = this.ctx(f.si, f.r, f.c);
    const v = evaluate(f.ast!, ctx);
    const cell = this.cellOf(f);
    if (cell?.dyn && isArr(v) && (v.data.length > 1 || (v.data[0]?.length ?? 0) > 1)) return this.spill(f, v.data);
    this.clearSpill(f);
    this.write(f, finalValue(v, ctx));
  }

  /** 删掉公式上次溢出出来的格子（连同 fa） */
  private clearSpill(f: FInfo): void {
    const sheet = this.book.sheets[f.si];
    const cell = this.cellOf(f);
    if (!cell?.dyn || !cell.fa) return;
    const g = parseRect(cell.fa);
    if (g)
      for (let r = g.r0; r <= g.r1; r++)
        for (let c = g.c0; c <= g.c1; c++) {
          const x = sheet.cells[r]?.[c];
          if (!x?.sp || (r === f.r && c === f.c)) continue;
          sheet.cells[r][c] = x.s ? ({ v: '', t: '', s: x.s } as SCell) : undefined;
        }
    delete cell.fa;
    sheet.dirty = true;
  }

  private spill(f: FInfo, data: Scalar[][]): void {
    const sheet = this.book.sheets[f.si];
    const anchor = this.cellOf(f)!;
    const old = anchor.fa ? parseRect(anchor.fa) : null;
    this.clearSpill(f);
    const rows = data.length;
    const cols = data[0]?.length ?? 0;
    let blocked = f.r + rows > 1048576 || f.c + cols > 16384;
    for (let r = f.r; r < f.r + rows && !blocked; r++)
      for (let c = f.c; c < f.c + cols && !blocked; c++) {
        if (r === f.r && c === f.c) continue;
        const x = sheet.cells[r]?.[c];
        if (x && (x.v !== '' || x.f || x.fraw)) blocked = true;
      }
    if (blocked) return this.write(f, new FErr('#SPILL!'));
    const put = (r: number, c: number, v: Scalar) => {
      const next: Pick<SCell, 'v' | 'k'> = isErr(v) ? { v: v.code, k: 'e' } : { v: v === null ? 0 : v };
      const row = (sheet.cells[r] ??= []);
      const cell: SCell = { v: next.v, t: '', sp: true };
      if (next.k) cell.k = next.k;
      const s = row[c]?.s;
      if (s) cell.s = s;
      refreshText(this.book, cell);
      row[c] = cell;
      if (!old || r < old.r0 || r > old.r1 || c < old.c0 || c > old.c1) this.grown.push({ si: f.si, r, c });
    };
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) if (r || c) put(f.r + r, f.c + c, data[r][c]);
    anchor.fa = `${colName(f.c)}${f.r + 1}:${colName(f.c + cols - 1)}${f.r + rows}`;
    this.write(f, data[0][0]);
    sheet.dirty = true;
  }

  private write(f: FInfo, v: Scalar) {
    const sheet = this.book.sheets[f.si];
    const cell = sheet.cells[f.r]?.[f.c];
    if (!cell) return;
    const next: Pick<SCell, 'v' | 'k'> = isErr(v) ? { v: v.code, k: 'e' } : { v: v === null ? 0 : v };
    if (cell.v === next.v && cell.k === next.k) return;
    cell.v = next.v;
    if (next.k) cell.k = next.k;
    else delete cell.k;
    refreshText(this.book, cell);
    sheet.dirty = true;
  }
}

/** 格子里的值 → 求值用的值（公式格子用它的结果） */
function cellValue(cell: SCell | undefined): Scalar {
  if (!cell) return null;
  if (cell.k === 'e') return new FErr(String(cell.v));
  if (cell.v === '' && (cell.f || cell.fraw)) return '';
  return cell.v;
}

/** 公式的最终结果：区域按隐式交集取一个值，数组取左上角 */
function finalValue(v: V, ctx: Ctx): Scalar {
  if (isRange(v)) return scalar(v, ctx);
  if (isArr(v)) return v.data[0]?.[0] ?? null;
  return v;
}

/** 公式的格式提示：TODAY / DATE / NOW 开头的公式在“常规”格子里自动用日期格式（同 Excel） */
export function formulaFormat(f: string): string | undefined {
  if (/^\s*(TODAY|DATE)\s*\(/i.test(f)) return 'yyyy/m/d';
  if (/^\s*NOW\s*\(/i.test(f)) return 'yyyy/m/d h:mm';
  return undefined;
}


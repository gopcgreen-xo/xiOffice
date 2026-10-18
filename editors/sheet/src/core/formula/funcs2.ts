/**
 * 第二批函数：数学 / 统计 / 文本 / 信息 / 查找与动态数组 / 日期 / 财务。
 * 由 eval.ts 在加载时通过 `moreFunctions(helpers)` 注册（辅助函数注入，避免两个文件互相 import）。
 * 返回数组的函数（FILTER / UNIQUE / SORT / SEQUENCE / TRANSPOSE / TEXTSPLIT …）的结果由 Engine 溢出到相邻格子。
 */
import type { ArrV, Ctx, FErr, RangeV, Scalar, V } from './eval';

export interface Helpers {
  E: Record<'div0' | 'value' | 'ref' | 'name' | 'num' | 'na', () => FErr>;
  mkErr: (code: string) => FErr;
  isErr: (v: unknown) => v is FErr;
  isRange: (v: V) => v is RangeV;
  isArr: (v: V) => v is ArrV;
  toArr: (v: V, ctx: Ctx) => Scalar[][];
  scalar: (v: V, ctx: Ctx) => Scalar;
  toNum: (v: Scalar, ctx: Ctx) => number | FErr;
  toStr: (v: Scalar) => string | FErr;
  toBool: (v: Scalar) => boolean | FErr;
  compare: (a: Scalar, b: Scalar) => number;
  numbers: (args: V[], ctx: Ctx) => number[] | FErr;
  values: (args: V[], ctx: Ctx) => Generator<{ v: Scalar; fromRef: boolean }>;
  num1: (v: V, ctx: Ctx) => number | FErr;
  str1: (v: V, ctx: Ctx) => string | FErr;
  err1: (...xs: unknown[]) => FErr | null;
  wildcard: (p: string) => RegExp;
  ymd: (serial: number, date1904?: boolean) => { y: number; m: number; d: number };
  dateSerial: (y: number, m: number, d: number, date1904?: boolean) => number;
  weekday: (serial: number, ctx: Ctx) => number;
  roundTo: (x: number, d: number, mode: 'round' | 'up' | 'down') => number;
  parseDate: (text: string, ctx: Ctx) => number | null;
}

type Fn = { fn: (args: V[], ctx: Ctx) => V };

const MAX_CELLS = 1_000_000;

export function moreFunctions(h: Helpers): Record<string, Fn> {
  const { E, isErr, isRange, isArr, toArr, scalar, toNum, toStr, toBool, compare, numbers, values, num1, str1, err1, wildcard, ymd, dateSerial, weekday, roundTo } = h;
  const arr = (data: Scalar[][]): ArrV => ({ kind: 'arr', data });
  const fin = (x: number): number | FErr => (Number.isFinite(x) ? Number(x.toPrecision(15)) : E.num());
  /** 可选参数：没写 / 空 → 默认值 */
  const opt = (v: V | undefined, d: number, ctx: Ctx): number | FErr => (v === undefined || v === null ? d : num1(v, ctx));
  const optBool = (v: V | undefined, d: boolean, ctx: Ctx): boolean | FErr => (v === undefined || v === null ? d : toBool(scalar(v, ctx)));
  const dims = (v: V) => (isRange(v) ? { r: v.r1 - v.r0 + 1, c: v.c1 - v.c0 + 1 } : isArr(v) ? { r: v.data.length, c: v.data[0]?.length ?? 0 } : { r: 1, c: 1 });
  const unary = (f: (x: number) => number | FErr): Fn => ({
    fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      return isErr(x) ? x : f(x);
    },
  });
  const binary = (f: (x: number, y: number) => number | FErr): Fn => ({
    fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      const y = num1(a[1], ctx);
      return err1(x, y) ?? f(x as number, y as number);
    },
  });
  const pairs = (a: V, b: V, ctx: Ctx): [number[], number[]] | FErr => {
    const A = toArr(a, ctx).flat();
    const B = toArr(b, ctx).flat();
    if (A.length !== B.length) return E.na();
    const xs: number[] = [];
    const ys: number[] = [];
    for (let i = 0; i < A.length; i++) {
      if (isErr(A[i])) return A[i] as FErr;
      if (isErr(B[i])) return B[i] as FErr;
      if (typeof A[i] === 'number' && typeof B[i] === 'number') {
        xs.push(A[i] as number);
        ys.push(B[i] as number);
      }
    }
    return [xs, ys];
  };
  const mean = (xs: number[]) => xs.reduce((p, q) => p + q, 0) / xs.length;
  /** 线性回归（最小二乘） */
  const regress = (a: V[], ctx: Ctx): { slope: number; intercept: number; r: number } | FErr => {
    // SLOPE / INTERCEPT / FORECAST / CORREL 参数是 (known_y, known_x)
    const p = pairs(a[0], a[1], ctx);
    if (isErr(p)) return p;
    const [ys, xs] = p;
    if (xs.length < 2) return E.div0();
    const mx = mean(xs);
    const my = mean(ys);
    let sxy = 0;
    let sxx = 0;
    let syy = 0;
    for (let i = 0; i < xs.length; i++) {
      sxy += (xs[i] - mx) * (ys[i] - my);
      sxx += (xs[i] - mx) ** 2;
      syy += (ys[i] - my) ** 2;
    }
    if (sxx === 0 || syy === 0) return E.div0();
    const slope = sxy / sxx;
    return { slope, intercept: my - slope * mx, r: sxy / Math.sqrt(sxx * syy) };
  };
  const percentile = (xs: number[], k: number): number | FErr => {
    if (!xs.length || k < 0 || k > 1) return E.num();
    const s = [...xs].sort((p, q) => p - q);
    const pos = (s.length - 1) * k;
    const lo = Math.floor(pos);
    const hi = Math.ceil(pos);
    return fin(s[lo] + (s[hi] - s[lo]) * (pos - lo));
  };
  const modeFn: Fn = {
    fn: (a, ctx) => {
      const xs = numbers(a, ctx);
      if (isErr(xs)) return xs;
      const count = new Map<number, number>();
      for (const x of xs) count.set(x, (count.get(x) ?? 0) + 1);
      let best: number | null = null;
      let n = 1;
      for (const [x, c] of count) if (c > n) [best, n] = [x, c];
      return best ?? E.na();
    },
  };
  const percentileFn: Fn = {
    fn: (a, ctx) => {
      const xs = numbers([a[0]], ctx);
      const k = num1(a[1], ctx);
      return isErr(xs) ? xs : isErr(k) ? k : percentile(xs, k);
    },
  };
  const quartileFn: Fn = {
    fn: (a, ctx) => {
      const xs = numbers([a[0]], ctx);
      const q = num1(a[1], ctx);
      if (isErr(xs)) return xs;
      if (isErr(q)) return q;
      const n = Math.trunc(q);
      return n < 0 || n > 4 ? E.num() : percentile(xs, n / 4);
    },
  };
  const gcd = (x: number, y: number): number => (y === 0 ? x : gcd(y, x % y));

  // ---------- 财务（Excel 的符号约定：付出为负、收入为正） ----------
  const pmt = (r: number, n: number, pv: number, fv: number, t: number) =>
    r === 0 ? -(pv + fv) / n : -(r * (pv * (1 + r) ** n + fv)) / ((1 + r * t) * ((1 + r) ** n - 1));
  const fv = (r: number, n: number, p: number, pv: number, t: number) => (r === 0 ? -(pv + p * n) : -(pv * (1 + r) ** n + (p * (1 + r * t) * ((1 + r) ** n - 1)) / r));
  const fin5 = (a: V[], ctx: Ctx, from: number): number[] | FErr => {
    // 从 from 开始的可选参数：fv / pv，type
    const out: number[] = [];
    for (let i = 0; i < 5; i++) {
      const x = i < from ? num1(a[i], ctx) : opt(a[i], 0, ctx);
      if (isErr(x)) return x;
      out.push(x);
    }
    return out;
  };
  const npv = (r: number, flows: number[]) => flows.reduce((s, x, i) => s + x / (1 + r) ** (i + 1), 0);
  /** 牛顿法求根，失败返回 null */
  const newton = (f: (x: number) => number, x0: number): number | null => {
    let x = x0;
    for (let i = 0; i < 100; i++) {
      const y = f(x);
      const d = (f(x + 1e-7) - y) / 1e-7;
      if (!Number.isFinite(y) || !Number.isFinite(d) || d === 0) return null;
      const nx = x - y / d;
      if (!Number.isFinite(nx)) return null;
      if (Math.abs(nx - x) < 1e-10) return nx;
      x = nx;
    }
    return null;
  };

  // ---------- 日期 ----------
  const serialOf = (v: V, ctx: Ctx): number | FErr => {
    const s = scalar(v, ctx);
    if (typeof s === 'string') {
      const d = h.parseDate(s, ctx);
      if (d !== null) return d;
    }
    return toNum(s, ctx);
  };
  const isoWeek = (serial: number, ctx: Ctx): number => {
    const d = Math.floor(serial);
    const wd = (weekday(d, ctx) + 6) % 7; // 0 = 星期一
    const thursday = d - wd + 3;
    const { y } = ymd(thursday, ctx.book.date1904);
    const jan1 = dateSerial(y, 1, 1, ctx.book.date1904);
    return Math.floor((thursday - jan1) / 7) + 1;
  };
  const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const timePart = (v: V, ctx: Ctx): number | FErr => {
    const x = serialOf(v, ctx);
    if (isErr(x)) return x;
    if (x < 0) return E.num();
    return Math.round((x - Math.floor(x)) * 86400) % 86400;
  };

  // ---------- 动态数组 ----------
  const truthy = (v: Scalar): boolean | FErr => (typeof v === 'string' ? E.value() : toBool(v));
  const keyOf = (row: Scalar[]) => row.map((v) => (typeof v === 'string' ? `s:${v.toLowerCase()}` : isErr(v) ? `e:${v.code}` : `${typeof v}:${v}`)).join('\u0001');
  const transposeM = (m: Scalar[][]): Scalar[][] => (m[0] ?? []).map((_, c) => m.map((row) => row[c]));
  /** 排序比较：空永远排最后 */
  const cmpBy = (asc: boolean) => (x: Scalar, y: Scalar) => {
    if (x === null || x === '') return y === null || y === '' ? 0 : 1;
    if (y === null || y === '') return -1;
    const d = compare(x, y);
    return asc ? d : -d;
  };
  const sortRows = (rows: Scalar[][], keys: { by: Scalar[]; asc: boolean }[]): Scalar[][] => {
    const idx = rows.map((_, i) => i);
    idx.sort((i, j) => {
      for (const k of keys) {
        const d = cmpBy(k.asc)(k.by[i], k.by[j]);
        if (d) return d;
      }
      return i - j;
    });
    return idx.map((i) => rows[i]);
  };
  const orderOf = (v: V | undefined, ctx: Ctx): boolean | FErr => {
    const o = opt(v, 1, ctx);
    return isErr(o) ? o : o === 1 ? true : o === -1 ? false : E.value();
  };
  const delims = (v: V | undefined, ctx: Ctx): string[] | FErr => {
    if (v === undefined || v === null) return [];
    const out: string[] = [];
    for (const x of toArr(v, ctx).flat()) {
      const s = toStr(x);
      if (isErr(s)) return s;
      if (s !== '') out.push(s);
    }
    return out;
  };
  const splitBy = (s: string, ds: string[], skipEmpty: boolean, ci: boolean): string[] => {
    if (!ds.length) return [s];
    const parts: string[] = [];
    let cur = '';
    const hay = ci ? s.toLowerCase() : s;
    const needles = ds.map((d) => (ci ? d.toLowerCase() : d));
    for (let i = 0; i < s.length; ) {
      const hit = needles.findIndex((d) => hay.startsWith(d, i));
      if (hit >= 0) {
        parts.push(cur);
        cur = '';
        i += needles[hit].length;
      } else cur += s[i++];
    }
    parts.push(cur);
    return skipEmpty ? parts.filter((p) => p !== '') : parts;
  };
  const textAround = (a: V[], ctx: Ctx, before: boolean): Scalar => {
    const s = str1(a[0], ctx);
    const d = str1(a[1], ctx);
    const n = opt(a[2], 1, ctx);
    const e = err1(s, d, n);
    if (e) return e;
    const text = s as string;
    const delim = d as string;
    const k = Math.trunc(n as number);
    if (k === 0 || Math.abs(k) > text.length + 1) return E.value();
    let at = -1;
    if (delim === '') at = k > 0 ? 0 : text.length;
    else if (k > 0) {
      let from = 0;
      for (let i = 0; i < k; i++) {
        at = text.indexOf(delim, from);
        if (at < 0) return E.na();
        from = at + delim.length;
      }
    } else {
      let from = text.length;
      for (let i = 0; i < -k; i++) {
        at = from < 0 ? -1 : text.lastIndexOf(delim, from);
        if (at < 0) return E.na();
        from = at - delim.length;
      }
    }
    return before ? text.slice(0, at) : text.slice(at + delim.length);
  };

  const F: Record<string, Fn> = {
    // ---------- 数学 ----------
    SQRT: unary((x) => (x < 0 ? E.num() : fin(Math.sqrt(x)))),
    POWER: binary((x, y) => (x === 0 && y === 0 ? E.num() : fin(Math.pow(x, y)))),
    SIGN: unary((x) => Math.sign(x) + 0),
    EXP: unary((x) => fin(Math.exp(x))),
    LN: unary((x) => (x <= 0 ? E.num() : fin(Math.log(x)))),
    LOG10: unary((x) => (x <= 0 ? E.num() : fin(Math.log10(x)))),
    LOG: { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      const b = opt(a[1], 10, ctx);
      const e = err1(x, b);
      if (e) return e;
      if ((x as number) <= 0 || (b as number) <= 0 || b === 1) return (b as number) === 1 ? E.div0() : E.num();
      return fin(Math.log(x as number) / Math.log(b as number));
    } },
    PI: { fn: () => Math.PI },
    SIN: unary((x) => fin(Math.sin(x))),
    COS: unary((x) => fin(Math.cos(x))),
    TAN: unary((x) => fin(Math.tan(x))),
    ASIN: unary((x) => (Math.abs(x) > 1 ? E.num() : fin(Math.asin(x)))),
    ACOS: unary((x) => (Math.abs(x) > 1 ? E.num() : fin(Math.acos(x)))),
    ATAN: unary((x) => fin(Math.atan(x))),
    ATAN2: binary((x, y) => (x === 0 && y === 0 ? E.div0() : fin(Math.atan2(y, x)))),
    DEGREES: unary((x) => fin((x * 180) / Math.PI)),
    RADIANS: unary((x) => fin((x * Math.PI) / 180)),
    RAND: { fn: () => Math.random() },
    RANDBETWEEN: binary((lo, hi) => (lo > hi ? E.num() : Math.floor(Math.ceil(lo) + Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)))),
    FACT: unary((x) => (x < 0 || x > 170 ? E.num() : fin([...Array(Math.trunc(x))].reduce((p, _, i) => p * (i + 1), 1)))),
    TRUNC: { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      const d = opt(a[1], 0, ctx);
      return err1(x, d) ?? roundTo(x as number, Math.trunc(d as number), 'down');
    } },
    QUOTIENT: binary((x, y) => (y === 0 ? E.div0() : Math.trunc(x / y))),
    CEILING: binary((x, s) => (s === 0 ? 0 : x > 0 && s < 0 ? E.num() : fin(Math.ceil(Number((x / s).toPrecision(15))) * s))),
    FLOOR: binary((x, s) => (s === 0 ? E.div0() : x > 0 && s < 0 ? E.num() : fin(Math.floor(Number((x / s).toPrecision(15))) * s))),
    'CEILING.MATH': { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      const s = opt(a[1], (x as number) >= 0 ? 1 : -1, ctx);
      const m = opt(a[2], 0, ctx);
      const e = err1(x, s, m);
      if (e) return e;
      const sig = Math.abs(s as number);
      if (sig === 0) return 0;
      const q = Number(((x as number) / sig).toPrecision(15));
      return fin((q < 0 && m ? Math.floor(q) : Math.ceil(q)) * sig);
    } },
    'FLOOR.MATH': { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      const s = opt(a[1], 1, ctx);
      const m = opt(a[2], 0, ctx);
      const e = err1(x, s, m);
      if (e) return e;
      const sig = Math.abs(s as number);
      if (sig === 0) return 0;
      const q = Number(((x as number) / sig).toPrecision(15));
      return fin((q < 0 && m ? Math.ceil(q) : Math.floor(q)) * sig);
    } },
    MROUND: binary((x, m) => (m === 0 ? 0 : x * m < 0 ? E.num() : fin(Math.round(Number((x / m).toPrecision(15))) * m))),
    GCD: { fn: (a, ctx) => {
      const xs = numbers(a, ctx);
      if (isErr(xs)) return xs;
      if (xs.some((x) => x < 0)) return E.num();
      return xs.reduce((p, q) => gcd(p, Math.trunc(q)), 0);
    } },
    LCM: { fn: (a, ctx) => {
      const xs = numbers(a, ctx);
      if (isErr(xs)) return xs;
      if (xs.some((x) => x < 0)) return E.num();
      return xs.reduce((p, q) => (p === 0 || q === 0 ? 0 : (p / gcd(p, Math.trunc(q))) * Math.trunc(q)), 1);
    } },
    SUMSQ: { fn: (a, ctx) => {
      const xs = numbers(a, ctx);
      return isErr(xs) ? xs : fin(xs.reduce((p, q) => p + q * q, 0));
    } },

    // ---------- 统计 ----------
    MODE: modeFn,
    'MODE.SNGL': modeFn,
    PERCENTILE: percentileFn,
    'PERCENTILE.INC': percentileFn,
    QUARTILE: quartileFn,
    'QUARTILE.INC': quartileFn,
    CORREL: { fn: (a, ctx) => {
      const r = regress(a, ctx);
      return isErr(r) ? r : fin(r.r);
    } },
    RSQ: { fn: (a, ctx) => {
      const r = regress(a, ctx);
      return isErr(r) ? r : fin(r.r * r.r);
    } },
    SLOPE: { fn: (a, ctx) => {
      const r = regress(a, ctx);
      return isErr(r) ? r : fin(r.slope);
    } },
    INTERCEPT: { fn: (a, ctx) => {
      const r = regress(a, ctx);
      return isErr(r) ? r : fin(r.intercept);
    } },
    FORECAST: { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      const r = regress(a.slice(1), ctx);
      if (isErr(x)) return x;
      return isErr(r) ? r : fin(r.intercept + r.slope * x);
    } },

    // ---------- 文本 ----------
    REPT: { fn: (a, ctx) => {
      const s = str1(a[0], ctx);
      const n = num1(a[1], ctx);
      const e = err1(s, n);
      if (e) return e;
      const k = Math.trunc(n as number);
      return k < 0 || (s as string).length * k > 32767 ? E.value() : (s as string).repeat(k);
    } },
    CHAR: { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      return isErr(x) ? x : x < 1 || x >= 256 ? E.value() : String.fromCharCode(Math.trunc(x));
    } },
    UNICHAR: { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      return isErr(x) ? x : x < 1 || x > 0x10ffff ? E.value() : String.fromCodePoint(Math.trunc(x));
    } },
    CODE: { fn: (a, ctx) => {
      const s = str1(a[0], ctx);
      return isErr(s) ? s : s === '' ? E.value() : s.charCodeAt(0);
    } },
    UNICODE: { fn: (a, ctx) => {
      const s = str1(a[0], ctx);
      return isErr(s) ? s : s === '' ? E.value() : s.codePointAt(0)!;
    } },
    PROPER: { fn: (a, ctx) => {
      const s = str1(a[0], ctx);
      return isErr(s) ? s : s.replace(/[A-Za-zÀ-ɏ]+/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase());
    } },
    CLEAN: { fn: (a, ctx) => {
      const s = str1(a[0], ctx);
      // eslint-disable-next-line no-control-regex
      return isErr(s) ? s : s.replace(/[\u0000-\u001f]/g, '');
    } },
    EXACT: { fn: (a, ctx) => {
      const x = str1(a[0], ctx);
      const y = str1(a[1], ctx);
      return err1(x, y) ?? x === y;
    } },
    REPLACE: { fn: (a, ctx) => {
      const s = str1(a[0], ctx);
      const st = num1(a[1], ctx);
      const n = num1(a[2], ctx);
      const t = str1(a[3], ctx);
      const e = err1(s, st, n, t);
      if (e) return e;
      if ((st as number) < 1 || (n as number) < 0) return E.value();
      const i = Math.trunc(st as number) - 1;
      return (s as string).slice(0, i) + (t as string) + (s as string).slice(i + Math.trunc(n as number));
    } },
    T: { fn: (a, ctx) => {
      const v = scalar(a[0], ctx);
      return typeof v === 'string' || isErr(v) ? v : '';
    } },
    TEXTBEFORE: { fn: (a, ctx) => textAround(a, ctx, true) },
    TEXTAFTER: { fn: (a, ctx) => textAround(a, ctx, false) },
    TEXTSPLIT: { fn: (a, ctx) => {
      const s = str1(a[0], ctx);
      const cd = delims(a[1], ctx);
      const rd = delims(a[2], ctx);
      const skip = optBool(a[3], false, ctx);
      const mode = opt(a[4], 0, ctx);
      const e = err1(s, cd, rd, skip, mode);
      if (e) return e;
      if (!(cd as string[]).length && !(rd as string[]).length) return E.value();
      const pad = a[5] === undefined ? E.na() : scalar(a[5], ctx);
      const rows = splitBy(s as string, rd as string[], skip as boolean, mode === 1).map((r) => splitBy(r, cd as string[], skip as boolean, mode === 1));
      const w = Math.max(...rows.map((r) => r.length));
      return arr(rows.map((r) => [...r, ...Array.from({ length: w - r.length }, () => pad)]));
    } },

    // ---------- 信息 / 逻辑 ----------
    ISERR: { fn: (a, ctx) => {
      const v = scalar(a[0], ctx);
      return isErr(v) && v.code !== '#N/A';
    } },
    ISLOGICAL: { fn: (a, ctx) => typeof scalar(a[0], ctx) === 'boolean' },
    ISNONTEXT: { fn: (a, ctx) => typeof scalar(a[0], ctx) !== 'string' },
    ISEVEN: { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      return isErr(x) ? x : Math.trunc(Math.abs(x)) % 2 === 0;
    } },
    ISODD: { fn: (a, ctx) => {
      const x = num1(a[0], ctx);
      return isErr(x) ? x : Math.trunc(Math.abs(x)) % 2 === 1;
    } },
    N: { fn: (a, ctx) => {
      const v = scalar(a[0], ctx);
      return isErr(v) ? v : typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0;
    } },
    TYPE: { fn: (a, ctx) => {
      if (isArr(a[0]) || (isRange(a[0]) && dims(a[0]).r * dims(a[0]).c > 1)) return 64;
      const v = scalar(a[0], ctx);
      return isErr(v) ? 16 : typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 1;
    } },
    NA: { fn: () => E.na() },
    TRUE: { fn: () => true },
    FALSE: { fn: () => false },
    XOR: { fn: (a, ctx) => {
      let n = 0;
      for (const { v, fromRef } of values(a, ctx)) {
        if (isErr(v)) return v;
        if (v === null || (fromRef && typeof v === 'string')) continue;
        const b = toBool(v);
        if (isErr(b)) return b;
        if (b) n++;
      }
      return n % 2 === 1;
    } },
    SWITCH: { fn: (a, ctx) => {
      const x = scalar(a[0], ctx);
      if (isErr(x)) return x;
      const rest = a.slice(1);
      for (let i = 0; i + 1 < rest.length; i += 2) {
        const c = scalar(rest[i], ctx);
        if (!isErr(c) && compare(x, c) === 0 && typeof x === typeof c) return rest[i + 1];
      }
      return rest.length % 2 === 1 ? rest[rest.length - 1] : E.na();
    } },

    // ---------- 查找 / 引用 ----------
    ROWS: { fn: (a) => dims(a[0]).r },
    COLUMNS: { fn: (a) => dims(a[0]).c },
    ADDRESS: { fn: (a, ctx) => {
      const r = num1(a[0], ctx);
      const c = num1(a[1], ctx);
      const abs = opt(a[2], 1, ctx);
      const a1 = optBool(a[3], true, ctx);
      const e = err1(r, c, abs, a1);
      if (e) return e;
      const R = Math.trunc(r as number);
      const C = Math.trunc(c as number);
      if (R < 1 || C < 1 || R > 1048576 || C > 16384) return E.value();
      const kind = Math.trunc(abs as number);
      if (kind < 1 || kind > 4) return E.value();
      let col = '';
      for (let n = C; n > 0; n = Math.floor((n - 1) / 26)) col = String.fromCharCode(65 + ((n - 1) % 26)) + col;
      let text: string;
      if (a1) text = `${kind === 1 || kind === 3 ? '$' : ''}${col}${kind === 1 || kind === 2 ? '$' : ''}${R}`;
      else text = `R${kind === 1 || kind === 2 ? R : `[${R}]`}C${kind === 1 || kind === 3 ? C : `[${C}]`}`;
      if (a[4] !== undefined && a[4] !== null) {
        const sh = str1(a[4], ctx);
        if (isErr(sh)) return sh;
        return `${/^[A-Za-z0-9_]+$/.test(sh) ? sh : `'${sh.replace(/'/g, "''")}'`}!${text}`;
      }
      return text;
    } },
    LOOKUP: { fn: (a, ctx) => {
      const needle = scalar(a[0], ctx);
      if (isErr(needle)) return needle;
      const look = toArr(a[1], ctx);
      const wide = (look[0]?.length ?? 0) > look.length;
      const keys = wide ? (look[0] ?? []) : look.map((row) => row[0]);
      let best = -1;
      for (let i = 0; i < keys.length; i++) {
        const v = keys[i];
        if (v === null || typeof v !== typeof needle) continue;
        if (compare(v, needle) <= 0) best = i;
        else break;
      }
      if (best < 0) return E.na();
      if (a[2] === undefined) return (wide ? look[look.length - 1][best] : look[best][look[best].length - 1]) ?? 0;
      const res = toArr(a[2], ctx).flat();
      return res[best] ?? E.na();
    } },
    XMATCH: { fn: (a, ctx) => {
      const needle = scalar(a[0], ctx);
      const mode = opt(a[2], 0, ctx);
      const dir = opt(a[3], 1, ctx);
      const e = err1(needle, mode, dir);
      if (e) return e;
      const list = toArr(a[1], ctx).flat();
      const re = mode === 2 && typeof needle === 'string' ? wildcard(needle) : null;
      const order = list.map((_, i) => i);
      if ((dir as number) < 0) order.reverse();
      let best = -1;
      for (const i of order) {
        const v = list[i];
        if (v === null) continue;
        if (re ? typeof v === 'string' && re.test(v) : typeof v === typeof needle && compare(v, needle) === 0) return i + 1;
        if (mode === 0 || mode === 2 || typeof v !== typeof needle) continue;
        const d = compare(v, needle);
        if (mode === -1 && d < 0 && (best < 0 || compare(v, list[best]) > 0)) best = i;
        if (mode === 1 && d > 0 && (best < 0 || compare(v, list[best]) < 0)) best = i;
      }
      return best < 0 ? E.na() : best + 1;
    } },
    TRANSPOSE: { fn: (a, ctx) => arr(transposeM(toArr(a[0], ctx))) },
    SEQUENCE: { fn: (a, ctx) => {
      const r = opt(a[0], 1, ctx);
      const c = opt(a[1], 1, ctx);
      const s = opt(a[2], 1, ctx);
      const st = opt(a[3], 1, ctx);
      const e = err1(r, c, s, st);
      if (e) return e;
      const R = Math.trunc(r as number);
      const C = Math.trunc(c as number);
      if (R < 1 || C < 1) return E.value();
      if (R * C > MAX_CELLS) return E.num();
      return arr(Array.from({ length: R }, (_, i) => Array.from({ length: C }, (_, j) => fin((s as number) + (i * C + j) * (st as number)) as number)));
    } },
    FILTER: { fn: (a, ctx) => {
      const data = toArr(a[0], ctx);
      const inc = toArr(a[1], ctx);
      const byRow = inc.length === data.length && (inc[0]?.length ?? 0) === 1;
      const byCol = !byRow && inc.length === 1 && (inc[0]?.length ?? 0) === (data[0]?.length ?? 0);
      if (!byRow && !byCol) return E.value();
      const flags = byCol ? inc[0] : inc.map((r) => r[0]);
      const keep: boolean[] = [];
      for (const f of flags) {
        if (isErr(f)) return f;
        const b = truthy(f);
        if (isErr(b)) return b;
        keep.push(b);
      }
      const out = byCol ? data.map((row) => row.filter((_, i) => keep[i])) : data.filter((_, i) => keep[i]);
      if (!out.length || !out[0].length) return a[2] === undefined ? h.mkErr('#CALC!') : (a[2] ?? null) as V;
      return arr(out);
    } },
    UNIQUE: { fn: (a, ctx) => {
      const byCol = optBool(a[1], false, ctx);
      const once = optBool(a[2], false, ctx);
      const e = err1(byCol, once);
      if (e) return e;
      const m = toArr(a[0], ctx);
      const rows = byCol ? transposeM(m) : m;
      const seen = new Map<string, number>();
      rows.forEach((row) => seen.set(keyOf(row), (seen.get(keyOf(row)) ?? 0) + 1));
      const done = new Set<string>();
      const out: Scalar[][] = [];
      for (const row of rows) {
        const k = keyOf(row);
        if (done.has(k) || (once && seen.get(k)! > 1)) continue;
        done.add(k);
        out.push(row);
      }
      if (!out.length) return h.mkErr('#CALC!');
      return arr(byCol ? transposeM(out) : out);
    } },
    SORT: { fn: (a, ctx) => {
      const idx = opt(a[1], 1, ctx);
      const asc = orderOf(a[2], ctx);
      const byCol = optBool(a[3], false, ctx);
      const e = err1(idx, asc, byCol);
      if (e) return e;
      const m = toArr(a[0], ctx);
      const rows = byCol ? transposeM(m) : m;
      const k = Math.trunc(idx as number) - 1;
      if (k < 0 || k >= (rows[0]?.length ?? 0)) return E.value();
      const sorted = sortRows(rows, [{ by: rows.map((r) => r[k]), asc: asc as boolean }]);
      return arr(byCol ? transposeM(sorted) : sorted);
    } },
    SORTBY: { fn: (a, ctx) => {
      const m = toArr(a[0], ctx);
      const byCol = m.length === 1 && (m[0]?.length ?? 0) > 1 && a.slice(1).every((x, i) => i % 2 || dims(x).r === 1);
      const rows = byCol ? transposeM(m) : m;
      const keys: { by: Scalar[]; asc: boolean }[] = [];
      for (let i = 1; i < a.length; i += 2) {
        const by = toArr(a[i], ctx).flat();
        if (by.length !== rows.length) return E.value();
        const asc = orderOf(a[i + 1], ctx);
        if (isErr(asc)) return asc;
        keys.push({ by, asc });
      }
      if (!keys.length) return E.value();
      const sorted = sortRows(rows, keys);
      return arr(byCol ? transposeM(sorted) : sorted);
    } },
    VSTACK: { fn: (a, ctx) => {
      const ms = a.map((x) => toArr(x, ctx));
      const w = Math.max(...ms.map((m) => m[0]?.length ?? 0));
      return arr(ms.flatMap((m) => m.map((row) => [...row, ...Array.from({ length: w - row.length }, () => E.na())])));
    } },
    HSTACK: { fn: (a, ctx) => {
      const ms = a.map((x) => toArr(x, ctx));
      const h2 = Math.max(...ms.map((m) => m.length));
      return arr(Array.from({ length: h2 }, (_, r) => ms.flatMap((m) => m[r] ?? Array.from({ length: m[0]?.length ?? 0 }, () => E.na()))));
    } },

    // ---------- 日期时间 ----------
    DAYS: { fn: (a, ctx) => {
      const e = serialOf(a[0], ctx);
      const s = serialOf(a[1], ctx);
      return err1(e, s) ?? Math.floor(e as number) - Math.floor(s as number);
    } },
    DATEVALUE: { fn: (a, ctx) => {
      const s = str1(a[0], ctx);
      if (isErr(s)) return s;
      const d = h.parseDate(s, ctx);
      return d === null ? E.value() : Math.floor(d);
    } },
    WEEKNUM: { fn: (a, ctx) => {
      const s = serialOf(a[0], ctx);
      const t = opt(a[1], 1, ctx);
      const e = err1(s, t);
      if (e) return e;
      if ((s as number) < 0) return E.num();
      const type = Math.trunc(t as number);
      if (type === 21) return isoWeek(s as number, ctx);
      // 一周从哪天开始：1 / 17 = 星期日，2 / 11 = 星期一，12–16 = 星期二…星期六
      const start = type === 1 || type === 17 ? 0 : type === 2 || type === 11 ? 1 : type >= 12 && type <= 16 ? type - 10 : -1;
      if (start < 0) return E.num();
      const d = Math.floor(s as number);
      const { y } = ymd(d, ctx.book.date1904);
      const jan1 = dateSerial(y, 1, 1, ctx.book.date1904);
      const back = (weekday(jan1, ctx) - start + 7) % 7;
      return Math.floor((d - jan1 + back) / 7) + 1;
    } },
    ISOWEEKNUM: { fn: (a, ctx) => {
      const s = serialOf(a[0], ctx);
      return isErr(s) ? s : s < 0 ? E.num() : isoWeek(s, ctx);
    } },
    YEARFRAC: { fn: (a, ctx) => {
      const s0 = serialOf(a[0], ctx);
      const e0 = serialOf(a[1], ctx);
      const b = opt(a[2], 0, ctx);
      const er = err1(s0, e0, b);
      if (er) return er;
      let s = Math.floor(s0 as number);
      let e = Math.floor(e0 as number);
      if (s > e) [s, e] = [e, s];
      const A = ymd(s, ctx.book.date1904);
      const B = ymd(e, ctx.book.date1904);
      const days = e - s;
      switch (Math.trunc(b as number)) {
        case 0: {
          // 美国 30/360
          let d1 = A.d;
          let d2 = B.d;
          const lastFeb = (y: number, m: number, d: number) => m === 2 && d === (isLeap(y) ? 29 : 28);
          if (lastFeb(A.y, A.m, A.d) && lastFeb(B.y, B.m, B.d)) d2 = 30;
          if (lastFeb(A.y, A.m, A.d)) d1 = 30;
          if (d2 === 31 && d1 >= 30) d2 = 30;
          if (d1 === 31) d1 = 30;
          return fin(((B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1)) / 360);
        }
        case 1: {
          if (A.y === B.y) return fin(days / (isLeap(A.y) ? 366 : 365));
          let total = 0;
          for (let y = A.y; y <= B.y; y++) total += isLeap(y) ? 366 : 365;
          return fin(days / (total / (B.y - A.y + 1)));
        }
        case 2:
          return fin(days / 360);
        case 3:
          return fin(days / 365);
        case 4:
          return fin(((B.y - A.y) * 360 + (B.m - A.m) * 30 + (Math.min(B.d, 30) - Math.min(A.d, 30))) / 360);
        default:
          return E.num();
      }
    } },
    HOUR: { fn: (a, ctx) => {
      const t = timePart(a[0], ctx);
      return isErr(t) ? t : Math.floor(t / 3600);
    } },
    MINUTE: { fn: (a, ctx) => {
      const t = timePart(a[0], ctx);
      return isErr(t) ? t : Math.floor((t % 3600) / 60);
    } },
    SECOND: { fn: (a, ctx) => {
      const t = timePart(a[0], ctx);
      return isErr(t) ? t : t % 60;
    } },
    TIME: { fn: (a, ctx) => {
      const hh = num1(a[0], ctx);
      const mm = num1(a[1], ctx);
      const ss = num1(a[2], ctx);
      const e = err1(hh, mm, ss);
      if (e) return e;
      const total = Math.trunc(hh as number) * 3600 + Math.trunc(mm as number) * 60 + Math.trunc(ss as number);
      if (total < 0) return E.num();
      return fin((total % 86400) / 86400);
    } },

    // ---------- 财务 ----------
    PMT: { fn: (a, ctx) => {
      const x = fin5(a, ctx, 3);
      if (isErr(x)) return x;
      const [r, n, pv, f, t] = x;
      return n === 0 ? E.num() : fin(pmt(r, n, pv, f, t));
    } },
    FV: { fn: (a, ctx) => {
      const x = fin5(a, ctx, 3);
      if (isErr(x)) return x;
      const [r, n, p, pv, t] = x;
      return fin(fv(r, n, p, pv, t));
    } },
    PV: { fn: (a, ctx) => {
      const x = fin5(a, ctx, 3);
      if (isErr(x)) return x;
      const [r, n, p, f, t] = x;
      return r === 0 ? fin(-(f + p * n)) : fin(-(f + (p * (1 + r * t) * ((1 + r) ** n - 1)) / r) / (1 + r) ** n);
    } },
    NPER: { fn: (a, ctx) => {
      const x = fin5(a, ctx, 3);
      if (isErr(x)) return x;
      const [r, p, pv, f, t] = x;
      if (r === 0) return p === 0 ? E.num() : fin(-(pv + f) / p);
      const num = p * (1 + r * t) - f * r;
      const den = p * (1 + r * t) + pv * r;
      return num / den <= 0 || r <= -1 ? E.num() : fin(Math.log(num / den) / Math.log(1 + r));
    } },
    IPMT: { fn: (a, ctx) => {
      const x = num1(a[1], ctx);
      const rest = fin5([a[0], a[2], a[3], a[4], a[5]], ctx, 3);
      if (isErr(x)) return x;
      if (isErr(rest)) return rest;
      const [r, n, pv, f, t] = rest;
      const per = Math.trunc(x);
      if (per < 1 || per > n) return E.num();
      const p = pmt(r, n, pv, f, t);
      if (t === 1) return per === 1 ? 0 : fin((fv(r, per - 2, p, pv, 1) - p) * r);
      return fin(fv(r, per - 1, p, pv, 0) * r);
    } },
    PPMT: { fn: (a, ctx) => {
      const x = num1(a[1], ctx);
      const rest = fin5([a[0], a[2], a[3], a[4], a[5]], ctx, 3);
      if (isErr(x)) return x;
      if (isErr(rest)) return rest;
      const [r, n, pv, f, t] = rest;
      const per = Math.trunc(x);
      if (per < 1 || per > n) return E.num();
      const p = pmt(r, n, pv, f, t);
      const ip = t === 1 ? (per === 1 ? 0 : (fv(r, per - 2, p, pv, 1) - p) * r) : fv(r, per - 1, p, pv, 0) * r;
      return fin(p - ip);
    } },
    NPV: { fn: (a, ctx) => {
      const r = num1(a[0], ctx);
      const xs = numbers(a.slice(1), ctx);
      if (isErr(r)) return r;
      if (isErr(xs)) return xs;
      return r === -1 ? E.div0() : fin(npv(r, xs));
    } },
    IRR: { fn: (a, ctx) => {
      const xs = numbers([a[0]], ctx);
      const g = opt(a[1], 0.1, ctx);
      if (isErr(xs)) return xs;
      if (isErr(g)) return g;
      if (!xs.some((x) => x > 0) || !xs.some((x) => x < 0)) return E.num();
      const f = (r: number) => xs.reduce((s, x, i) => s + x / (1 + r) ** i, 0);
      for (const start of [g, 0.1, -0.5, 0.5, 1]) {
        const r = newton(f, start);
        if (r !== null && r > -1) return fin(r);
      }
      return E.num();
    } },
    RATE: { fn: (a, ctx) => {
      const x = fin5(a, ctx, 3);
      const g = opt(a[5], 0.1, ctx);
      if (isErr(x)) return x;
      if (isErr(g)) return g;
      const [n, p, pv, f, t] = x;
      const fn = (r: number) => (r === 0 ? pv + p * n + f : pv * (1 + r) ** n + (p * (1 + r * t) * ((1 + r) ** n - 1)) / r + f);
      for (const start of [g, 0.1, 0.01, 0.5]) {
        const r = newton(fn, start || 0.0001);
        if (r !== null && r > -1) return fin(r);
      }
      return E.num();
    } },
    SLN: { fn: (a, ctx) => {
      const c = num1(a[0], ctx);
      const s = num1(a[1], ctx);
      const l = num1(a[2], ctx);
      const e = err1(c, s, l);
      return e ?? (l === 0 ? E.div0() : fin(((c as number) - (s as number)) / (l as number)));
    } },
  };
  return F;
}

/** 返回数组、会向相邻格子溢出的函数（顶层调用时） */
export const SPILL_FUNCTIONS = new Set(['FILTER', 'UNIQUE', 'SORT', 'SORTBY', 'SEQUENCE', 'TRANSPOSE', 'TEXTSPLIT', 'HSTACK', 'VSTACK']);

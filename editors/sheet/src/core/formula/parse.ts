/**
 * 公式语法：词法 + 运算符优先级解析成语法树。
 * 优先级（从低到高，同 Excel）：比较 = <> < > <= >=；连接 &；加减；乘除；乘方 ^；百分号 %；负号；区域 :。
 */
import { mapRefs, matchRef, type Ref } from './refs';

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: string }
  | { t: 'ref'; ref: Ref }
  | { t: 'name'; name: string }
  | { t: 'fn'; name: string; args: Node[] }
  | { t: 'un'; op: '-' | '+' | '%'; a: Node }
  | { t: 'bin'; op: string; a: Node; b: Node }
  | { t: 'arr'; rows: Node[][] }
  | { t: 'missing' };

export class ParseError extends Error {}

type Tok =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'err'; v: string }
  | { k: 'ref'; ref: Ref }
  | { k: 'fn'; v: string }
  | { k: 'name'; v: string }
  | { k: 'op'; v: string };

const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let s = '';
      for (;;) {
        if (j >= src.length) throw new ParseError('字符串没有结束的引号');
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"';
            j += 2;
            continue;
          }
          break;
        }
        s += src[j++];
      }
      out.push({ k: 'str', v: s });
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const e = ERRORS.find((x) => src.slice(i, i + x.length).toUpperCase() === x);
      if (!e) throw new ParseError('无法识别的错误值');
      out.push({ k: 'err', v: e });
      i += e.length;
      continue;
    }
    // 引用（可能带表名）优先于数字：1:3 是整行引用
    const m = matchRef(src, i);
    if (m) {
      out.push({ k: 'ref', ref: m.ref });
      i += m.len;
      continue;
    }
    const num = src.slice(i).match(/^(\d+\.?\d*|\.\d+)(e[-+]?\d+)?/i);
    if (num) {
      out.push({ k: 'num', v: Number(num[0]) });
      i += num[0].length;
      continue;
    }
    const word = src.slice(i).match(/^[\p{L}_\\][\p{L}\p{N}_.]*/u);
    if (word) {
      let j = i + word[0].length;
      while (src[j] === ' ') j++;
      if (src[j] === '(') out.push({ k: 'fn', v: word[0].toUpperCase().replace(/^_XLFN\.|^_XLWS\./, '') });
      else out.push({ k: 'name', v: word[0] });
      i += word[0].length;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      out.push({ k: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%(),;:{}'.includes(ch)) {
      out.push({ k: 'op', v: ch });
      i++;
      continue;
    }
    throw new ParseError(`无法识别的字符 ${ch}`);
  }
  return out;
}

const BIN: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export function parseFormula(src: string): Node {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const isOp = (v: string) => peek()?.k === 'op' && (peek() as { v: string }).v === v;
  const expect = (v: string) => {
    if (!isOp(v)) throw new ParseError(`缺少 ${v}`);
    p++;
  };

  const expr = (minPrec: number): Node => {
    let a = unary();
    for (;;) {
      const t = peek();
      if (t?.k !== 'op' || !(t.v in BIN)) break;
      const prec = BIN[t.v];
      if (prec < minPrec) break;
      p++;
      // ^ 在 Excel 里也是左结合
      const b = expr(prec + 1);
      a = { t: 'bin', op: t.v, a, b };
    }
    return a;
  };
  const unary = (): Node => {
    if (isOp('-') || isOp('+')) {
      const op = (toks[p++] as { v: '-' | '+' }).v;
      // 负号优先级高于 ^ 以外的运算（Excel：-2^2 = 4）
      return { t: 'un', op, a: unary() };
    }
    return postfix();
  };
  const postfix = (): Node => {
    let a = range();
    while (isOp('%')) {
      p++;
      a = { t: 'un', op: '%', a };
    }
    return a;
  };
  const range = (): Node => {
    let a = primary();
    while (isOp(':')) {
      p++;
      const b = primary();
      a = { t: 'bin', op: ':', a, b };
    }
    return a;
  };
  const primary = (): Node => {
    const t = toks[p++];
    if (!t) throw new ParseError('公式不完整');
    switch (t.k) {
      case 'num':
        return { t: 'num', v: t.v };
      case 'str':
        return { t: 'str', v: t.v };
      case 'err':
        return { t: 'err', v: t.v };
      case 'ref':
        return { t: 'ref', ref: t.ref };
      case 'name': {
        const up = t.v.toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') return { t: 'bool', v: up === 'TRUE' };
        return { t: 'name', name: t.v };
      }
      case 'fn': {
        expect('(');
        const args: Node[] = [];
        if (isOp(')')) {
          p++;
          return { t: 'fn', name: t.v, args };
        }
        for (;;) {
          if (isOp(',') || isOp(')')) args.push({ t: 'missing' });
          else args.push(expr(1));
          if (isOp(',')) {
            p++;
            if (isOp(')')) {
              args.push({ t: 'missing' });
              p++;
              break;
            }
            continue;
          }
          expect(')');
          break;
        }
        return { t: 'fn', name: t.v, args };
      }
      case 'op':
        if (t.v === '(') {
          const e = expr(1);
          expect(')');
          return e;
        }
        if (t.v === '{') {
          const rows: Node[][] = [[]];
          for (;;) {
            rows[rows.length - 1].push(unary());
            if (isOp(',')) p++;
            else if (isOp(';')) {
              p++;
              rows.push([]);
            } else {
              expect('}');
              break;
            }
          }
          return { t: 'arr', rows };
        }
        throw new ParseError(`多余的 ${t.v}`);
    }
  };
  const root = expr(1);
  if (p < toks.length) throw new ParseError('公式里有多余的内容');
  return root;
}

/** Excel 2010 以后新增的函数：写进文件时要带 _xlfn. 前缀，否则旧版 / Excel 自己读出来是 #NAME? */
export const FUTURE_FUNCTIONS = new Set([
  'IFS', 'IFNA', 'CONCAT', 'TEXTJOIN', 'XLOOKUP', 'XMATCH', 'MAXIFS', 'MINIFS', 'SWITCH', 'DAYS', 'ISOWEEKNUM', 'FILTER', 'SORT', 'UNIQUE',
  'SEQUENCE', 'LET', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'CEILING.MATH', 'FLOOR.MATH', 'RANK.EQ', 'RANK.AVG',
  'SORTBY', 'TEXTSPLIT', 'TEXTBEFORE', 'TEXTAFTER', 'HSTACK', 'VSTACK', 'UNICHAR', 'UNICODE', 'PERCENTILE.INC', 'QUARTILE.INC', 'MODE.SNGL',
]);
/** 这几个在文件里的前缀是 _xlfn._xlws.（工作表函数），不是 _xlfn. */
const WORKSHEET_FUNCTIONS = new Set(['FILTER', 'SORT']);

/** 改写公式里的函数名（跳过字符串） */
function mapFunctionNames(f: string, fn: (name: string) => string): string {
  let out = '';
  let i = 0;
  while (i < f.length) {
    if (f[i] === '"') {
      let j = i + 1;
      while (j < f.length && !(f[j] === '"' && f[j + 1] !== '"')) j += f[j] === '"' ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (f[i] === "'") {
      let j = i + 1;
      while (j < f.length && !(f[j] === "'" && f[j + 1] !== "'")) j += f[j] === "'" ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const w = f.slice(i).match(/^[\p{L}_\\][\p{L}\p{N}_.]*(?=\s*\()/u);
    if (w && !/[\p{L}\p{N}_.]/u.test(f[i - 1] ?? '')) {
      out += fn(w[0]);
      i += w[0].length;
      continue;
    }
    out += f[i++];
  }
  return out;
}

/** 文件里的公式 → 编辑器显示的公式（去掉 _xlfn. / _xlws. 前缀） */
export function fromFileFormula(f: string): string {
  if (!f.includes('_xl')) return f;
  return mapFunctionNames(f, (n) => n.replace(/^_xlfn\.|^_xlws\./i, ''));
}

/** 编辑器的公式 → 写进文件的公式（新函数加 _xlfn. 前缀） */
export function toFileFormula(f: string): string {
  return mapFunctionNames(f, (n) => (FUTURE_FUNCTIONS.has(n.toUpperCase()) ? `_xlfn.${WORKSHEET_FUNCTIONS.has(n.toUpperCase()) ? '_xlws.' : ''}${n.toUpperCase()}` : n));
}

/** 用户输入的公式规范化：引用和函数名转大写（=sum(a1:b2) → SUM(A1:B2)） */
export function normalizeFormula(f: string): string {
  return mapFunctionNames(mapRefs(f, (r) => r), (n) => n.toUpperCase());
}

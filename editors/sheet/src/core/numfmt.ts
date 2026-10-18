/**
 * Excel 数字格式（只读显示用的常见子集）：常规、小数位、千分位、百分比、科学计数、
 * 货币等前后缀文字、正负零分段，以及日期 / 时间（含中文“年月日”“星期”）。
 */

/** 内置格式（中文区域设置下的显示） */
const BUILTIN: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '¥#,##0;¥-#,##0',
  6: '¥#,##0;[Red]¥-#,##0',
  7: '¥#,##0.00;¥-#,##0.00',
  8: '¥#,##0.00;[Red]¥-#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'yyyy/m/d',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'yyyy/m/d h:mm',
  27: 'yyyy"年"m"月"',
  28: 'm"月"d"日"',
  29: 'm"月"d"日"',
  30: 'm-d-yy',
  31: 'yyyy"年"m"月"d"日"',
  32: 'h"时"mm"分"',
  33: 'h"时"mm"分"ss"秒"',
  34: '上午/下午h"时"mm"分"',
  35: '上午/下午h"时"mm"分"ss"秒"',
  36: 'yyyy"年"m"月"',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
  50: 'yyyy"年"m"月"',
  51: 'm"月"d"日"',
  52: 'yyyy"年"m"月"',
  53: 'm"月"d"日"',
  54: 'm"月"d"日"',
  55: '上午/下午h"时"mm"分"',
  56: '上午/下午h"时"mm"分"ss"秒"',
  57: 'yyyy"年"m"月"',
  58: 'm"月"d"日"',
};

export function builtinFormat(id: number): string {
  return BUILTIN[id] ?? 'General';
}

/** 按 `;` 分段（引号、反斜杠转义内的不算） */
function sections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') quoted = !quoted;
    if (!quoted && ch === '\\') {
      cur += ch + (code[i + 1] ?? '');
      i++;
      continue;
    }
    if (!quoted && ch === ';') {
      out.push(cur);
      cur = '';
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

/** 去掉引号、转义和方括号，只留下能判断格式类别的字符 */
function bare(section: string): string {
  return section
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/_./g, '')
    .replace(/\*./g, '')
    .replace(/\[(?!h+\]|m+\]|s+\])[^\]]*\]/gi, '');
}

const DATE_RE = /[ymdhse]|aaa|\[(h+|m+|s+)\]|上午\/下午|am\/pm|a\/p/i;

/** 分段是否是日期 / 时间格式（科学计数的 E+00 不算） */
function dateLike(section: string): boolean {
  return DATE_RE.test(bare(section).replace(/general/gi, '').replace(/e[+-]/gi, ''));
}

export function isDateFormat(code: string): boolean {
  return dateLike(sections(code)[0]);
}

/** “常规”格式：最多 11 位有效数字 */
export function formatGeneral(v: number): string {
  if (!Number.isFinite(v)) return String(v);
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  const abs = Math.abs(v);
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) {
    return v
      .toExponential(5)
      .replace(/\.?0+e/, 'e')
      .replace(/e([+-])(\d)$/, 'e$10$2')
      .toUpperCase();
  }
  return String(Number(v.toPrecision(11)));
}

/** 取掉转义后的文字前后缀 */
function literal(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      out += s.slice(i + 1, end < 0 ? undefined : end);
      i = end < 0 ? s.length : end;
    } else if (ch === '\\') {
      out += s[i + 1] ?? '';
      i++;
    } else if (ch === '_') {
      out += ' ';
      i++;
    } else if (ch === '*') {
      i++;
    } else if (ch === '[') {
      const end = s.indexOf(']', i);
      const inner = s.slice(i + 1, end < 0 ? undefined : end);
      // [$¥-804] → ¥；[Red]、[$-804]、[>=100] 等忽略
      const cur = inner.match(/^\$([^-]*)/);
      if (cur) out += cur[1];
      i = end < 0 ? s.length : end;
    } else out += ch;
  }
  return out;
}

function groupThousands(int: string): string {
  return int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatNumberSection(v: number, section: string): string {
  // 找出数字占位部分：第一个 0#? 到最后一个 0#?.,（引号 / 转义 / 方括号以外）
  const mask: boolean[] = [];
  let quoted = false;
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      quoted = !quoted;
      mask.push(false);
    } else if (quoted) mask.push(false);
    else if (ch === '\\' || ch === '_' || ch === '*') {
      mask.push(false, false);
      i++;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      const n = (end < 0 ? section.length : end) - i + 1;
      for (let k = 0; k < n; k++) mask.push(false);
      i += n - 1;
    } else mask.push(true);
  }
  const isPh = (i: number) => mask[i] && /[0#?]/.test(section[i]);
  let start = -1;
  let end = -1;
  // 科学计数 E+00 的指数位不算尾数占位符
  let limit = section.length;
  for (let i = 0; i < section.length; i++) {
    if (mask[i] && /[eE]/.test(section[i]) && /[+-]/.test(section[i + 1] ?? '')) {
      limit = i;
      break;
    }
  }
  for (let i = 0; i < limit; i++) {
    if (isPh(i)) {
      if (start < 0) start = i;
      end = i;
    }
  }
  if (start < 0) return literal(section);
  // 占位符之后紧跟的 . , E+00 仍属于数字部分
  while (end + 1 < limit && mask[end + 1] && /[0#?.,]/.test(section[end + 1])) end++;
  let pattern = section.slice(start, end + 1);
  let prefix = section.slice(0, start);
  let suffix = section.slice(end + 1);
  const expo = suffix.match(/^[eE]([+-])([0#]+)/) ?? null;
  if (expo) suffix = suffix.slice(expo[0].length);
  // 前缀里有小数点（如 .00）也算数字部分
  if (prefix.endsWith('.') && !prefix.endsWith('\\.')) {
    pattern = `.${pattern}`;
    prefix = prefix.slice(0, -1);
  }

  let x = v;
  const pct = (bare(prefix + suffix).match(/%/g) ?? []).length;
  x *= 100 ** pct;
  // 末尾逗号：每个缩小 1000 倍
  const trailing = pattern.match(/,+$/)?.[0].length ?? 0;
  if (trailing) {
    x /= 1000 ** trailing;
    pattern = pattern.slice(0, -trailing);
  }
  const dot = pattern.indexOf('.');
  const intPat = dot < 0 ? pattern : pattern.slice(0, dot);
  const fracPat = dot < 0 ? '' : pattern.slice(dot + 1);
  const maxDec = (fracPat.match(/[0#?]/g) ?? []).length;
  const minDec = (fracPat.match(/0/g) ?? []).length;
  const minInt = (intPat.match(/0/g) ?? []).length;
  const thousands = intPat.includes(',');

  let body: string;
  if (expo) {
    const [m, e] = x.toExponential(maxDec).split('e');
    const exp = Number(e);
    const sign = exp < 0 ? '-' : expo[1] === '+' ? '+' : '';
    body = `${m}E${sign}${String(Math.abs(exp)).padStart(expo[2].length, '0')}`;
  } else {
    let [int, frac = ''] = Math.abs(x).toFixed(maxDec).split('.');
    while (frac.length > minDec && frac.endsWith('0')) frac = frac.slice(0, -1);
    if (int === '0' && minInt === 0) int = '';
    int = int.padStart(minInt, '0');
    if (thousands) int = groupThousands(int);
    body = frac ? `${int}.${frac}` : int || '0';
    if (int === '' && frac) body = `.${frac}`;
  }
  const neg = !expo && x < 0 && Number(Math.abs(x).toFixed(maxDec)) !== 0;
  return (neg ? '-' : '') + literal(prefix) + body + literal(suffix);
}

const WEEK_CN = ['日', '一', '二', '三', '四', '五', '六'];
const WEEK_EN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MONTH_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

type Tok = { k: 'lit'; s: string } | { k: 'y' | 'M' | 'd' | 'h' | 'm' | 's' | 'ampm' | 'cn-ampm' | 'aaa' | 'frac' | 'H' | 'N' | 'S'; n: number };

function dateTokens(section: string): Tok[] {
  const toks: Tok[] = [];
  const s = section;
  for (let i = 0; i < s.length; ) {
    const ch = s[i];
    const low = ch.toLowerCase();
    const run = (c: string) => {
      let n = 0;
      while (i + n < s.length && s[i + n].toLowerCase() === c) n++;
      return n;
    };
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      toks.push({ k: 'lit', s: s.slice(i + 1, end < 0 ? undefined : end) });
      i = end < 0 ? s.length : end + 1;
    } else if (ch === '\\') {
      toks.push({ k: 'lit', s: s[i + 1] ?? '' });
      i += 2;
    } else if (ch === '_') {
      toks.push({ k: 'lit', s: ' ' });
      i += 2;
    } else if (ch === '*') i += 2;
    else if (ch === '[') {
      const end = s.indexOf(']', i);
      const inner = s.slice(i + 1, end < 0 ? undefined : end).toLowerCase();
      if (/^h+$/.test(inner)) toks.push({ k: 'H', n: inner.length });
      else if (/^m+$/.test(inner)) toks.push({ k: 'N', n: inner.length });
      else if (/^s+$/.test(inner)) toks.push({ k: 'S', n: inner.length });
      i = end < 0 ? s.length : end + 1;
    } else if (s.startsWith('上午/下午', i)) {
      toks.push({ k: 'cn-ampm', n: 1 });
      i += 5;
    } else if (/^am\/pm/i.test(s.slice(i))) {
      toks.push({ k: 'ampm', n: 2 });
      i += 5;
    } else if (/^a\/p/i.test(s.slice(i))) {
      toks.push({ k: 'ampm', n: 1 });
      i += 3;
    } else if (low === 'y' || low === 'e') {
      const n = low === 'e' ? 4 : run('y');
      toks.push({ k: 'y', n });
      i += low === 'e' ? 1 : n;
    } else if (low === 'm') {
      const n = run('m');
      toks.push({ k: 'M', n });
      i += n;
    } else if (low === 'd') {
      const n = run('d');
      toks.push({ k: 'd', n });
      i += n;
    } else if (low === 'a' && run('a') >= 3) {
      const n = run('a');
      toks.push({ k: 'aaa', n });
      i += n;
    } else if (low === 'h') {
      const n = run('h');
      toks.push({ k: 'h', n });
      i += n;
    } else if (low === 's') {
      const n = run('s');
      toks.push({ k: 's', n });
      i += n;
    } else if (ch === '.' && s[i + 1] === '0' && toks.some((t) => t.k === 's' || t.k === 'S')) {
      let n = 0;
      while (s[i + 1 + n] === '0') n++;
      toks.push({ k: 'frac', n });
      i += 1 + n;
    } else {
      toks.push({ k: 'lit', s: ch });
      i++;
    }
  }
  // m / mm 紧跟在小时之后或在秒之前时表示分钟
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.k !== 'M' || t.n > 2) continue;
    const prev = toks.slice(0, i).reverse().find((x) => x.k !== 'lit');
    const next = toks.slice(i + 1).find((x) => x.k !== 'lit');
    if (prev?.k === 'h' || prev?.k === 'H' || next?.k === 's' || next?.k === 'S') toks[i] = { k: 'm', n: t.n };
  }
  return toks;
}

/** Excel 日期序列号 → 各部分（UTC） */
function serialParts(serial: number, date1904: boolean) {
  const days = date1904 ? serial + 1462 : serial;
  const ms = Math.round((days - 25569) * 86400000);
  const d = new Date(ms);
  return {
    y: d.getUTCFullYear(),
    M: d.getUTCMonth() + 1,
    d: d.getUTCDate(),
    w: d.getUTCDay(),
    h: d.getUTCHours(),
    m: d.getUTCMinutes(),
    s: d.getUTCSeconds(),
    ms: d.getUTCMilliseconds(),
  };
}

function formatDateSection(v: number, section: string, date1904: boolean): string {
  const toks = dateTokens(section);
  const hasFrac = toks.some((t) => t.k === 'frac');
  // 不显示小数秒时按秒四舍五入
  const serial = hasFrac ? v : Math.round(v * 86400) / 86400;
  if (serial < 0) return formatGeneral(v);
  const p = serialParts(serial, date1904);
  const twelve = toks.some((t) => t.k === 'ampm' || t.k === 'cn-ampm');
  const pad = (n: number, w: number) => String(n).padStart(w, '0');
  let out = '';
  for (const t of toks) {
    switch (t.k) {
      case 'lit':
        out += t.s;
        break;
      case 'y':
        out += t.n <= 2 ? pad(p.y % 100, 2) : String(p.y);
        break;
      case 'M':
        out += t.n === 1 ? String(p.M) : t.n === 2 ? pad(p.M, 2) : t.n === 3 ? MONTH_EN[p.M - 1].slice(0, 3) : t.n === 5 ? MONTH_EN[p.M - 1][0] : MONTH_EN[p.M - 1];
        break;
      case 'd':
        out += t.n === 1 ? String(p.d) : t.n === 2 ? pad(p.d, 2) : t.n === 3 ? WEEK_EN[p.w].slice(0, 3) : WEEK_EN[p.w];
        break;
      case 'aaa':
        out += t.n === 3 ? WEEK_CN[p.w] : `星期${WEEK_CN[p.w]}`;
        break;
      case 'h': {
        const h = twelve ? p.h % 12 || 12 : p.h;
        out += t.n >= 2 ? pad(h, 2) : String(h);
        break;
      }
      case 'm':
        out += t.n >= 2 ? pad(p.m, 2) : String(p.m);
        break;
      case 's':
        out += t.n >= 2 ? pad(p.s, 2) : String(p.s);
        break;
      case 'frac':
        out += `.${pad(p.ms, 3).slice(0, t.n)}`;
        break;
      case 'ampm':
        out += t.n === 1 ? (p.h < 12 ? 'A' : 'P') : p.h < 12 ? 'AM' : 'PM';
        break;
      case 'cn-ampm':
        out += p.h < 12 ? '上午' : '下午';
        break;
      case 'H':
        out += pad(Math.floor(serial * 24), t.n);
        break;
      case 'N':
        out += pad(Math.floor(serial * 1440), t.n);
        break;
      case 'S':
        out += pad(Math.floor(serial * 86400), t.n);
        break;
    }
  }
  return out;
}

/** 按格式代码显示数值 */
export function formatNumber(v: number, code: string, date1904 = false): string {
  const secs = sections(code);
  let sec = secs[0];
  let x = v;
  if (secs.length > 1 && v < 0) {
    sec = secs[1];
    x = -v;
  } else if (secs.length > 2 && v === 0) sec = secs[2];
  const trimmed = bare(sec).trim();
  // 只有文字的分段（如零值显示 "-"）
  if (!trimmed) return literal(sec);
  // 只有 ? 占位的分段（会计格式的零值 "-"??）：? 显示为空格
  if (/^\?+$/.test(trimmed)) return literal(sec.replace(/\?/g, ' '));
  if (trimmed.includes('@')) return formatGeneral(v);
  if (/general/i.test(trimmed)) return literal(sec.replace(/general/i, '\u0000')).replace('\u0000', formatGeneral(x));
  if (dateLike(sec)) return formatDateSection(x, sec, date1904);
  if (/\?\/\?/.test(sec)) return formatGeneral(v);
  return formatNumberSection(x, sec);
}

/**
 * 邮件合并域的格式开关（Word 域的 \# 数字图片、\@ 日期时间图片、\* 大小写）。纯函数，不依赖文档。
 * 值不是数字 / 日期时原样返回（Word 也是不转换）。
 */

// ---------------- 数字 \# ----------------

/** 取数字：去掉千分位逗号、空格、常见货币符号和百分号后能解析的才算 */
export function parseNumber(v: string): number | null {
  const s = v.trim().replace(/[,，\s¥￥$€£]/g, '').replace(/%$/, '');
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(s)) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

/** 数字图片的一节（正数；负数；零，用 ; 分隔）：前后字面文字、整数位、小数位、千分位 */
interface NumPic {
  prefix: string;
  suffix: string;
  /** 整数部分至少几位（0 的个数） */
  minInt: number;
  minDec: number;
  maxDec: number;
  group: boolean;
  /** 这一节里没有数字位（如“零”）：只输出文字 */
  literalOnly?: boolean;
}

function splitPictures(pic: string): string[] {
  // 引号里的 ; 不分节
  const out: string[] = [];
  let cur = '';
  let q = '';
  for (const c of pic) {
    if (q) {
      if (c === q) q = '';
      cur += c;
    } else if (c === "'" || c === '"') {
      q = c;
      cur += c;
    } else if (c === ';') {
      out.push(cur);
      cur = '';
    } else cur += c;
  }
  out.push(cur);
  return out;
}

function parseNumPic(pic: string): NumPic {
  // 第一个 # 或 0 开始到最后一个 # 0 , . 为止是数字部分，前后是字面文字（引号里的原样）
  const lit = (s: string) => s.replace(/'([^']*)'|"([^"]*)"/g, (_, a: string, b: string) => a ?? b);
  const m = pic.match(/^([^#0]*)([#0][#0,.]*)(.*)$/s);
  if (!m) return { prefix: lit(pic), suffix: '', minInt: 0, minDec: 0, maxDec: 0, group: false, literalOnly: true };
  const [, prefix, core, suffix] = m;
  const dot = core.indexOf('.');
  const ip = dot < 0 ? core : core.slice(0, dot);
  const dp = dot < 0 ? '' : core.slice(dot + 1).replace(/,/g, '');
  return {
    prefix: lit(prefix),
    suffix: lit(suffix),
    minInt: (ip.match(/0/g) ?? []).length,
    minDec: (dp.match(/0/g) ?? []).length,
    maxDec: dp.length,
    group: ip.includes(','),
  };
}

function renderNum(abs: number, p: NumPic): string {
  if (p.literalOnly) return p.prefix;
  let s = abs.toFixed(p.maxDec);
  if (p.maxDec > p.minDec) {
    // 去掉多余的小数末尾 0，但保留 minDec 位
    const [i, d = ''] = s.split('.');
    let dd = d;
    while (dd.length > p.minDec && dd.endsWith('0')) dd = dd.slice(0, -1);
    s = dd ? `${i}.${dd}` : i;
  }
  let [int, dec] = s.split('.');
  if (int === '0' && p.minInt === 0 && dec) int = '';
  else int = int.padStart(p.minInt, '0');
  if (p.group) int = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return `${p.prefix}${int}${dec ? `.${dec}` : ''}${p.suffix}`;
}

/** Word 数字图片：0 必有的位、# 可有可无的位、, 千分位、. 小数点，其它字符照写；可用 ; 分正数 / 负数 / 零 */
export function formatNumber(value: string, pic: string): string {
  const n = parseNumber(value);
  if (n === null) return value;
  const parts = splitPictures(pic);
  if (n === 0 && parts[2] !== undefined) return renderNum(0, parseNumPic(parts[2]));
  if (n < 0) {
    if (parts[1] !== undefined) return renderNum(-n, parseNumPic(parts[1]));
    return `-${renderNum(-n, parseNumPic(parts[0]))}`;
  }
  return renderNum(n, parseNumPic(parts[0]));
}

// ---------------- 日期 \@ ----------------

const CN_NUM = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二'];
const CN_DAY = ['日', '一', '二', '三', '四', '五', '六'];

interface When {
  y: number;
  mo: number;
  d: number;
  h: number;
  mi: number;
  s: number;
  /** 值里写了时间 */
  hasTime: boolean;
}

/** 认日期：2026-10-04、2026/10/4、2026.10.4、2026年10月4日、20261004，后面可带 时:分(:秒)；Excel 序列号（天数，可带小数） */
export function parseWhen(v: string): When | null {
  const s = v.trim();
  const time = (rest: string) => {
    const t = rest.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    return t ? { h: +t[1], mi: +t[2], s: t[3] ? +t[3] : 0, hasTime: true } : { h: 0, mi: 0, s: 0, hasTime: false };
  };
  const ok = (y: number, mo: number, d: number, rest: string): When | null => {
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const dt = new Date(Date.UTC(y, mo - 1, d));
    if (dt.getUTCMonth() !== mo - 1) return null;
    return { y, mo, d, ...time(rest) };
  };
  let m = s.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(.*)$/);
  if (m) return ok(+m[1], +m[2], +m[3], m[4]);
  m = s.match(/^(\d{4})(\d{2})(\d{2})(.*)$/);
  if (m && s.length >= 8 && !/^\d{9,}/.test(s)) return ok(+m[1], +m[2], +m[3], m[4]);
  if (/^\d{4,6}(\.\d+)?$/.test(s)) {
    const serial = Number(s);
    if (serial >= 1 && serial < 2958466) {
      const ms = Math.round((serial - 25569) * 86400000);
      const dt = new Date(ms);
      return { y: dt.getUTCFullYear(), mo: dt.getUTCMonth() + 1, d: dt.getUTCDate(), h: dt.getUTCHours(), mi: dt.getUTCMinutes(), s: dt.getUTCSeconds(), hasTime: serial % 1 !== 0 };
    }
  }
  return null;
}

const two = (n: number) => String(n).padStart(2, '0');

/** Word 日期时间图片：yyyy yy MMMM MMM MM M dddd ddd dd d HH H hh h mm m ss s AM/PM am/pm；引号里的文字照写，其它字符（年 月 日 - / 等）照写。月份、星期用中文名 */
export function formatDate(value: string, pic: string): string {
  const w = parseWhen(value);
  if (!w) return value;
  const dow = new Date(Date.UTC(w.y, w.mo - 1, w.d)).getUTCDay();
  const h12 = w.h % 12 === 0 ? 12 : w.h % 12;
  return pic.replace(/yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|AM\/PM|am\/pm|A\/P|a\/p|'[^']*'|"[^"]*"|[\s\S]/g, (t) => {
    switch (t) {
      case 'yyyy': return String(w.y);
      case 'yy': return two(w.y % 100);
      case 'MMMM': return `${CN_NUM[w.mo - 1]}月`;
      case 'MMM': return `${w.mo}月`;
      case 'MM': return two(w.mo);
      case 'M': return String(w.mo);
      case 'dddd': return `星期${CN_DAY[dow]}`;
      case 'ddd': return `周${CN_DAY[dow]}`;
      case 'dd': return two(w.d);
      case 'd': return String(w.d);
      case 'HH': return two(w.h);
      case 'H': return String(w.h);
      case 'hh': return two(h12);
      case 'h': return String(h12);
      case 'mm': return two(w.mi);
      case 'm': return String(w.mi);
      case 'ss': return two(w.s);
      case 's': return String(w.s);
      case 'AM/PM': return w.h < 12 ? 'AM' : 'PM';
      case 'am/pm': return w.h < 12 ? 'am' : 'pm';
      case 'A/P': return w.h < 12 ? 'A' : 'P';
      case 'a/p': return w.h < 12 ? 'a' : 'p';
      default: return t[0] === "'" || t[0] === '"' ? t.slice(1, -1) : t;
    }
  });
}

// ---------------- 大小写 \* ----------------

export type CaseFmt = 'Upper' | 'Lower' | 'FirstCap' | 'Caps';

export function applyCase(value: string, fmt: string): string {
  switch (fmt.toLowerCase()) {
    case 'upper':
      return value.toUpperCase();
    case 'lower':
      return value.toLowerCase();
    case 'firstcap':
      return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
    case 'caps':
      return value.toLowerCase().replace(/(^|[^\p{L}\p{N}'])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase());
    default:
      return value;
  }
}

export interface FormatSwitches {
  numFmt?: string;
  dateFmt?: string;
  caseFmt?: string;
}

/** 按开关格式化一个值：先数字 / 日期，再大小写（与 Word 的顺序一致）；数字图片和日期图片都给了时，值像日期用日期，否则用数字 */
export function formatValue(value: string, f: FormatSwitches): string {
  let v = value;
  if (v !== '') {
    if (f.dateFmt && (!f.numFmt || parseWhen(v))) v = formatDate(v, f.dateFmt);
    else if (f.numFmt) v = formatNumber(v, f.numFmt);
  }
  if (f.caseFmt) v = applyCase(v, f.caseFmt);
  return v;
}

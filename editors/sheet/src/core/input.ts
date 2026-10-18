/**
 * 输入框里的文字 → 单元格内容（同 Excel 的自动识别）：
 * 公式（= 开头）、数字（含千分位、科学计数）、百分比、货币 ¥、日期 / 时间、TRUE / FALSE、' 开头强制为文字。
 * 与 Excel 不同的一点：前导 0 的数字（编号）和超过 15 位的数字（身份证号、订单号）保持文字，不会被改掉。
 */
import { isDateFormat } from './numfmt';

export type Parsed =
  | { kind: 'empty' }
  | { kind: 'formula'; f: string }
  | { kind: 'value'; v: string | number | boolean; fmt?: string };

/** 1900 日期系统的序列号（Excel 把 1900-02-29 当成存在的日子，1900-03-01 起与真实日期一致） */
export function dateSerial(y: number, m: number, d: number, date1904 = false): number {
  const utc = Date.UTC(y, m - 1, d);
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return Math.round((utc - base) / 86400000);
}

const validDate = (y: number, m: number, d: number) => {
  if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 9999) return false;
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
};

function parseTime(s: string): { frac: number; seconds: boolean } | null {
  const m = s.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (!m) return null;
  const h = Number(m[1]);
  const mi = Number(m[2]);
  const se = Number(m[3] ?? 0);
  if (h > 23 || mi > 59 || se > 59) return null;
  return { frac: (h * 3600 + mi * 60 + se) / 86400, seconds: m[3] !== undefined };
}

function parseDate(s: string, date1904: boolean, year: number): { serial: number; fmt: string } | null {
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (m && validDate(+m[1], +m[2], +m[3])) return { serial: dateSerial(+m[1], +m[2], +m[3], date1904), fmt: 'yyyy/m/d' };
  m = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if (m && validDate(+m[1], +m[2], +m[3])) return { serial: dateSerial(+m[1], +m[2], +m[3], date1904), fmt: 'yyyy"年"m"月"d"日"' };
  m = s.match(/^(\d{4})年(\d{1,2})月$/);
  if (m && validDate(+m[1], +m[2], 1)) return { serial: dateSerial(+m[1], +m[2], 1, date1904), fmt: 'yyyy"年"m"月"' };
  // 月/日（当年）
  m = s.match(/^(\d{1,2})[/-](\d{1,2})$/);
  if (m && validDate(year, +m[1], +m[2])) return { serial: dateSerial(year, +m[1], +m[2], date1904), fmt: 'm"月"d"日"' };
  m = s.match(/^(\d{1,2})月(\d{1,2})日$/);
  if (m && validDate(year, +m[1], +m[2])) return { serial: dateSerial(year, +m[1], +m[2], date1904), fmt: 'm"月"d"日"' };
  return null;
}

const NUM = /^[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?(?:[eE][-+]?\d+)?$/;

function parseNumber(s: string): { v: number; thousands: boolean; decimals: number } | null {
  if (!/\d/.test(s) || !NUM.test(s)) return null;
  const digits = s.replace(/[-+,.]|[eE].*$/g, '');
  // 前导 0 的编号、超过 15 位的长数字按文字
  if (/^[-+]?0\d/.test(s.replace(/,/g, '')) || digits.length > 15) return null;
  const v = Number(s.replace(/,/g, ''));
  if (!Number.isFinite(v)) return null;
  return { v, thousands: s.includes(','), decimals: s.match(/\.(\d+)(?![\d])/)?.[1].length ?? 0 };
}

/**
 * currentFmt：格子现有的数字格式。“文本”格式（@）下一律当文字；
 * 已有非“常规”格式时数字沿用它，不再套新格式（同 Excel）。
 */
export function parseInput(raw: string, currentFmt = 'General', date1904 = false, now = new Date()): Parsed {
  if (raw === '') return { kind: 'empty' };
  if (currentFmt === '@') return { kind: 'value', v: raw };
  if (raw.startsWith("'")) return raw.length > 1 ? { kind: 'value', v: raw.slice(1) } : { kind: 'empty' };
  if (raw.startsWith('=') && raw.length > 1) return { kind: 'formula', f: raw.slice(1) };
  const s = raw.trim();
  if (!s) return { kind: 'value', v: raw };
  const keep = currentFmt !== 'General';
  const upper = s.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') return { kind: 'value', v: upper === 'TRUE' };

  const n = parseNumber(s);
  if (n) {
    if (keep || !n.thousands) return { kind: 'value', v: n.v };
    return { kind: 'value', v: n.v, fmt: n.decimals ? '#,##0.00' : '#,##0' };
  }
  const pct = s.match(/^(.+)%$/);
  if (pct) {
    const p = parseNumber(pct[1].trim());
    if (p) {
      const v = Number((p.v / 100).toPrecision(15));
      return keep ? { kind: 'value', v } : { kind: 'value', v, fmt: p.decimals ? `0.${'0'.repeat(p.decimals)}%` : '0%' };
    }
  }
  const cur = s.match(/^([-+]?)[¥￥]\s*(.+)$/);
  if (cur) {
    const p = parseNumber(cur[2]);
    if (p) {
      const v = cur[1] === '-' ? -p.v : p.v;
      return keep ? { kind: 'value', v } : { kind: 'value', v, fmt: p.decimals ? '"¥"#,##0.00' : '"¥"#,##0' };
    }
  }
  const time = parseTime(s);
  if (time) return { kind: 'value', v: time.frac, fmt: keep && isDateFormat(currentFmt) ? undefined : time.seconds ? 'h:mm:ss' : 'h:mm' };
  const [dPart, tPart] = s.split(/\s+/, 2);
  const date = parseDate(dPart, date1904, now.getFullYear());
  if (date) {
    if (!tPart && s.split(/\s+/).length === 1) return { kind: 'value', v: date.serial, fmt: keep && isDateFormat(currentFmt) ? undefined : date.fmt };
    const t = tPart && s.split(/\s+/).length === 2 ? parseTime(tPart) : null;
    if (t) return { kind: 'value', v: date.serial + t.frac, fmt: keep && isDateFormat(currentFmt) ? undefined : t.seconds ? 'yyyy/m/d h:mm:ss' : 'yyyy/m/d h:mm' };
  }
  return { kind: 'value', v: raw };
}

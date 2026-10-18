/**
 * 格式工具栏的纯逻辑：常用数字格式、增减小数位、边框（按格子在区域里的位置决定改哪几条边）、调色板。
 */
import { formatGeneral, isDateFormat } from './numfmt';
import { applyTint } from './color';
import type { Range } from './workbook';
import type { StyleChange } from './styles';

/** 数字格式下拉里的选项（中文版 Excel 的“常规 / 数值 / 货币 / 会计专用 …”） */
export const NUMBER_FORMATS: { label: string; code: string; sample?: string }[] = [
  { label: '常规', code: 'General' },
  { label: '数值', code: '0.00' },
  { label: '千分位', code: '#,##0.00' },
  { label: '货币', code: '"¥"#,##0.00;"¥"-#,##0.00' },
  { label: '会计专用', code: '_ "¥"* #,##0.00_ ;_ "¥"* -#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ ' },
  { label: '百分比', code: '0.00%' },
  { label: '科学记数', code: '0.00E+00' },
  { label: '短日期', code: 'yyyy/m/d' },
  { label: '长日期', code: 'yyyy"年"m"月"d"日"' },
  { label: '时间', code: 'h:mm:ss' },
  { label: '文本', code: '@' },
];

/** 格式代码 → 下拉里显示的名字（不是常用格式时为“自定义”） */
export function formatLabel(code: string): string {
  const hit = NUMBER_FORMATS.find((f) => f.code === code);
  if (hit) return hit.label;
  if (isDateFormat(code)) return /[hs]/i.test(code.replace(/"[^"]*"/g, '')) && !/[yd]/i.test(code.replace(/"[^"]*"/g, '')) ? '时间' : '日期';
  if (/%/.test(code.replace(/"[^"]*"/g, ''))) return '百分比';
  return '自定义';
}

/** 一个分段里数字部分的小数位 +1 / -1（引号里的文字不动） */
function stepSection(sec: string, dir: 1 | -1): string {
  // 找到最后一个数字占位符（不在引号、方括号里）
  let quoted = false;
  let bracket = false;
  let last = -1;
  let dot = -1;
  for (let i = 0; i < sec.length; i++) {
    const ch = sec[i];
    if (ch === '"') quoted = !quoted;
    else if (!quoted && ch === '[') bracket = true;
    else if (!quoted && ch === ']') bracket = false;
    else if (!quoted && !bracket && (ch === '\\' || ch === '_' || ch === '*')) i++;
    else if (!quoted && !bracket && /[0#?]/.test(ch)) last = i;
    else if (!quoted && !bracket && ch === '.' && last >= 0) dot = i;
  }
  if (last < 0) return sec;
  if (dir > 0) return dot >= 0 ? sec.slice(0, last + 1) + '0' + sec.slice(last + 1) : sec.slice(0, last + 1) + '.0' + sec.slice(last + 1);
  if (dot < 0) return sec;
  // 去掉最后一位小数；没有小数位了连小数点一起去掉
  const decimals = last - dot;
  if (decimals <= 1) return sec.slice(0, dot) + sec.slice(last + 1);
  return sec.slice(0, last) + sec.slice(last + 1);
}

/**
 * 增加 / 减少小数位（同 Excel 的按钮）：常规格式按当前值显示的位数开始；日期、文本格式不变。
 */
export function stepDecimals(code: string, dir: 1 | -1, sample?: number): string {
  if (code === '@' || isDateFormat(code)) return code;
  if (code === 'General') {
    const text = typeof sample === 'number' ? formatGeneral(sample) : '0';
    if (/E/i.test(text)) return dir > 0 ? '0.000E+00' : '0.0E+00';
    const d = text.split('.')[1]?.length ?? 0;
    const n = Math.max(0, d + dir);
    return n ? `0.${'0'.repeat(n)}` : '0';
  }
  // 按 ; 分段（引号里的不算），每段都改
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (const ch of code) {
    if (ch === '"') quoted = !quoted;
    if (ch === ';' && !quoted) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out.map((s) => stepSection(s, dir)).join(';');
}

export type BorderKind = 'all' | 'outside' | 'thickOutside' | 'none' | 'top' | 'bottom' | 'left' | 'right' | 'doubleBottom' | 'thickBottom';

export const BORDER_ITEMS: { kind: BorderKind; label: string }[] = [
  { kind: 'bottom', label: '下框线' },
  { kind: 'top', label: '上框线' },
  { kind: 'left', label: '左框线' },
  { kind: 'right', label: '右框线' },
  { kind: 'none', label: '无框线' },
  { kind: 'all', label: '所有框线' },
  { kind: 'outside', label: '外侧框线' },
  { kind: 'thickOutside', label: '粗外侧框线' },
  { kind: 'doubleBottom', label: '双底框线' },
  { kind: 'thickBottom', label: '粗下框线' },
];

/** 区域里 (r, c) 这一格要改的边；没有要改的返回 null */
export function borderFor(kind: BorderKind, g: Range, r: number, c: number, color?: string): StyleChange['border'] | null {
  const line = (style: string) => ({ style, ...(color && { color }) });
  const top = r === g.r0;
  const bottom = r === g.r1;
  const left = c === g.c0;
  const right = c === g.c1;
  const out: NonNullable<StyleChange['border']> = {};
  switch (kind) {
    case 'all':
      return { left: line('thin'), right: line('thin'), top: line('thin'), bottom: line('thin') };
    case 'none':
      return { left: null, right: null, top: null, bottom: null };
    case 'outside':
    case 'thickOutside': {
      const s = kind === 'outside' ? 'thin' : 'medium';
      if (top) out.top = line(s);
      if (bottom) out.bottom = line(s);
      if (left) out.left = line(s);
      if (right) out.right = line(s);
      break;
    }
    case 'top':
      if (top) out.top = line('thin');
      break;
    case 'bottom':
      if (bottom) out.bottom = line('thin');
      break;
    case 'left':
      if (left) out.left = line('thin');
      break;
    case 'right':
      if (right) out.right = line('thin');
      break;
    case 'doubleBottom':
      if (bottom) out.bottom = line('double');
      break;
    case 'thickBottom':
      if (bottom) out.bottom = line('medium');
      break;
  }
  return Object.keys(out).length ? out : null;
}

/** Excel 调色板：10 个主题色各一列（原色 + 4 档变浅 / 变深），下面一行标准色 */
const THEME_BASE = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
export const STANDARD_COLORS = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'].map((x) => `#${x.toLowerCase()}`);

export function themePalette(): string[][] {
  // 每行一个档：原色、浅 80%、浅 60%、浅 40%、深 25%、深 50%（白 / 黑两列 Excel 用的是 5%~50% 的深浅，这里统一按 tint）
  const tints = [0, 0.8, 0.6, 0.4, -0.25, -0.5];
  return tints.map((t) =>
    THEME_BASE.map((b, i) => {
      // 白色变浅没有意义：换成变深；黑色变深没有意义：换成变浅
      const tt = i === 0 && t > 0 ? -t / 4 : i === 1 && t < 0 ? -t : t;
      return `#${(t === 0 ? b : applyTint(b, tt)).toLowerCase()}`;
    }),
  );
}

/** 常用字体与字号（中文版 Excel 的列表） */
/** 字体下拉框的两组：中文字体（等线放最后）和西文字体 */
export const SHEET_FONTS_EA = ['宋体', '黑体', '微软雅黑', '楷体', '仿宋', '等线'];
export const SHEET_FONTS_LATIN = ['Arial', 'Times New Roman', 'Calibri'];
export const SHEET_FONTS = [...SHEET_FONTS_EA, ...SHEET_FONTS_LATIN];
export const SHEET_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** 字号对应的自动行高（磅）：中文 Excel 里 11 号等线约 14.25 磅，按比例放大 */
export function autoRowPt(fontSize: number): number {
  return Math.round(fontSize * 1.3 * 4) / 4;
}

// ---------- 设置单元格格式：数字分类 ----------

/** “设置单元格格式 → 数字”的分类（同 Excel 中文版） */
export type FmtCategory = 'general' | 'number' | 'currency' | 'accounting' | 'percent' | 'scientific' | 'date' | 'time' | 'text' | 'custom';

export const FMT_CATEGORIES: { cat: FmtCategory; label: string }[] = [
  { cat: 'general', label: '常规' },
  { cat: 'number', label: '数值' },
  { cat: 'currency', label: '货币' },
  { cat: 'accounting', label: '会计专用' },
  { cat: 'date', label: '日期' },
  { cat: 'time', label: '时间' },
  { cat: 'percent', label: '百分比' },
  { cat: 'scientific', label: '科学记数' },
  { cat: 'text', label: '文本' },
  { cat: 'custom', label: '自定义' },
];

/** 日期 / 时间分类里可选的格式 */
export const DATE_CODES = ['yyyy/m/d', 'yyyy"年"m"月"d"日"', 'yyyy-mm-dd', 'yyyy"年"m"月"', 'm"月"d"日"', 'yyyy/m/d h:mm'];
export const TIME_CODES = ['h:mm:ss', 'h:mm', 'h"时"mm"分"', 'h"时"mm"分"ss"秒"', 'h:mm AM/PM'];

/** 带小数位、千位分隔的分类 */
export const FMT_DECIMALS: FmtCategory[] = ['number', 'currency', 'accounting', 'percent', 'scientific'];

export interface FmtParts {
  cat: FmtCategory;
  /** 小数位数 0–30 */
  decimals: number;
  /** 数值：使用千位分隔符 */
  thousands: boolean;
}

/** 格式代码 → 分类、小数位、千位分隔（打开对话框时选中对应的分类） */
export function fmtParts(code: string): FmtParts {
  const plain = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  const first = plain.split(';')[0];
  const decimals = first.match(/\.(0+)/)?.[1].length ?? 0;
  const thousands = /#,##0/.test(first);
  const base = { decimals, thousands };
  if (code === 'General' || code === '') return { cat: 'general', decimals: 2, thousands: false };
  if (code === '@') return { cat: 'text', decimals: 2, thousands: false };
  if (isDateFormat(code)) return { cat: formatLabel(code) === '时间' ? 'time' : 'date', decimals: 2, thousands: false };
  if (/E\+/i.test(first)) return { cat: 'scientific', ...base };
  if (/%/.test(first)) return { cat: 'percent', ...base };
  if (/\*/.test(first) && /[¥$]/.test(code)) return { cat: 'accounting', ...base };
  if (/[¥$]/.test(code)) return { cat: 'currency', ...base };
  // 用原始代码判断：带单位（0.00"吨"）、颜色等的是自定义
  if (/^(#,##0|0)(\.0+)?$/.test(code.split(';')[0].trim())) return { cat: 'number', ...base };
  return { cat: 'custom', decimals: 2, thousands: false };
}

/** 分类 + 小数位 + 千位分隔 → 格式代码；日期 / 时间 / 自定义用 code 原样 */
export function buildFormat(p: FmtParts, code = ''): string {
  const d = Math.max(0, Math.min(30, Math.round(p.decimals)));
  const dec = d ? `.${'0'.repeat(d)}` : '';
  switch (p.cat) {
    case 'general':
      return 'General';
    case 'text':
      return '@';
    case 'number':
      return `${p.thousands ? '#,##0' : '0'}${dec}`;
    case 'currency':
      return `"¥"#,##0${dec};"¥"-#,##0${dec}`;
    case 'accounting':
      return `_ "¥"* #,##0${dec}_ ;_ "¥"* -#,##0${dec}_ ;_ "¥"* "-"${'?'.repeat(d)}_ ;_ @_ `;
    case 'percent':
      return `0${dec}%`;
    case 'scientific':
      return `0${dec}E+00`;
    default:
      return code || 'General';
  }
}

/**
 * 字体 / 段落对话框和工具条用到的纯逻辑：带单位的长度（字符、行、磅、厘米…）、字号（中文字号名、增大 / 减小字号）、
 * 段落对话框的表单与 ParaProps 之间的换算（只把改过的项写回，没动的项一点不变）。
 */
import { CN_SIZES, type Align, type LineSpacing, type ParaProps } from './model';

export type LenUnit = 'char' | 'line' | 'pt' | 'cm' | 'mm' | 'in';

const UNIT_NAMES: [RegExp, LenUnit][] = [
  [/^(字符|字|ch|chars?)$/i, 'char'],
  [/^(行|li|lines?)$/i, 'line'],
  [/^(磅|pt|points?)$/i, 'pt'],
  [/^(厘米|公分|cm)$/i, 'cm'],
  [/^(毫米|mm)$/i, 'mm'],
  [/^(英寸|in|")$/i, 'in'],
];

const UNIT_LABEL: Record<LenUnit, string> = { char: '字符', line: '行', pt: '磅', cm: '厘米', mm: '毫米', in: '英寸' };

/** “2 字符”“0.5行”“12pt”“1.5 厘米”→ 数值和单位；没写单位按 def；读不出来返回 null */
export function parseLength(text: string, def: LenUnit): { value: number; unit: LenUnit } | null {
  const m = text.trim().match(/^(-?\d+(?:\.\d+)?|-?\.\d+)\s*(\S*)$/);
  if (!m) return null;
  const value = Number(m[1]);
  if (!Number.isFinite(value)) return null;
  if (!m[2]) return { value, unit: def };
  const hit = UNIT_NAMES.find(([re]) => re.test(m[2]));
  return hit ? { value, unit: hit[1] } : null;
}

/** 换成磅：字符按 charPt（一个字符多宽），行按 linePt */
export function lengthToPt(v: { value: number; unit: LenUnit }, charPt: number, linePt = 12): number {
  switch (v.unit) {
    case 'char':
      return v.value * charPt;
    case 'line':
      return v.value * linePt;
    case 'cm':
      return (v.value * 72) / 2.54;
    case 'mm':
      return (v.value * 72) / 25.4;
    case 'in':
      return v.value * 72;
    default:
      return v.value;
  }
}

/** 数值 + 单位的显示（最多两位小数，去掉末尾的 0） */
export function formatLength(value: number, unit: LenUnit): string {
  return `${+value.toFixed(2)} ${UNIT_LABEL[unit]}`;
}

// ---------------- 字号 ----------------

/** “小四”“13.5”“13.5 磅”→ 磅；Word 的范围是 1–1638 磅，按 0.5 磅取整 */
export function parseFontSize(text: string): number | null {
  const t = text.trim().replace(/\s*(磅|pt)$/i, '');
  const cn = CN_SIZES.find(([name]) => name === t);
  if (cn) return cn[1];
  if (!/^\d+(\.\d+)?$|^\.\d+$/.test(t)) return null;
  const v = Math.round(Number(t) * 2) / 2;
  return v >= 1 && v <= 1638 ? v : null;
}

/** 增大 / 减小字号（Word 的 Ctrl+Shift+> / <）走的字号表；72 磅以上每次 10 磅 */
const SIZE_STEPS = [...new Set([...CN_SIZES.map(([, v]) => v), 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72])].filter((v) => v <= 72).sort((a, b) => a - b);

export function stepFontSize(size: number, dir: 1 | -1): number {
  if (dir > 0) {
    if (size >= 72) return Math.min(1638, Math.floor(size / 10) * 10 + 10);
    return SIZE_STEPS.find((v) => v > size + 0.01) ?? 80;
  }
  if (size > 80) return Math.ceil(size / 10) * 10 - 10;
  if (size > 72) return 72;
  const below = SIZE_STEPS.filter((v) => v < size - 0.01);
  return below.length ? below[below.length - 1] : Math.max(1, size - 1);
}

/** 每次 1 磅（Word 的 Ctrl+] / Ctrl+[） */
export function nudgeFontSize(size: number, dir: 1 | -1): number {
  return Math.min(1638, Math.max(1, Math.round(size) + dir));
}

// ---------------- 段落对话框 ----------------

export type LineChoice = 'single' | '1.5' | 'double' | 'atLeast' | 'exact' | 'multiple';

export interface ParaForm {
  align: Align;
  /** 大纲级别 0–8；null = 正文文本 */
  outline: number | null;
  left: string;
  right: string;
  special: 'none' | 'first' | 'hanging';
  specialBy: string;
  before: string;
  after: string;
  lineRule: LineChoice;
  /** 多倍行距的倍数，或最小值 / 固定值的磅数 */
  lineAt: string;
  snapToGrid: boolean;
  widowControl: boolean;
  keepNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
}

/** 段落格式 → 表单。charPt = 这一段一个字符多宽（按字符的缩进用它换算） */
export function paraFormOf(p: ParaProps, charPt: number): ParaForm {
  const chars = (pt: number) => formatLength(pt / charPt, 'char');
  const first = p.firstLineChars !== null ? p.firstLineChars * charPt : p.firstLine;
  const space = (lines: number | null, pt: number, auto?: boolean) => (lines !== null ? formatLength(lines, 'line') : auto ? '自动' : formatLength(pt, 'pt'));
  const l = p.line;
  const lineRule: LineChoice =
    l.rule === 'exact' ? 'exact' : l.rule === 'atLeast' ? 'atLeast' : l.value === 1 ? 'single' : l.value === 1.5 ? '1.5' : l.value === 2 ? 'double' : 'multiple';
  return {
    align: p.align,
    outline: p.outline ?? null,
    left: chars(p.indentLeft),
    right: chars(p.indentRight),
    special: first > 0.001 ? 'first' : first < -0.001 ? 'hanging' : 'none',
    specialBy: first ? (p.firstLineChars !== null ? formatLength(Math.abs(p.firstLineChars), 'char') : chars(Math.abs(first))) : formatLength(2, 'char'),
    before: space(p.spaceBeforeLines, p.spaceBefore, p.spaceBeforeAuto),
    after: space(p.spaceAfterLines, p.spaceAfter, p.spaceAfterAuto),
    lineRule,
    lineAt: l.rule === 'auto' ? String(+l.value.toFixed(2)) : formatLength(l.value, 'pt'),
    snapToGrid: p.snapToGrid,
    widowControl: p.widowControl,
    keepNext: p.keepNext,
    keepLines: p.keepLines,
    pageBreakBefore: p.pageBreakBefore,
  };
}

/** 表单里读不出来的项（返回第一个出错的说明），都能读时返回 null */
export function checkParaForm(f: ParaForm): string | null {
  const len = (t: string, def: LenUnit) => parseLength(t, def);
  if (!len(f.left, 'char')) return `左侧缩进“${f.left}”无法识别（例：2 字符、0.5 厘米、12 磅）`;
  if (!len(f.right, 'char')) return `右侧缩进“${f.right}”无法识别`;
  if (f.special !== 'none' && !len(f.specialBy, 'char')) return `缩进值“${f.specialBy}”无法识别`;
  for (const [name, t] of [['段前', f.before], ['段后', f.after]] as const) {
    const v = t.trim() === '自动' ? { value: 0, unit: 'line' as const } : len(t, 'line');
    if (!v || v.value < 0 || (v.unit !== 'line' && v.unit !== 'pt' && v.unit !== 'cm' && v.unit !== 'mm')) return `${name}“${t}”无法识别（例：0.5 行、6 磅、自动）`;
  }
  if (f.lineRule === 'multiple') {
    const v = Number(f.lineAt);
    if (!(v > 0 && v <= 132)) return `多倍行距“${f.lineAt}”应为 0.06–132 之间的数`;
  }
  if (f.lineRule === 'exact' || f.lineRule === 'atLeast') {
    const v = len(f.lineAt, 'pt');
    if (!v || v.value < 0 || (f.lineRule === 'exact' && v.value === 0) || v.unit === 'char' || v.unit === 'line') return `行距设置值“${f.lineAt}”无法识别（例：20 磅）`;
  }
  return null;
}

/**
 * 表单 → 要改的段落格式：只有和打开对话框时（init）不同的项才写进去，所以没动的项保持原值（多段时各段自己的值也不变）。
 * 调用前先用 checkParaForm 检查。
 */
export function paraPatchOf(f: ParaForm, init: ParaForm, charPt: number): Partial<ParaProps> {
  const out: Partial<ParaProps> = {};
  const pt = (t: string, def: LenUnit) => Math.round(lengthToPt(parseLength(t, def)!, charPt) * 100) / 100;
  if (f.align !== init.align) out.align = f.align;
  if (f.outline !== init.outline) out.outline = f.outline ?? undefined;
  if (f.left !== init.left) out.indentLeft = pt(f.left, 'char');
  if (f.right !== init.right) out.indentRight = pt(f.right, 'char');
  if (f.special !== init.special || f.specialBy !== init.specialBy) {
    if (f.special === 'none') Object.assign(out, { firstLine: 0, firstLineChars: null });
    else {
      const v = parseLength(f.specialBy, 'char')!;
      const sign = f.special === 'hanging' ? -1 : 1;
      // 按字符设的记字符数（排版时按当时的字宽换算，同 Word 的 firstLineChars / hangingChars），磅值同时写上
      out.firstLine = sign * Math.round(lengthToPt(v, charPt) * 100) / 100;
      out.firstLineChars = v.unit === 'char' ? sign * Math.abs(v.value) : null;
    }
  }
  const space = (t: string, side: 'Before' | 'After') => {
    if (t.trim() === '自动') return Object.assign(out, { [`space${side}Auto`]: true, [`space${side}Lines`]: null });
    const v = parseLength(t, 'line')!;
    if (v.unit === 'line') Object.assign(out, { [`space${side}Lines`]: v.value, [`space${side}`]: 0, [`space${side}Auto`]: false });
    else Object.assign(out, { [`space${side}`]: Math.round(lengthToPt(v, charPt) * 100) / 100, [`space${side}Lines`]: null, [`space${side}Auto`]: false });
  };
  if (f.before !== init.before) space(f.before, 'Before');
  if (f.after !== init.after) space(f.after, 'After');
  if (f.lineRule !== init.lineRule || f.lineAt !== init.lineAt) {
    const line: LineSpacing =
      f.lineRule === 'single'
        ? { rule: 'auto', value: 1 }
        : f.lineRule === '1.5'
          ? { rule: 'auto', value: 1.5 }
          : f.lineRule === 'double'
            ? { rule: 'auto', value: 2 }
            : f.lineRule === 'multiple'
              ? { rule: 'auto', value: Math.round(Number(f.lineAt) * 100) / 100 }
              : { rule: f.lineRule, value: pt(f.lineAt, 'pt') };
    out.line = line;
  }
  for (const k of ['snapToGrid', 'widowControl', 'keepNext', 'keepLines', 'pageBreakBefore'] as const) if (f[k] !== init[k]) out[k] = f[k];
  return out;
}

/** 行距下拉换选项时，设置值换成该行距常用的值（同 Word） */
export function lineAtFor(rule: LineChoice, cur: ParaForm): string {
  if (rule === 'multiple') return cur.lineRule === 'multiple' ? cur.lineAt : '3';
  if (rule === 'exact' || rule === 'atLeast') return cur.lineRule === 'exact' || cur.lineRule === 'atLeast' ? cur.lineAt : formatLength(12, 'pt');
  return rule === 'single' ? '1' : rule === '1.5' ? '1.5' : '2';
}

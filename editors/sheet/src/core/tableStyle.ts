/**
 * Excel 表格（插入 → 表格，xl/tables）的内置样式显示：标题行底色、隔行底纹、汇总行加粗、首 / 末列加粗。
 * 只认内置样式名（TableStyleLight1–21 / Medium1–28 / Dark1–11），颜色取自主题的强调色；自定义表格样式暂不显示。
 * 单元格自己的格式（底色、字体颜色、加粗）优先于表格样式（同 Excel）；边框、列隔行暂不画。
 * 样式是按 Excel 内置样式的配色规则近似出来的（标题行 / 隔行的色调与深浅），不是逐像素一致。
 */
import { applyTint, type CellBorder } from './color';
import type { CellStyle } from './styles';
import type { Range } from './workbook';

export interface SheetTable {
  /** 表格部件路径 xl/tables/tableN.xml */
  path: string;
  ref: Range;
  headerRows: number;
  totalsRows: number;
  style?: string;
  rowStripes: boolean;
  firstCol: boolean;
  lastCol: boolean;
  /** 标题行有筛选下拉按钮（表格自带 autoFilter） */
  filter?: boolean;
}

interface Spec {
  /** 强调色序号 0–5；null = 黑 / 灰 */
  accent: number | null;
  /** 标题行底色：深色（强调色）/ 浅色（强调色的淡色）/ 无（只加粗） */
  header: 'solid' | 'tint' | 'plain';
  /** 隔行底纹 */
  stripe: 'tint' | 'dark' | 'none';
  /** 深色样式：整个表格底色 + 白字 */
  dark: boolean;
  /** 内置样式的细线：Medium 每行下面一条强调色淡线；Light 只在标题行下面 */
  lines: 'rows' | 'header' | 'none';
}

/** 样式名 → 配色规则；认不出的返回 null */
export function specOf(name: string | undefined): Spec | null {
  const m = name?.match(/^TableStyle(Light|Medium|Dark)(\d+)$/);
  if (!m) return null;
  const n = Number(m[2]);
  const group = (n - 1) % 7; // 每 7 个一组：黑 / 灰，再是 6 个强调色
  const accent = group === 0 ? null : group - 1;
  if (m[1] === 'Dark') return { accent: n === 1 ? null : (n - 2) % 6, header: 'solid', stripe: 'dark', dark: true, lines: 'none' };
  if (m[1] === 'Light') {
    const band = Math.floor((n - 1) / 7); // 0: 标题加粗 + 隔行；1: 标题深色底；2: 网格
    return { accent, header: band === 1 ? 'solid' : 'plain', stripe: band === 1 ? 'none' : 'tint', dark: false, lines: band === 0 ? 'header' : 'none' };
  }
  const band = Math.floor((n - 1) / 7); // 0–1: 标题深色；2: 标题淡色；3: 标题淡色
  return { accent, header: band >= 2 ? 'tint' : 'solid', stripe: 'tint', dark: false, lines: 'rows' };
}

/** 主题色：theme 里 0 lt1、1 dk1、2 lt2、3 dk2、4–9 强调色 1–6 */
function base(theme: string[], accent: number | null): string {
  return accent === null ? (theme[1] ?? '000000') : (theme[4 + accent] ?? '4472C4');
}

/** 某个格子上表格样式给的格式（不在任何表格里 / 样式不认识 / 没有格式时 undefined） */
export function tableStyleAt(tables: SheetTable[] | undefined, theme: string[], r: number, c: number): CellStyle | undefined {
  if (!tables) return undefined;
  for (const t of tables) {
    const g = t.ref;
    if (r < g.r0 || r > g.r1 || c < g.c0 || c > g.c1) continue;
    const spec = specOf(t.style);
    if (!spec) return undefined;
    const accent = base(theme, spec.accent);
    const out: CellStyle = {};
    const line = spec.lines === 'rows' ? { b: { w: 1, style: 'solid' as const, color: `#${spec.accent === null ? 'A6A6A6' : applyTint(accent, 0.4)}` } } : undefined;
    const isHeader = r < g.r0 + t.headerRows;
    const isTotal = t.totalsRows > 0 && r > g.r1 - t.totalsRows;
    if (isHeader) {
      out.bold = true;
      if (line) out.border = line;
      else if (spec.lines === 'header') out.border = { b: { w: 1, style: 'solid', color: `#${accent}` } };
      if (spec.header === 'solid') {
        out.fill = `#${spec.accent === null && !spec.dark ? base(theme, null) : accent}`;
        out.color = '#FFFFFF';
      } else if (spec.header === 'tint') out.fill = `#${applyTint(accent, 0.6)}`;
      return out;
    }
    if (isTotal) {
      out.bold = true;
      if (line) out.border = line;
      if (spec.dark) {
        out.fill = `#${spec.accent === null ? '000000' : applyTint(accent, -0.4)}`;
        out.color = '#FFFFFF';
      }
      return out;
    }
    const k = r - (g.r0 + t.headerRows);
    if (line) out.border = line;
    if (spec.dark) {
      out.fill = `#${applyTint(accent, t.rowStripes && k % 2 === 1 ? -0.5 : -0.25)}`;
      out.color = '#FFFFFF';
    } else if (spec.stripe === 'tint' && t.rowStripes && k % 2 === 0) {
      out.fill = `#${spec.accent === null ? 'D9D9D9' : applyTint(accent, 0.8)}`;
    }
    if ((t.firstCol && c === g.c0) || (t.lastCol && c === g.c1)) out.bold = true;
    return Object.keys(out).length ? out : undefined;
  }
  return undefined;
}

/**
 * 表格样式在下，单元格自己设的格式在上。默认的黑色字体颜色（文件里常写成 theme=1 的黑）不算“自己设的”，
 * 否则标题行的白字会被盖成黑字（Excel 里标题行显示白字）。
 */
export function overTable(tbl: CellStyle, own: CellStyle | undefined): CellStyle {
  const out: CellStyle = { ...tbl };
  for (const [k, v] of Object.entries(own ?? {})) {
    if (v === undefined) continue;
    if (k === 'color' && typeof v === 'string' && v.toLowerCase() === '#000000') continue;
    if (k === 'border') {
      out.border = { ...tbl.border, ...(v as CellBorder) };
      continue;
    }
    (out as Record<string, unknown>)[k] = v;
  }
  return out;
}

/**
 * SpreadsheetML 颜色（`<color>` / `<fgColor>` / `<bgColor>`）→ CSS `#rrggbb`。
 * 规则（ECMA-376 第 1 部分 18.8.3 / 18.8.19 / 18.8.20）：
 * - `rgb="AARRGGBB"`：直接取 RRGGBB（Excel 显示时忽略透明度）；
 * - `indexed="n"`：旧调色板，`<colors><indexedColors>` 可覆盖；64 = 系统前景色（黑），65 = 系统背景色（白）；
 * - `theme="n"`：主题色，Excel 的序号把深浅两对对调：0 = lt1，1 = dk1，2 = lt2，3 = dk2，4–9 = accent1–6，10 = hlink，11 = folHlink；
 * - `tint`：在 HLS 空间调亮度，tint < 0 时 L × (1 + tint)，tint > 0 时 L × (1 − tint) + tint；
 * - `auto="1"`：自动色，返回 undefined（由使用方决定缺省色）。
 */
import { xmlAttr } from '@wz/editor-shared/xlsx';

/** 默认 64 色调色板（indexed 0–63） */
export const DEFAULT_INDEXED: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

/** Office 默认主题（文档里没有 theme1.xml 时用），按 Excel 的 theme 序号排列 */
export const DEFAULT_THEME: readonly string[] = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72',
];

export interface Palette {
  /** indexed 0–63（已应用 indexedColors 覆盖） */
  indexed: string[];
  /** 按 Excel theme 序号排列的 RRGGBB */
  theme: string[];
}

/** 读 styles.xml 的 `<indexedColors>` 和 theme1.xml 的 `<a:clrScheme>` */
export function readPalette(stylesXml: string, themeXml: string): Palette {
  const indexed = [...DEFAULT_INDEXED];
  const custom = stylesXml.match(/<indexedColors\b[^>]*>([\s\S]*?)<\/indexedColors>/)?.[1];
  if (custom) {
    let i = 0;
    for (const m of custom.matchAll(/<rgbColor\b[^>]*>/g)) {
      const rgb = hex6(xmlAttr(m[0], 'rgb'));
      if (rgb && i < 64) indexed[i] = rgb;
      i++;
    }
  }
  const theme = [...DEFAULT_THEME];
  const scheme = themeXml.match(/<(?:\w+:)?clrScheme\b[^>]*>([\s\S]*?)<\/(?:\w+:)?clrScheme>/)?.[1];
  if (scheme) {
    const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    order.forEach((name, i) => {
      const body = scheme.match(new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${name}>`))?.[1] ?? '';
      const srgb = body.match(/<(?:\w+:)?srgbClr\b[^>]*>/)?.[0];
      const sys = body.match(/<(?:\w+:)?sysClr\b[^>]*>/)?.[0];
      const rgb = hex6(srgb ? xmlAttr(srgb, 'val') : sys ? xmlAttr(sys, 'lastClr') : undefined);
      if (rgb) theme[i] = rgb;
    });
  }
  return { indexed, theme };
}

/** 一个颜色元素（`<color .../>` 等的开始标签）→ `#rrggbb`；自动色或无法识别时 undefined */
export function resolveColor(tag: string | undefined, pal: Palette): string | undefined {
  if (!tag) return undefined;
  if (xmlAttr(tag, 'auto') === '1' || xmlAttr(tag, 'auto') === 'true') return undefined;
  let rgb: string | undefined;
  const raw = xmlAttr(tag, 'rgb');
  const idx = xmlAttr(tag, 'indexed');
  const th = xmlAttr(tag, 'theme');
  if (raw !== undefined) rgb = hex6(raw);
  else if (th !== undefined) rgb = pal.theme[Number(th)];
  else if (idx !== undefined) {
    const n = Number(idx);
    rgb = n === 64 ? '000000' : n === 65 ? 'FFFFFF' : pal.indexed[n];
  }
  if (!rgb) return undefined;
  const tint = Number(xmlAttr(tag, 'tint') ?? 0);
  return `#${(tint ? applyTint(rgb, tint) : rgb).toLowerCase()}`;
}

/** AARRGGBB / RRGGBB → RRGGBB */
function hex6(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const s = v.trim();
  if (/^[0-9a-f]{8}$/i.test(s)) return s.slice(2).toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(s)) return s.toUpperCase();
  return undefined;
}

/** 按 tint 调亮 / 调暗（HLS 亮度） */
export function applyTint(rgb: string, tint: number): string {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  const d = max - min;
  if (d) {
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  l = Math.max(0, Math.min(1, l));
  const out = s === 0 ? [l, l, l] : hlsToRgb(h, l, s);
  return out.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function hlsToRgb(h: number, l: number, s: number): number[] {
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)];
}

/** 一条单元格边框 */
export interface BorderSide {
  /** 线宽（px） */
  w: number;
  style: 'solid' | 'dashed' | 'dotted' | 'double';
  /** CSS #rrggbb；自动色（或没写颜色）时没有，按黑色画 */
  color?: string;
}

/** 单元格四边的边框（没有的边不画） */
export interface CellBorder {
  l?: BorderSide;
  r?: BorderSide;
  t?: BorderSide;
  b?: BorderSide;
}

/** ST_BorderStyle → 线宽与线型 */
const BORDER_STYLES: Record<string, Omit<BorderSide, 'color'>> = {
  thin: { w: 1, style: 'solid' },
  hair: { w: 1, style: 'dotted' },
  dotted: { w: 1, style: 'dotted' },
  dashed: { w: 1, style: 'dashed' },
  dashDot: { w: 1, style: 'dashed' },
  dashDotDot: { w: 1, style: 'dashed' },
  medium: { w: 2, style: 'solid' },
  mediumDashed: { w: 2, style: 'dashed' },
  mediumDashDot: { w: 2, style: 'dashed' },
  mediumDashDotDot: { w: 2, style: 'dashed' },
  slantDashDot: { w: 2, style: 'dashed' },
  thick: { w: 3, style: 'solid' },
  double: { w: 3, style: 'double' },
};

/**
 * `<border>` 的内容 → 四边边框（ECMA-376 18.8.4）。`start` / `end` 是新版写法，等同 `left` / `right`；
 * 没写 style 或 style="none" 的边没有边框。对角线（diagonal）不支持。
 */
export function resolveBorder(borderXml: string, pal: Palette): CellBorder | undefined {
  const side = (...names: string[]): BorderSide | undefined => {
    for (const name of names) {
      const m = borderXml.match(new RegExp(`<${name}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${name}>)`));
      if (!m) continue;
      const kind = BORDER_STYLES[xmlAttr(m[1], 'style') ?? 'none'];
      if (!kind) return undefined;
      const color = resolveColor((m[2] ?? '').match(/<color\b[^>]*>/)?.[0], pal);
      return color ? { ...kind, color } : { ...kind };
    }
    return undefined;
  };
  const out: CellBorder = {};
  const l = side('left', 'start');
  const r = side('right', 'end');
  const t = side('top');
  const b = side('bottom');
  if (l) out.l = l;
  if (r) out.r = r;
  if (t) out.t = t;
  if (b) out.b = b;
  return l || r || t || b ? out : undefined;
}

/** 图案填充里前景色所占比例：只画一种平均色（不画图案本身） */
const PATTERN_DENSITY: Record<string, number> = {
  darkGray: 0.75,
  mediumGray: 0.5,
  lightGray: 0.25,
  gray125: 0.125,
  gray0625: 0.0625,
};

/**
 * `<fill>` → 单元格底色。`solid` 用前景色 fgColor；其它图案按前景色占比把 fgColor 混到 bgColor 上（近似）；
 * 渐变取各色标的平均色；`none` / 缺省 patternType 无底色。
 */
export function resolveFill(fillXml: string, pal: Palette): string | undefined {
  const grad = fillXml.match(/<gradientFill\b[\s\S]*?<\/gradientFill>/)?.[0];
  if (grad) {
    const stops = [...grad.matchAll(/<color\b[^>]*>/g)].map((m) => resolveColor(m[0], pal)).filter((c): c is string => !!c);
    return stops.length ? average(stops) : undefined;
  }
  const pf = fillXml.match(/<patternFill\b([^>]*?)(?:\/>|>([\s\S]*?)<\/patternFill>)/);
  if (!pf) return undefined;
  const type = xmlAttr(pf[1], 'patternType') ?? 'none';
  if (type === 'none') return undefined;
  const inner = pf[2] ?? '';
  // 缺省前景色是系统前景色（黑），缺省背景色是系统背景色（白）
  const fg = resolveColor(inner.match(/<fgColor\b[^>]*>/)?.[0], pal) ?? '#000000';
  if (type === 'solid') return fg;
  const bg = resolveColor(inner.match(/<bgColor\b[^>]*>/)?.[0], pal) ?? '#ffffff';
  const k = PATTERN_DENSITY[type] ?? (type.startsWith('dark') ? 0.5 : 0.25);
  return mix(fg, bg, k);
}

function mix(a: string, b: string, k: number): string {
  const ca = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16));
  const cb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16));
  return `#${ca.map((v, i) => Math.round(v * k + cb[i] * (1 - k)).toString(16).padStart(2, '0')).join('')}`;
}

function average(colors: string[]): string {
  const sum = [0, 0, 0];
  for (const c of colors) [1, 3, 5].forEach((i, j) => (sum[j] += parseInt(c.slice(i, i + 2), 16)));
  return `#${sum.map((v) => Math.round(v / colors.length).toString(16).padStart(2, '0')).join('')}`;
}

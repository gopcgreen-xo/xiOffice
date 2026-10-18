/**
 * 形状（DrawingML 预设形状）的外形和颜色计算，纯函数（排版、绘制、读写共用）。
 *
 * 外形按 DrawingML presetShapeDefinitions 的公式和缺省调整值（avLst 为空时的值）算，
 * 坐标是框内的磅值（左上角 0, 0）。读不到调整值时一律用缺省值（多数文档用的就是缺省值）。
 * 不在 SHAPE_GEOMS 里的预设形状、自由曲线（custGeom）不当作形状读入（docx 原样保留，.doc 记进“不支持的内容”）。
 */

export type PathCmd = ['M', number, number] | ['L', number, number] | ['C', number, number, number, number, number, number] | ['Z'];

/** 支持显示的预设形状 */
export const SHAPE_GEOMS = new Set([
  'rect',
  'roundRect',
  'ellipse',
  'diamond',
  'triangle',
  'rtTriangle',
  'parallelogram',
  'trapezoid',
  'hexagon',
  'octagon',
  'pentagon',
  'homePlate',
  'chevron',
  'plus',
  'star5',
  'rightArrow',
  'leftArrow',
  'upArrow',
  'downArrow',
  'line',
  'straightConnector1',
  'flowChartProcess',
  'flowChartAlternateProcess',
  'flowChartDecision',
]);

/** 只有线、没有面积的形状：不填充，文字不显示 */
export const isLineGeom = (g: string | undefined): boolean => g === 'line' || g === 'straightConnector1';

/** 四分之一圆弧的贝塞尔控制点系数 */
const K = 0.5522847498;

function poly(pts: [number, number][]): PathCmd[] {
  const out: PathCmd[] = pts.map(([x, y], i) => [i ? 'L' : 'M', x, y] as PathCmd);
  out.push(['Z']);
  return out;
}

function ellipse(x: number, y: number, w: number, h: number): PathCmd[] {
  const rx = w / 2;
  const ry = h / 2;
  const cx = x + rx;
  const cy = y + ry;
  return [
    ['M', cx + rx, cy],
    ['C', cx + rx, cy + ry * K, cx + rx * K, cy + ry, cx, cy + ry],
    ['C', cx - rx * K, cy + ry, cx - rx, cy + ry * K, cx - rx, cy],
    ['C', cx - rx, cy - ry * K, cx - rx * K, cy - ry, cx, cy - ry],
    ['C', cx + rx * K, cy - ry, cx + rx, cy - ry * K, cx + rx, cy],
    ['Z'],
  ];
}

function roundRect(w: number, h: number, r: number): PathCmd[] {
  const k = r * (1 - K);
  return [
    ['M', r, 0],
    ['L', w - r, 0],
    ['C', w - k, 0, w, k, w, r],
    ['L', w, h - r],
    ['C', w, h - k, w - k, h, w - r, h],
    ['L', r, h],
    ['C', k, h, 0, h - k, 0, h - r],
    ['L', 0, r],
    ['C', 0, k, k, 0, r, 0],
    ['Z'],
  ];
}

/** 五角星：外接椭圆贴着框（DrawingML star5 的 hf / vf 让星形撑满框，这里按撑满处理），内径比 0.382 */
function star5(w: number, h: number): PathCmd[] {
  const pts: [number, number][] = [];
  // 外顶点所在的椭圆：上顶点贴上边，下面两个顶点贴下边
  const ry = h / (1 + Math.cos(Math.PI / 5));
  const rx = w / 2 / Math.sin((2 * Math.PI) / 5);
  const cx = w / 2;
  const cy = ry;
  for (let i = 0; i < 10; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / 5;
    const f = i % 2 ? 0.382 : 1;
    pts.push([cx + rx * f * Math.cos(a), cy + ry * f * Math.sin(a)]);
  }
  return poly(pts);
}

/** 形状的轮廓（框内坐标，磅）；不认识的形状按矩形 */
export function shapePath(geom: string | undefined, w: number, h: number): PathCmd[] {
  const ss = Math.min(w, h);
  switch (geom) {
    case 'roundRect':
    case 'flowChartAlternateProcess':
      return roundRect(w, h, ss * 0.16667);
    case 'ellipse':
      return ellipse(0, 0, w, h);
    case 'diamond':
    case 'flowChartDecision':
      return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'triangle':
      return poly([[w / 2, 0], [w, h], [0, h]]);
    case 'rtTriangle':
      return poly([[0, 0], [w, h], [0, h]]);
    case 'parallelogram': {
      const a = ss * 0.25;
      return poly([[a, 0], [w, 0], [w - a, h], [0, h]]);
    }
    case 'trapezoid': {
      const a = ss * 0.25;
      return poly([[0, h], [a, 0], [w - a, 0], [w, h]]);
    }
    case 'hexagon': {
      const a = ss * 0.25;
      return poly([[0, h / 2], [a, 0], [w - a, 0], [w, h / 2], [w - a, h], [a, h]]);
    }
    case 'octagon': {
      const a = ss * 0.29289;
      return poly([[a, 0], [w - a, 0], [w, a], [w, h - a], [w - a, h], [a, h], [0, h - a], [0, a]]);
    }
    case 'pentagon':
      return poly([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]);
    case 'homePlate': {
      const x1 = w - ss * 0.5;
      return poly([[0, 0], [x1, 0], [w, h / 2], [x1, h], [0, h]]);
    }
    case 'chevron': {
      const a = ss * 0.5;
      return poly([[0, 0], [w - a, 0], [w, h / 2], [w - a, h], [0, h], [a, h / 2]]);
    }
    case 'plus': {
      const a = ss * 0.25;
      return poly([[0, a], [a, a], [a, 0], [w - a, 0], [w - a, a], [w, a], [w, h - a], [w - a, h - a], [w - a, h], [a, h], [a, h - a], [0, h - a]]);
    }
    case 'star5':
      return star5(w, h);
    case 'rightArrow':
    case 'leftArrow': {
      const y1 = h / 4;
      const y2 = (h * 3) / 4;
      const x1 = w - ss * 0.5;
      const pts: [number, number][] = [[0, y1], [x1, y1], [x1, 0], [w, h / 2], [x1, h], [x1, y2], [0, y2]];
      return poly(geom === 'rightArrow' ? pts : pts.map(([x, y]) => [w - x, y]));
    }
    case 'downArrow':
    case 'upArrow': {
      const x1 = w / 4;
      const x2 = (w * 3) / 4;
      const y1 = h - ss * 0.5;
      const pts: [number, number][] = [[x1, 0], [x2, 0], [x2, y1], [w, y1], [w / 2, h], [0, y1], [x1, y1]];
      return poly(geom === 'downArrow' ? pts : pts.map(([x, y]) => [x, h - y]));
    }
    case 'line':
    case 'straightConnector1':
      return [['M', 0, 0], ['L', w, h]];
    default:
      return poly([[0, 0], [w, 0], [w, h], [0, h]]);
  }
}

// ---------------- 颜色 ----------------

const toLin = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c: number) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
const clamp01 = (v: number) => Math.min(1, Math.max(0, v));
const hex = (rgb: number[]) => rgb.map((v) => Math.round(clamp01(v) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();

function rgbToHsl([r, g, b]: number[]): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb([h, s, l]: number[]): number[] {
  if (!s) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    const u = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    return u < 1 / 6 ? p + (q - p) * 6 * u : u < 1 / 2 ? q : u < 2 / 3 ? p + (q - p) * (2 / 3 - u) * 6 : p;
  };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)];
}

/**
 * DrawingML 颜色变换（a:shade / a:tint / a:lumMod / a:lumOff，val 为 1/1000 百分比，如 50000 = 50%）。
 * shade / tint 在线性 RGB 里算（Word 实测：accent1 4472C4 shade 50% = 2F528F），lumMod / lumOff 在 HSL 里算。
 * 其它变换（alpha、satMod…）忽略。
 */
export function applyColorMods(color: string, mods: { name: string; val: number }[]): string {
  let rgb = [0, 2, 4].map((i) => parseInt(color.slice(i, i + 2), 16) / 255);
  for (const { name, val } of mods) {
    const v = val / 100000;
    if (name === 'shade') rgb = rgb.map((c) => toSrgb(toLin(c) * v));
    else if (name === 'tint') rgb = rgb.map((c) => toSrgb(toLin(c) * v + (1 - v)));
    else if (name === 'lumMod' || name === 'lumOff') {
      const hsl = rgbToHsl(rgb);
      hsl[2] = clamp01(name === 'lumMod' ? hsl[2] * v : hsl[2] + v);
      rgb = hslToRgb(hsl);
    }
  }
  return hex(rgb);
}

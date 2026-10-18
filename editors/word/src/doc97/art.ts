/**
 * .doc 里的绘图对象（文本框、形状、浮动图片）。规范：[MS-DOC]（PlcfSpa、FSPA、OfficeArtContent、PlcftxbxTxt）、[MS-ODRAW]。
 *
 * - 正文里的 0x08 字符（带 fSpec）是锚点；表流的 PlcSpaMom（页眉页脚用 PlcSpaHdr）按 cp 给出 FSPA：
 *   形状 id（spid）、相对 bx / by 参照物的外框（twips）、环绕方式 wr、衬于文字下方 fBelowText。
 * - 表流 fcDggInfo 处是 OfficeArtContent：OfficeArtDggContainer（含图片仓库 BStore），然后每个绘图一个字节 dgglbl（0 正文、1 页眉页脚）
 *   + OfficeArtDgContainer。绘图里的 OfficeArtSpContainer 给形状类型（FSP 的 recInstance = MSOSPT）、翻转、属性表（OPT：填充、线条、
 *   文字边距、图片序号 pib、旋转、文本框 lTxid…）；组合是嵌套的 OfficeArtSpgrContainer，子形状的位置是组合坐标系（FSPGR）里的 ChildAnchor。
 * - 图片仓库的 BSE 在 .doc 里不带图片数据，数据在 WordDocument 流的 foDelay 处（OfficeArtBlip 记录）。
 * - 文本框文字在文本框故事里（PlcftxbxTxt 分段），FTXBXS.lid = 形状 id。
 *
 * 属性缺省值（[MS-ODRAW] 2.3）：填充开、白色；线条开、黑色、9525 EMU（0.75 磅）；文字边距左右 91440、上下 45720 EMU；
 * 环绕距离左右 114300 EMU、上下 0。
 *
 * 组合按子形状拆开（位置按组合坐标换算到页面），保存为 .docx 后是各自独立的形状。
 */
import type { FloatPos, Paragraph, TextBox } from '../core/model';
import { isLineGeom } from '../core/shapes';
import { i32, u16, u32 } from './sprm';

export interface ArtShape {
  spid: number;
  /** MSOSPT 形状类型（FSP 的 recInstance） */
  spt: number;
  flipH: boolean;
  flipV: boolean;
  /** 属性表（主 OPT 和第三 OPT 合并）：pid → 值 */
  opt: Map<number, number>;
  /** 子形状在组合坐标系里的位置 [左, 上, 右, 下] */
  child: [number, number, number, number] | null;
  /** 组合：自己的坐标系和子形状 */
  group?: { rect: [number, number, number, number]; kids: ArtShape[] };
}

export interface OfficeArt {
  /** 图片仓库（pib 从 1 数）：WordDocument 流里的位置和长度，或 BSE 自带的数据 */
  blips: ({ foDelay: number; size: number; inline: Uint8Array | null } | null)[];
  /** 正文（dgglbl 0）和页眉页脚（dgglbl 1）绘图里的顶层形状，按 spid */
  main: Map<number, ArtShape>;
  header: Map<number, ArtShape>;
}

function readOpt(d: Uint8Array, body: number, count: number, into: Map<number, number>) {
  for (let k = 0; k < count; k++) {
    const pid = u16(d, body + 6 * k) & 0x3fff;
    const v = u32(d, body + 6 * k + 2);
    const old = into.get(pid);
    // 布尔属性组（pid 低 6 位全 1，如 0x01BF 填充、0x01FF 线条）：主属性表和第三属性表各写一部分位，
    // 只取“有效位”（第 16 + i 位）置了的那些位，不能整个覆盖
    if ((pid & 0x3f) === 0x3f && old !== undefined) {
      const use = v >>> 16;
      into.set(pid, ((((old >>> 16) | use) << 16) | ((old & 0xffff & ~use) | (v & use))) >>> 0);
    } else into.set(pid, v);
  }
}

/** 一个 OfficeArtSpContainer 的内容 */
function readSp(d: Uint8Array, from: number, to: number): (ArtShape & { grpRect?: [number, number, number, number]; isGroup: boolean }) | null {
  let p = from;
  let sp: (ArtShape & { grpRect?: [number, number, number, number]; isGroup: boolean }) | null = null;
  const opt = new Map<number, number>();
  let child: ArtShape['child'] = null;
  let grpRect: [number, number, number, number] | undefined;
  while (p + 8 <= to) {
    const vi = u16(d, p);
    const type = u16(d, p + 2);
    const len = u32(d, p + 4);
    const body = p + 8;
    if (type === 0xf00a) {
      const flags = u32(d, body + 4);
      sp = { spid: u32(d, body), spt: vi >> 4, flipH: !!(flags & 0x40), flipV: !!(flags & 0x80), opt, child: null, isGroup: !!(flags & 1) };
    } else if (type === 0xf00b || type === 0xf122) readOpt(d, body, vi >> 4, opt);
    else if (type === 0xf00f && len >= 16) child = [i32(d, body), i32(d, body + 4), i32(d, body + 8), i32(d, body + 12)];
    else if (type === 0xf009 && len >= 16) grpRect = [i32(d, body), i32(d, body + 4), i32(d, body + 8), i32(d, body + 12)];
    p = body + len;
  }
  if (!sp) return null;
  sp.child = child;
  if (grpRect) sp.grpRect = grpRect;
  return sp;
}

/** OfficeArtSpgrContainer：第一个 SpContainer 是组合自己，后面是成员 */
function readGroup(d: Uint8Array, from: number, to: number): ArtShape | null {
  let p = from;
  let self: ReturnType<typeof readSp> = null;
  const kids: ArtShape[] = [];
  while (p + 8 <= to) {
    const type = u16(d, p + 2);
    const len = u32(d, p + 4);
    const body = p + 8;
    const end = Math.min(to, body + len);
    if (type === 0xf004) {
      const sp = readSp(d, body, end);
      if (sp && !self) self = sp;
      else if (sp) kids.push(sp);
    } else if (type === 0xf003) {
      const g = readGroup(d, body, end);
      if (g) kids.push(g);
    }
    p = end;
  }
  if (!self) return null;
  return { spid: self.spid, spt: self.spt, flipH: self.flipH, flipV: self.flipV, opt: self.opt, child: self.child, group: { rect: self.grpRect ?? [0, 0, 1, 1], kids } };
}

/** 一个绘图（OfficeArtDgContainer）里的顶层形状（跳过最外层组合自己的“画布”形状） */
function readDrawing(d: Uint8Array, from: number, to: number, out: Map<number, ArtShape>) {
  let p = from;
  while (p + 8 <= to) {
    const type = u16(d, p + 2);
    const len = u32(d, p + 4);
    const body = p + 8;
    const end = Math.min(to, body + len);
    if (type === 0xf003) {
      // 最外层组合（patriarch）：成员就是顶层形状
      const top = readGroup(d, body, end);
      for (const k of top?.group?.kids ?? []) out.set(k.spid, k);
    } else if (type === 0xf004) {
      const sp = readSp(d, body, end);
      if (sp) out.set(sp.spid, sp);
    }
    p = end;
  }
}

export function parseOfficeArt(art: Uint8Array): OfficeArt {
  const res: OfficeArt = { blips: [], main: new Map(), header: new Map() };
  if (art.length < 8) return res;
  // OfficeArtDggContainer
  const dggLen = u32(art, 4);
  const dggEnd = Math.min(art.length, 8 + dggLen);
  let p = 8;
  while (p + 8 <= dggEnd) {
    const type = u16(art, p + 2);
    const len = u32(art, p + 4);
    const body = p + 8;
    const end = Math.min(dggEnd, body + len);
    if (type === 0xf001) {
      let q = body;
      while (q + 8 <= end) {
        const t = u16(art, q + 2);
        const l = u32(art, q + 4);
        const b = q + 8;
        if (t === 0xf007 && l >= 36) {
          const cbName = art[b + 33] ?? 0;
          const inline = l > 36 + cbName ? art.subarray(b + 36 + cbName, Math.min(end, b + l)) : null;
          res.blips.push({ foDelay: u32(art, b + 28), size: u32(art, b + 20), inline });
        } else res.blips.push(null);
        q = b + l;
      }
    }
    p = end;
  }
  // 各个绘图：dgglbl + OfficeArtDgContainer
  p = dggEnd;
  while (p + 9 <= art.length) {
    const lbl = art[p]!;
    const len = u32(art, p + 5);
    const body = p + 9;
    const end = Math.min(art.length, body + len);
    readDrawing(art, body, end, lbl === 0 ? res.main : res.header);
    p = end;
  }
  return res;
}

/** FSPA：正文里一个绘图锚点 */
export interface Fspa {
  spid: number;
  /** 外框（twips，相对 bx / by 的参照物） */
  l: number;
  t: number;
  r: number;
  b: number;
  /** 0 页边距、1 页面、2 栏（水平）/ 段落（垂直） */
  bx: number;
  by: number;
  /** 环绕：0 / 2 四周型、1 上下型、3 无（浮于文字上方 / 衬于文字下方）、4 紧密型、5 穿越型 */
  wr: number;
  below: boolean;
}

/** PlcSpaMom / PlcSpaHdr：cp → FSPA */
export function parsePlcSpa(t: Uint8Array, fc: number, lcb: number): Map<number, Fspa> {
  const out = new Map<number, Fspa>();
  if (lcb < 4 + 30) return out;
  const n = (lcb - 4) / 30;
  for (let i = 0; i < n; i++) {
    const o = fc + 4 * (n + 1) + 26 * i;
    if (o + 26 > t.length) break;
    const f = u16(t, o + 20);
    out.set(u32(t, fc + 4 * i), {
      spid: u32(t, o),
      l: i32(t, o + 4),
      t: i32(t, o + 8),
      r: i32(t, o + 12),
      b: i32(t, o + 16),
      bx: (f >> 1) & 3,
      by: (f >> 3) & 3,
      wr: (f >> 5) & 15,
      below: !!(f & 0x4000),
    });
  }
  return out;
}

/** PlcftxbxTxt / PlcfHdrtxbxTxt：形状 id → 文本框故事在本故事里的 [起, 止) */
export function parseTxbxPlc(t: Uint8Array, fc: number, lcb: number): Map<number, [number, number]> {
  const out = new Map<number, [number, number]>();
  if (lcb < 4 + 26) return out;
  const n = (lcb - 4) / 26;
  for (let i = 0; i < n; i++) {
    const o = fc + 4 * (n + 1) + 22 * i;
    if (o + 22 > t.length) break;
    const lid = i32(t, o + 14);
    if (lid > 0) out.set(lid, [u32(t, fc + 4 * i), u32(t, fc + 4 * (i + 1))]);
  }
  return out;
}

/** MSOSPT → DrawingML 预设形状名（core/shapes.ts 的 SHAPE_GEOMS）；202 = 文本框（矩形，无 geom） */
const SPT_GEOM: Record<number, string> = {
  1: 'rect',
  2: 'roundRect',
  3: 'ellipse',
  4: 'diamond',
  5: 'triangle',
  6: 'rtTriangle',
  7: 'parallelogram',
  8: 'trapezoid',
  9: 'hexagon',
  10: 'octagon',
  11: 'plus',
  12: 'star5',
  13: 'rightArrow',
  15: 'homePlate',
  20: 'line',
  32: 'straightConnector1',
  55: 'chevron',
  56: 'pentagon',
  66: 'leftArrow',
  67: 'downArrow',
  68: 'upArrow',
  109: 'flowChartProcess',
  110: 'flowChartDecision',
  176: 'flowChartAlternateProcess',
};
const SPT_TEXTBOX = 202;
const SPT_PICTURE = 75;

/** OfficeArtCOLORREF：低 3 字节 R、G、B，最高字节是标志（调色板 / 配色方案 / 系统色，这里按缺省颜色处理） */
function colorRef(v: number | undefined, d: string): string {
  if (v === undefined || v >>> 24 !== 0) return d;
  return [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** 布尔属性组：位 bit 是值，bit + 16 是“这一位有效”；无效时用缺省值 */
function boolProp(v: number | undefined, bit: number, d: boolean): boolean {
  if (v === undefined || !(v & (1 << (bit + 16)))) return d;
  return !!(v & (1 << bit));
}

const EMU_PER_PT = 12700;

export interface ShapeEnv {
  /** pib（从 1 数）→ 登记好的图片 id */
  image(pib: number): string | null;
  /** 形状 id → 文本框文字（没有文字返回 null） */
  text(spid: number): Paragraph[] | null;
  /** 没有文字的段落（形状有文字时的空文本框） */
  emptyPara(): Paragraph;
  /** 记录不支持的内容 */
  lossy(what: string): void;
}

/**
 * 一个锚点的形状 → 文本框模型（组合拆成多个）。位置：外框相对 FSPA 的参照物；
 * 第三属性表里有 posrelh / posrelv / posh / posv（Word 2000 以后写的）时以它为准。
 */
export function shapeBoxes(sp: ArtShape, fspa: Fspa, env: ShapeEnv): TextBox[] {
  const o = sp.opt;
  const relH = ['margin', 'page', 'column', 'character'][o.get(0x0390) ?? fspa.bx] ?? 'column';
  const relV = ['margin', 'page', 'paragraph', 'line'][o.get(0x0392) ?? fspa.by] ?? 'paragraph';
  const alignH = ['', 'left', 'center', 'right', 'inside', 'outside'][o.get(0x038f) ?? 0] || undefined;
  const alignV = ['', 'top', 'center', 'bottom', 'inside', 'outside'][o.get(0x0391) ?? 0] || undefined;
  const wrap: TextBox['wrap'] = (['square', 'topAndBottom', 'square', 'none', 'tight', 'through'] as const)[fspa.wr] ?? 'square';
  const behind = fspa.below || boolProp(o.get(0x03bf), 5, false);
  const dist = {
    l: (o.get(0x0384) ?? 114300) / EMU_PER_PT,
    t: (o.get(0x0385) ?? 0) / EMU_PER_PT,
    r: (o.get(0x0386) ?? 114300) / EMU_PER_PT,
    b: (o.get(0x0387) ?? 0) / EMU_PER_PT,
  };
  const out: TextBox[] = [];
  const emit = (s: ArtShape, x: number, y: number, w: number, h: number) => {
    if (s.group) {
      const [gl, gt, gr, gb] = s.group.rect;
      const sx = gr !== gl ? w / (gr - gl) : 1;
      const sy = gb !== gt ? h / (gb - gt) : 1;
      for (const k of s.group.kids) {
        const c = k.child ?? [gl, gt, gr, gb];
        emit(k, x + (c[0] - gl) * sx, y + (c[1] - gt) * sy, (c[2] - c[0]) * sx, (c[3] - c[1]) * sy);
      }
      return;
    }
    const box = shapeBox(s, w, h, env);
    if (!box) return;
    const pos = (rel: string, offset: number, align: string | undefined): FloatPos => ({ rel, offset, ...(align && s === sp ? { align } : {}) });
    out.push({ ...box, posH: pos(relH, x, alignH), posV: pos(relV, y, alignV), wrap, dist, behind });
  };
  // 旋转 45°–135°、225°–315° 的形状，FSPA 里存的是转过 90° 后的外框（宽高互换、中心不变）
  let { l, t, r, b } = fspa;
  const rot = (((o.get(0x0004) ?? 0) | 0) / 65536 + 360) % 360;
  if (!sp.group && ((rot >= 45 && rot < 135) || (rot >= 225 && rot < 315))) {
    const cx = (l + r) / 2;
    const cy = (t + b) / 2;
    const hw = (b - t) / 2;
    const hh = (r - l) / 2;
    [l, t, r, b] = [cx - hw, cy - hh, cx + hw, cy + hh];
  }
  emit(sp, l / 20, t / 20, (r - l) / 20, (b - t) / 20);
  return out;
}

/** 单个形状（不是组合）的外观；不支持的形状返回 null 并记录 */
function shapeBox(s: ArtShape, w: number, h: number, env: ShapeEnv): Omit<TextBox, 'posH' | 'posV' | 'wrap' | 'dist' | 'behind'> | null {
  const o = s.opt;
  if (boolProp(o.get(0x03bf), 1, false)) return null; // fHidden
  const base = {
    w: Math.max(0, w),
    h: Math.max(0, h),
    inset: { l: (o.get(0x0081) ?? 91440) / EMU_PER_PT, t: (o.get(0x0082) ?? 45720) / EMU_PER_PT, r: (o.get(0x0083) ?? 91440) / EMU_PER_PT, b: (o.get(0x0084) ?? 45720) / EMU_PER_PT },
    vAlign: ([0, 3].includes(o.get(0x0087) ?? 0) ? 'top' : [1, 4].includes(o.get(0x0087)!) ? 'center' : [2, 5].includes(o.get(0x0087)!) ? 'bottom' : 'top') as TextBox['vAlign'],
  };
  const pib = o.get(0x0104);
  if (s.spt === SPT_PICTURE || (pib && !SPT_GEOM[s.spt] && s.spt !== SPT_TEXTBOX)) {
    const id = pib ? env.image(pib) : null;
    if (!id) {
      env.lossy('图片、公式或对象（格式暂不支持）');
      return null;
    }
    return { ...base, fill: null, line: null, inset: { l: 0, t: 0, r: 0, b: 0 }, vAlign: 'top', paragraphs: [], image: { id, w: base.w, h: base.h } };
  }
  const geom = s.spt === SPT_TEXTBOX ? undefined : SPT_GEOM[s.spt];
  if (s.spt !== SPT_TEXTBOX && !geom) {
    env.lossy('自由曲线、艺术字等形状');
    return null;
  }
  const filled = boolProp(o.get(0x01bf), 4, true);
  const lined = boolProp(o.get(0x01ff), 3, true);
  const text = env.text(s.spid);
  const rot = ((o.get(0x0004) ?? 0) | 0) / 65536;
  return {
    ...base,
    fill: filled && !isLineGeom(geom) ? colorRef(o.get(0x0181), 'FFFFFF') : null,
    line: lined ? { color: colorRef(o.get(0x01c0), '000000'), w: (o.get(0x01cb) ?? 9525) / EMU_PER_PT } : null,
    paragraphs: text ?? (s.spt === SPT_TEXTBOX ? [env.emptyPara()] : []),
    ...(geom ? { geom } : {}),
    ...(s.flipH ? { flipH: true } : {}),
    ...(s.flipV ? { flipV: true } : {}),
    ...(rot ? { rot: ((rot % 360) + 360) % 360 } : {}),
  };
}

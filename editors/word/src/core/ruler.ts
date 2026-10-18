/**
 * 标尺：把段落缩进、制表位和页边距换成标尺上的位置，拖动标尺上的标记再换回格式。
 * 全部是纯函数（单位磅，位置从“段落区域左边”算起；区域是版心、分栏的一栏）。
 *
 * 规则取自 Word / OOXML：
 * - 首行缩进标记（上三角）= 左缩进 + 首行缩进（悬挂缩进为负的首行缩进，w:hanging）；
 * - 悬挂缩进标记（下三角）只移动左缩进，首行位置不变，所以首行缩进 / 悬挂值随之改变；
 * - 左缩进标记（方块）连同首行标记一起移动：左缩进变，首行缩进值不变；
 * - 右缩进标记在区域右边往左数；
 * - 以“字符”为单位的首行缩进（firstLineChars）按字号折成磅显示；一旦用标尺改首行就以磅为准（清掉字符数）；
 * - 缩进可以是负数（伸到页边距里），但不超过纸边，也不让左右缩进互相越过。
 *
 * 还没有支持：表格单元格里的缩进（标尺只显示页边距）、页眉页脚、竖直标尺、小数点对齐制表位、
 * 多栏页面的栏间距拖动。
 */
import { PT_PER_CM, columnBoxes, textWidth, type PageGeom, type PageSetup, type ParaProps, type TabStop } from './model';

export type TabAlign = TabStop['align'];

/** 标尺上的位置（磅，从段落区域左边算起） */
export interface RulerMarks {
  /** 首行起点 */
  first: number;
  /** 左缩进（悬挂位置） */
  left: number;
  /** 右缩进标记（区域右边往左缩进后的位置） */
  right: number;
  tabs: TabStop[];
}

/** 首行缩进折成磅：字符数优先（1 字 = 字号） */
export function firstLinePt(p: Pick<ParaProps, 'firstLine' | 'firstLineChars'>, fontSize: number): number {
  return p.firstLineChars != null ? (p.firstLineChars / 100) * fontSize : p.firstLine;
}

export function rulerMarks(p: ParaProps, areaW: number, fontSize: number): RulerMarks {
  return {
    first: p.indentLeft + firstLinePt(p, fontSize),
    left: p.indentLeft,
    right: areaW - p.indentRight,
    tabs: p.tabs ?? [],
  };
}

export type IndentMarker = 'first' | 'hanging' | 'left' | 'right';

export interface RulerArea {
  /** 区域左边到页面左边的距离 */
  x: number;
  w: number;
  /** 区域右边到页面右边的距离 */
  r: number;
}

/**
 * 段落所在区域（版心的一栏）：lineX 是排版结果里该段落的区域左边（页面坐标），
 * 在分栏页上取最近的一栏；一栏时就是版心。
 */
export function paragraphArea(setup: PageSetup, lineX: number): RulerArea {
  const boxes = columnBoxes(setup);
  let best = boxes[0] ?? { x: 0, w: textWidth(setup) };
  for (const b of boxes) if (Math.abs(setup.marginLeft + b.x - lineX) < Math.abs(setup.marginLeft + best.x - lineX)) best = b;
  const x = setup.marginLeft + best.x;
  return { x, w: best.w, r: setup.width - x - best.w };
}

/** 拖到 x（区域坐标，磅）后的段落格式修改；没有变化时返回 null */
export function indentPatch(
  marker: IndentMarker,
  x: number,
  p: ParaProps,
  area: RulerArea,
  fontSize: number,
): Partial<ParaProps> | null {
  const marks = rulerMarks(p, area.w, fontSize);
  const minX = -area.x; // 不超过纸的左边
  const r = (v: number) => Math.round(v * 100) / 100;
  let patch: Partial<ParaProps>;
  switch (marker) {
    case 'first': {
      const at = Math.min(Math.max(x, minX), area.w + area.r);
      patch = { firstLine: r(at - p.indentLeft), firstLineChars: null };
      break;
    }
    case 'hanging': {
      // 只动左缩进，首行位置不变
      const at = Math.min(Math.max(x, minX), marks.right - 1);
      patch = { indentLeft: r(at), firstLine: r(marks.first - at), firstLineChars: null };
      break;
    }
    case 'left': {
      // 左缩进和首行标记一起动：首行缩进值不变；两个标记都不能越出纸边或越过右缩进
      const lo = minX - Math.min(0, marks.first - p.indentLeft);
      const hi = marks.right - 1 - Math.max(0, marks.first - p.indentLeft);
      patch = { indentLeft: r(Math.min(Math.max(x, lo), Math.max(lo, hi))) };
      break;
    }
    case 'right': {
      const at = Math.min(Math.max(x, Math.max(marks.left, marks.first) + 1), area.w + area.r);
      patch = { indentRight: r(area.w - at) };
      break;
    }
  }
  const changed = (Object.keys(patch) as (keyof ParaProps)[]).some((k) => patch[k] !== p[k]);
  return changed ? patch : null;
}

/** 拖动页边距线：side 的新位置 x（页面坐标，磅）；版心至少留 minText 磅。返回新的页面几何，没有变化时返回 null */
export function marginPatch(side: 'left' | 'right', x: number, setup: PageSetup, minText = 36): PageGeom | null {
  const geom = marginGeom(setup);
  if (side === 'left') {
    const m = Math.min(Math.max(x, 0), setup.width - setup.marginRight - minText);
    geom.marginLeft = Math.round(m * 100) / 100;
  } else {
    const m = Math.min(Math.max(setup.width - x, 0), setup.width - setup.marginLeft - minText);
    geom.marginRight = Math.round(m * 100) / 100;
  }
  return geom.marginLeft === setup.marginLeft && geom.marginRight === setup.marginRight ? null : geom;
}

/** 拖动纵向标尺上的页边距线：side 的新位置 y（页面坐标，磅）；版心至少留 minText 磅 */
export function marginPatchV(side: 'top' | 'bottom', y: number, setup: PageSetup, minText = 36): PageGeom | null {
  const geom = marginGeom(setup);
  if (side === 'top') geom.marginTop = Math.round(Math.min(Math.max(y, 0), setup.height - setup.marginBottom - minText) * 100) / 100;
  else geom.marginBottom = Math.round(Math.min(Math.max(setup.height - y, 0), setup.height - setup.marginTop - minText) * 100) / 100;
  return geom.marginTop === setup.marginTop && geom.marginBottom === setup.marginBottom ? null : geom;
}

function marginGeom(setup: PageSetup): PageGeom {
  return Object.fromEntries(
    (['width', 'height', 'marginTop', 'marginBottom', 'marginLeft', 'marginRight', 'headerDistance', 'footerDistance'] as const).map((k) => [k, setup[k]]),
  ) as PageGeom;
}

const SAME_TAB = 1.5;

/** 在 pos 处加一个制表位（已有相近的制表位就改它的对齐方式）；位置在区域外时不加 */
export function addTab(tabs: TabStop[] | undefined, pos: number, align: TabAlign, areaW: number): TabStop[] | null {
  if (pos < 0 || pos > areaW) return null;
  const list = (tabs ?? []).filter((t) => Math.abs(t.pos - pos) >= SAME_TAB);
  list.push({ pos: Math.round(pos * 100) / 100, align, leader: 'none' });
  return list.sort((a, b) => a.pos - b.pos);
}

/** 把第 i 个制表位移到 pos；pos 为 null 表示拖出标尺，删除。保留对齐方式和前导符 */
export function moveTab(tabs: TabStop[], i: number, pos: number | null, areaW: number): TabStop[] {
  const rest = tabs.filter((_, k) => k !== i);
  const t = tabs[i];
  if (!t || pos === null || pos < 0 || pos > areaW) return rest;
  const others = rest.filter((o) => Math.abs(o.pos - pos) >= SAME_TAB);
  others.push({ ...t, pos: Math.round(pos * 100) / 100 });
  return others.sort((a, b) => a.pos - b.pos);
}

/** 命中第几个制表位（容差磅）；没有返回 -1 */
export function tabAt(tabs: TabStop[], pos: number, tol: number): number {
  let best = -1;
  let d = tol;
  tabs.forEach((t, i) => {
    const k = Math.abs(t.pos - pos);
    if (k <= d) (d = k), (best = i);
  });
  return best;
}

/** 吸附到 step 磅的整数倍（按住 Alt 不吸附） */
export function snap(v: number, step: number): number {
  return step > 0 ? Math.round(v / step) * step : v;
}

/** 标尺刻度：厘米，每 0.5 厘米一条短线、每厘米一个数字；origin 为 0 点（磅），范围 [from, to] 磅 */
export function rulerTicks(origin: number, from: number, to: number): { pt: number; cm: number; major: boolean; half: boolean }[] {
  const out: { pt: number; cm: number; major: boolean; half: boolean }[] = [];
  const first = Math.ceil((from - origin) / (PT_PER_CM / 2));
  const last = Math.floor((to - origin) / (PT_PER_CM / 2));
  for (let i = first; i <= last; i++) {
    out.push({ pt: origin + (i * PT_PER_CM) / 2, cm: i / 2, major: i % 2 === 0, half: i % 2 !== 0 });
  }
  return out;
}

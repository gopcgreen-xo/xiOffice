/**
 * 行列分组（大纲）：组合 / 取消组合、折叠 / 展开、按级别显示。
 *
 * 规则同 Excel：
 * - 每行 / 列有分组级别 0–7（`outlineLevel`）；一组 = 紧挨着汇总项、级别比汇总项高的连续行 / 列。
 * - 汇总项默认在明细下方（行）/ 右侧（列），`<outlinePr summaryBelow / summaryRight>` 可以改成上方 / 左侧。
 * - 折叠 = 汇总项记 `collapsed`，它管的那组行 / 列隐藏；嵌套时外层折叠会连内层一起藏起来。
 * - 级别按钮 k：只显示级别 ≤ k-1 的行 / 列。
 * 隐藏状态以 `hidden` 为准（显示、保存都读它），`collapsed` 只用来决定展开时要显示哪些。
 */
import { mergeDefs, splitDefs, type Tx } from './ops';
import { deriveAxes, type ColDef, type RowDef, type SheetData } from './workbook';

export type Axis = 'row' | 'col';
export const MAX_LEVEL = 7;

const colDefAt = (sheet: SheetData, c: number): ColDef | undefined => sheet.colDefs?.find((d) => c >= d.min && c <= d.max);

export function levelAt(sheet: SheetData, axis: Axis, i: number): number {
  if (i < 0) return 0;
  return (axis === 'row' ? sheet.rowDefs?.[i]?.level : colDefAt(sheet, i)?.level) ?? 0;
}

export function collapsedAt(sheet: SheetData, axis: Axis, i: number): boolean {
  return !!(axis === 'row' ? sheet.rowDefs?.[i]?.collapsed : colDefAt(sheet, i)?.collapsed);
}

/** 当前工作表里最大的分组级别（没有分组 = 0） */
export function maxLevel(sheet: SheetData, axis: Axis): number {
  if (axis === 'row') return Math.max(0, ...Object.values(sheet.rowDefs ?? {}).map((d) => d.level ?? 0));
  return Math.max(0, ...(sheet.colDefs ?? []).map((d) => d.level ?? 0));
}

/** 汇总项在明细的后面（行：下方；列：右侧）吗 */
export function summaryAfter(sheet: SheetData, axis: Axis): boolean {
  return axis === 'row' ? sheet.outline?.summaryBelow !== false : sheet.outline?.summaryRight !== false;
}

/** 汇总项 s 管着的那一组：紧挨着它、级别比它高的连续行 / 列（返回 [起, 止]，没有返回 null） */
export function groupOf(sheet: SheetData, axis: Axis, s: number): [number, number] | null {
  const L = levelAt(sheet, axis, s);
  const step = summaryAfter(sheet, axis) ? -1 : 1;
  const a = s + step;
  if (a < 0 || levelAt(sheet, axis, a) <= L) return null;
  let b = a;
  while (b + step >= 0 && levelAt(sheet, axis, b + step) > L) b += step;
  return step < 0 ? [b, a] : [a, b];
}

/** i 是某一组的汇总项吗（它旁边的行 / 列级别比它高） */
export const isSummary = (sheet: SheetData, axis: Axis, i: number): boolean => groupOf(sheet, axis, i) !== null;

// ---------- 修改 ----------

/** 逐行 / 列改属性；改完没有任何属性的行定义删掉 */
function edit(tx: Tx, si: number, axis: Axis, from: number, to: number, fn: (d: { level?: number; collapsed?: boolean; hidden?: boolean }, i: number) => void): void {
  const sheet = tx.book.sheets[si];
  tx.meta(si);
  if (axis === 'row') {
    const defs = (sheet.rowDefs ??= {});
    for (let r = from; r <= to; r++) {
      const d: RowDef = { ...defs[r] };
      fn(d, r);
      if (Object.keys(d).length) defs[r] = d;
      else delete defs[r];
    }
    return;
  }
  const defs = splitDefs(sheet.colDefs ?? [], from, to);
  for (let c = from; c <= to; c++) {
    let d = defs.find((x) => x.min === c && x.max === c);
    if (!d) defs.push((d = { min: c, max: c }));
    fn(d, c);
  }
  sheet.colDefs = mergeDefs(defs);
}

const clean = (d: { level?: number; collapsed?: boolean }) => {
  if (!d.level) {
    delete d.level;
    delete d.collapsed;
  }
  if (!d.collapsed) delete d.collapsed;
};

/** 按折叠标记重新算哪些行 / 列该隐藏（只动有分组级别的行 / 列） */
function refreshHidden(tx: Tx, si: number, axis: Axis): void {
  const sheet = tx.book.sheets[si];
  const grouped: number[] = [];
  const collapsed: number[] = [];
  if (axis === 'row') {
    for (const [k, d] of Object.entries(sheet.rowDefs ?? {})) {
      if (d.level) grouped.push(Number(k));
      if (d.collapsed) collapsed.push(Number(k));
    }
  } else {
    for (const d of sheet.colDefs ?? []) {
      for (let c = d.min; c <= Math.min(d.max, d.min + 16384); c++) {
        if (d.level) grouped.push(c);
        if (d.collapsed) collapsed.push(c);
      }
    }
  }
  const hide = new Set<number>();
  for (const s of collapsed) {
    const g = groupOf(sheet, axis, s);
    if (g) for (let i = g[0]; i <= g[1]; i++) hide.add(i);
  }
  const lo = Math.min(...grouped, ...collapsed);
  const hi = Math.max(...grouped, ...collapsed);
  if (!Number.isFinite(lo)) return;
  // 汇总项已经没有可折叠的组了：去掉过期的折叠标记
  const stale = new Set(collapsed.filter((s) => !groupOf(sheet, axis, s)));
  const mark = new Set(grouped);
  edit(tx, si, axis, lo, hi, (d, i) => {
    if (stale.has(i)) delete d.collapsed;
    if (!mark.has(i)) return;
    if (hide.has(i)) d.hidden = true;
    else delete d.hidden;
  });
  deriveAxes(sheet);
}

/** 组合：选中的行 / 列级别 +1 */
export function groupRange(tx: Tx, si: number, axis: Axis, from: number, to: number): void {
  edit(tx, si, axis, from, to, (d) => {
    d.level = Math.min(MAX_LEVEL, (d.level ?? 0) + 1);
  });
  deriveAxes(tx.book.sheets[si]);
}

/** 取消组合：选中的行 / 列级别 -1；回到 0 的行 / 列显示出来 */
export function ungroupRange(tx: Tx, si: number, axis: Axis, from: number, to: number): void {
  edit(tx, si, axis, from, to, (d) => {
    d.level = Math.max(0, (d.level ?? 0) - 1);
    if (!d.level) delete d.hidden;
    clean(d);
  });
  refreshHidden(tx, si, axis);
}

/** 清除选中范围的所有分组 */
export function clearGroups(tx: Tx, si: number, axis: Axis, from: number, to: number): void {
  edit(tx, si, axis, from, to, (d) => {
    delete d.level;
    delete d.collapsed;
    delete d.hidden;
  });
  refreshHidden(tx, si, axis);
}

/** 折叠 / 展开汇总项 s 管着的那一组 */
export function toggleGroup(tx: Tx, si: number, axis: Axis, s: number): void {
  const sheet = tx.book.sheets[si];
  if (!groupOf(sheet, axis, s)) return;
  const next = !collapsedAt(sheet, axis, s);
  edit(tx, si, axis, s, s, (d) => {
    if (next) d.collapsed = true;
    else delete d.collapsed;
  });
  refreshHidden(tx, si, axis);
}

/** 级别按钮 k（1 起）：只显示级别 ≤ k-1 的行 / 列 */
export function showLevel(tx: Tx, si: number, axis: Axis, k: number): void {
  const sheet = tx.book.sheets[si];
  const top = maxLevel(sheet, axis);
  if (!top) return;
  // 汇总项的范围：有分组的行 / 列前后各 1
  let lo = Infinity;
  let hi = -1;
  if (axis === 'row') {
    for (const [key, d] of Object.entries(sheet.rowDefs ?? {})) if (d.level) (lo = Math.min(lo, Number(key))), (hi = Math.max(hi, Number(key)));
  } else {
    for (const d of sheet.colDefs ?? []) if (d.level) (lo = Math.min(lo, d.min)), (hi = Math.max(hi, d.max));
  }
  if (hi < 0) return;
  const want = new Map<number, boolean>();
  for (let i = Math.max(0, lo - 1); i <= hi + 1; i++) if (isSummary(sheet, axis, i)) want.set(i, levelAt(sheet, axis, i) >= k - 1);
  edit(tx, si, axis, Math.max(0, lo - 1), hi + 1, (d, i) => {
    const w = want.get(i);
    if (w === undefined) return;
    if (w) d.collapsed = true;
    else delete d.collapsed;
  });
  refreshHidden(tx, si, axis);
}

import { describe, expect, it } from 'vitest';
import { PT_PER_CM, newDocument, newParagraph } from './model';
import { addTab, indentPatch, marginPatch, marginPatchV, moveTab, paragraphArea, rulerMarks, rulerTicks, snap, tabAt } from './ruler';

const doc = newDocument();
const setup = doc.setup;
const base = newParagraph().props;
const area = paragraphArea(setup, setup.marginLeft);

describe('标尺：段落缩进（Word 的通用规则）', () => {
  it('首行标记 = 左缩进 + 首行缩进；悬挂缩进为负的首行缩进', () => {
    expect(rulerMarks({ ...base, indentLeft: 36, firstLine: 24 }, area.w, 12)).toMatchObject({ first: 60, left: 36 });
    expect(rulerMarks({ ...base, indentLeft: 36, firstLine: -18 }, area.w, 12)).toMatchObject({ first: 18, left: 36 });
  });

  it('以字符为单位的首行缩进按字号折成磅，用标尺改首行后以磅为准', () => {
    const p = { ...base, firstLine: 0, firstLineChars: 200 };
    expect(rulerMarks(p, area.w, 12).first).toBe(24);
    expect(indentPatch('first', 30, p, area, 12)).toEqual({ firstLine: 30, firstLineChars: null });
  });

  it('拖首行标记只改首行缩进，左缩进不变', () => {
    const p = { ...base, indentLeft: 20 };
    expect(indentPatch('first', 50, p, area, 12)).toEqual({ firstLine: 30, firstLineChars: null });
    expect(indentPatch('first', 5, p, area, 12)).toEqual({ firstLine: -15, firstLineChars: null });
  });

  it('拖悬挂标记：左缩进变，首行位置不变（首行缩进值随之改变）', () => {
    const p = { ...base, indentLeft: 0, firstLine: 24 }; // 首行在 24
    expect(indentPatch('hanging', 40, p, area, 12)).toEqual({ indentLeft: 40, firstLine: -16, firstLineChars: null });
  });

  it('拖左缩进标记：首行标记一起移动，首行缩进值不变', () => {
    const p = { ...base, indentLeft: 10, firstLine: -10 };
    expect(indentPatch('left', 30, p, area, 12)).toEqual({ indentLeft: 30 });
  });

  it('拖右缩进标记：从区域右边往左数', () => {
    expect(indentPatch('right', area.w - 25, base, area, 12)).toEqual({ indentRight: 25 });
  });

  it('可以伸到页边距里，但不超过纸边，也不越过对方', () => {
    expect(indentPatch('hanging', -9999, base, area, 12)).toMatchObject({ indentLeft: -area.x });
    expect(indentPatch('right', 9999, base, area, 12)).toEqual({ indentRight: -area.r });
    const p = { ...base, indentRight: 100 };
    const hang = indentPatch('hanging', 9999, p, area, 12)!;
    expect(hang.indentLeft!).toBeLessThan(area.w - 100);
    // 悬挂缩进的首行标记在左缩进左边时，整体左移也不让首行标记越出纸边
    const q = { ...base, indentLeft: 0, firstLine: -30 };
    expect(indentPatch('left', -9999, q, area, 12)!.indentLeft).toBe(-area.x + 30);
  });

  it('没有变化返回 null', () => {
    expect(indentPatch('first', 0, base, area, 12)).toBeNull();
    expect(indentPatch('left', 0, base, area, 12)).toBeNull();
  });
});

describe('标尺：页边距与制表位', () => {
  it('拖页边距线：改左 / 右页边距，版心至少留一点', () => {
    const g = marginPatch('left', 3 * PT_PER_CM, setup)!;
    expect(g.marginLeft).toBeCloseTo(3 * PT_PER_CM, 1);
    expect(g.marginRight).toBe(setup.marginRight);
    const r = marginPatch('right', setup.width - 2 * PT_PER_CM, setup)!;
    expect(r.marginRight).toBeCloseTo(2 * PT_PER_CM, 1);
    expect(setup.width - setup.marginRight - marginPatch('left', 9999, setup)!.marginLeft).toBeCloseTo(36, 1);
    expect(marginPatch('left', setup.marginLeft, setup)).toBeNull();
  });

  it('拖纵向标尺的页边距线：改上 / 下页边距，版心至少留一点', () => {
    const t = marginPatchV('top', 3 * PT_PER_CM, setup)!;
    expect(t.marginTop).toBeCloseTo(3 * PT_PER_CM, 1);
    expect(t.marginBottom).toBe(setup.marginBottom);
    const b = marginPatchV('bottom', setup.height - 2 * PT_PER_CM, setup)!;
    expect(b.marginBottom).toBeCloseTo(2 * PT_PER_CM, 1);
    expect(setup.height - setup.marginBottom - marginPatchV('top', 99999, setup)!.marginTop).toBeCloseTo(36, 1);
    expect(marginPatchV('top', -50, setup)!.marginTop).toBe(0);
    expect(marginPatchV('top', setup.marginTop, setup)).toBeNull();
  });

  it('制表位：添加、相近的合并、排序、拖出标尺删除、移动保留对齐方式', () => {
    let tabs = addTab(undefined, 100, 'left', area.w)!;
    tabs = addTab(tabs, 50, 'center', area.w)!;
    expect(tabs.map((t) => [t.pos, t.align])).toEqual([[50, 'center'], [100, 'left']]);
    expect(addTab(tabs, 100.5, 'right', area.w)!.map((t) => t.align)).toEqual(['center', 'right']);
    expect(addTab(tabs, area.w + 10, 'left', area.w)).toBeNull();
    const moved = moveTab(tabs, 0, 150, area.w);
    expect(moved.map((t) => [t.pos, t.align])).toEqual([[100, 'left'], [150, 'center']]);
    expect(moveTab(tabs, 1, null, area.w)).toHaveLength(1);
    expect(tabAt(tabs, 52, 4)).toBe(0);
    expect(tabAt(tabs, 75, 4)).toBe(-1);
  });

  it('吸附与刻度', () => {
    expect(snap(10, 2.835)).toBeCloseTo(11.34, 1);
    expect(snap(10, 0)).toBe(10);
    const ticks = rulerTicks(100, 100, 100 + 2 * PT_PER_CM);
    expect(ticks.filter((t) => t.major).map((t) => t.cm)).toEqual([0, 1, 2]);
    expect(ticks.filter((t) => t.half)).toHaveLength(2);
  });

  it('分栏页上取段落所在的那一栏', () => {
    const s = { ...setup, cols: { num: 2, space: 20, sep: false } };
    const second = paragraphArea(s, s.marginLeft + (s.width - s.marginLeft - s.marginRight - 20) / 2 + 20);
    const first = paragraphArea(s, s.marginLeft);
    expect(second.x).toBeGreaterThan(first.x);
    expect(first.w).toBeCloseTo(second.w, 5);
  });
});

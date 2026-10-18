import { describe, expect, it } from 'vitest';
import { layoutDocument, hitTest, type Measurer } from './layout';
import { columnBoxes, newDocument, newParagraph, textWidth, type Doc } from './model';
import { isCJK } from './text';

/** 假字体：汉字 1 em，西文 0.5 em */
const fake: Measurer = {
  width: (ch) => (isCJK(ch) ? 1 : 0.5),
  metrics: () => ({ ascent: 0.88, descent: 0.12 }),
};

function twoCols(): Doc {
  const d = newDocument();
  d.setup.cols = { num: 2, space: 21, sep: true };
  return d;
}

describe('分栏（w:cols）', () => {
  it('等宽分栏：栏宽 = (版心 − 间距) / 2，文字按栏宽换行，排满第一栏再排第二栏', () => {
    const d = twoCols();
    const [c0, c1] = columnBoxes(d.setup);
    expect(c0.w).toBeCloseTo((textWidth(d.setup) - 21) / 2);
    expect(c1.x).toBeCloseTo(c0.w + 21);
    d.body = Array.from({ length: 60 }, () => newParagraph('国'.repeat(30)));
    const L = layoutDocument(d, fake);
    const p0 = L.pages[0];
    // 每行字数按栏宽
    expect(p0.body[0].line.end - p0.body[0].line.start).toBe(Math.floor((c0.w + 0.01) / 10.5));
    const xs = [...new Set(p0.body.map((l) => Math.round(l.x)))];
    expect(xs).toEqual([Math.round(d.setup.marginLeft), Math.round(d.setup.marginLeft + c1.x)]);
    // 第二栏从版心顶部开始，在第一栏之后
    const second = p0.body.find((l) => l.x > d.setup.marginLeft + 1)!;
    expect(second.y).toBeCloseTo(d.setup.marginTop);
    expect(p0.body.indexOf(second)).toBeGreaterThan(0);
    // 分隔线在两栏间距中间
    expect(p0.borders.some((b) => Math.abs(b.x1 - (d.setup.marginLeft + c0.w + 10.5)) < 0.01 && b.x1 === b.x2)).toBe(true);
  });

  it('连续分节符前的分栏平衡：两栏高度相同，后一节从栏下面开始', () => {
    const d = newDocument();
    const end = newParagraph('国', { sectionBreak: 'continuous' });
    end.sectionCols = { num: 2, space: 21, sep: false };
    d.body = [...Array.from({ length: 9 }, () => newParagraph('国')), end, newParagraph('后一节')];
    const L = layoutDocument(d, fake);
    const body = L.pages[0].body;
    const left = body.filter((l) => l.x < d.setup.marginLeft + 1 && l.path[0] < 10);
    const right = body.filter((l) => l.x > d.setup.marginLeft + 1);
    expect(left.length).toBe(5);
    expect(right.length).toBe(5);
    const last = body[body.length - 1];
    expect(last.path[0]).toBe(10);
    expect(last.x).toBeCloseTo(d.setup.marginLeft);
    expect(last.y).toBeCloseTo(Math.max(...left.map((l) => l.y + l.line.height)));
  });

  it('分栏符（w:br w:type="column"）之后的内容从下一栏顶部开始；最后一栏的分栏符换页', () => {
    const d = twoCols();
    d.body = [newParagraph('甲\v乙'), newParagraph('丙\v丁')];
    const L = layoutDocument(d, fake);
    const at = (s: string) => {
      for (const pg of L.pages) for (const l of pg.body) if (l.line.glyphs.some((g) => g.ch === s)) return { page: pg.index, x: l.x, y: l.y };
      throw new Error(s);
    };
    const [, c1] = columnBoxes(d.setup);
    expect(at('乙')).toEqual({ page: 0, x: d.setup.marginLeft + c1.x, y: d.setup.marginTop });
    expect(at('丙').page).toBe(0);
    expect(at('丁')).toMatchObject({ page: 1, x: d.setup.marginLeft, y: d.setup.marginTop });
  });

  it('不等宽分栏：段落换到另一栏时按那一栏的宽度重新断行，段首缩进只在段首', () => {
    const d = newDocument();
    const tw = textWidth(d.setup);
    d.setup.cols = { num: 2, space: 15, sep: false, widths: [{ w: 105, space: 15 }, { w: tw - 120, space: 0 }] };
    d.body = [newParagraph('国'.repeat(500), { firstLineChars: 2 })];
    const p0 = layoutDocument(d, fake).pages[0].body;
    const left = p0.filter((l) => l.x < d.setup.marginLeft + 1);
    const right = p0.filter((l) => l.x > d.setup.marginLeft + 1);
    const n = (l: (typeof p0)[number]) => l.line.end - l.line.start;
    expect(n(left[0])).toBe(8); // 首行缩进两字
    expect(n(left[1])).toBe(10);
    expect(n(right[0])).toBe(Math.floor((tw - 120 + 0.01) / 10.5));
    expect(right[0].line.start).toBe(left[left.length - 1].line.end);
  });

  it('Word 实测：只含连续分节符的空段落不占高度', () => {
    const d = newDocument();
    const br = newParagraph('', { sectionBreak: 'continuous' });
    br.sectionCols = undefined;
    d.body = [newParagraph('甲'), br, newParagraph('乙')];
    d.setup.cols = { num: 2, space: 21, sep: false };
    const body = layoutDocument(d, fake).pages[0].body;
    const y = (bi: number) => body.find((l) => l.path[0] === bi)!;
    expect(y(1).line.height).toBe(0);
    expect(y(2).y).toBeCloseTo(y(0).y + y(0).line.height);
  });

  it('点击第二栏下方的空白处定位到第二栏，而不是第一栏同一高度的行', () => {
    const d = twoCols();
    d.body = Array.from({ length: 50 }, () => newParagraph('国'));
    // 最后一段放在第二栏顶部附近：在第一栏排满后加几段
    const L = layoutDocument(d, fake);
    const p0 = L.pages[0];
    const firstRight = p0.body.findIndex((l) => l.x > d.setup.marginLeft + 1);
    const n = firstRight + 3;
    d.body = d.body.slice(0, n);
    const L2 = layoutDocument(d, fake);
    const [, c1] = columnBoxes(d.setup);
    const pos = hitTest(L2, 0, d.setup.marginLeft + c1.x + 5, d.setup.marginTop + 500, 'body')!;
    expect(pos.path[0]).toBe(n - 1);
  });
});

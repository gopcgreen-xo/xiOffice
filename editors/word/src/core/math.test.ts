// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { layoutMath } from './math';
import { layoutDocument, type Measurer } from './layout';
import { FIELD_CHAR, newDocument, newParagraph } from './model';

/** 假字体：每个字 0.5 em，上伸 0.8 下伸 0.2 */
const fake: Measurer = { width: () => 0.5, metrics: () => ({ ascent: 0.8, descent: 0.2 }) };
const r = (t: string, sty = '') => `<m:r>${sty ? `<m:rPr><m:sty m:val="${sty}"/></m:rPr>` : ''}<m:t>${t}</m:t></m:r>`;
const math = (inner: string, para = false) => (para ? `<m:oMathPara><m:oMath>${inner}</m:oMath></m:oMathPara>` : `<m:oMath>${inner}</m:oMath>`);
const texts = (xml: string) => layoutMath(xml, 10, '宋体', fake)!.ops.filter((o) => o.t === 'text') as { s: string; x: number; y: number; size: number; italic: boolean }[];

describe('公式（OMML）显示', () => {
  it('文字：字母斜体、数字和运算符正体；关系符两边 5/18 em、二元运算符两边 4/18 em', () => {
    const t = texts(math(r('a=b+1')));
    expect(t.map((o) => o.s).join('')).toBe('a=b+1');
    expect(t.map((o) => o.italic)).toEqual([true, false, true, false, false]);
    // a 宽 5，= 前后各 25/9
    expect(t[1].x).toBeCloseTo(5 + 50 / 18);
    expect(t[2].x).toBeCloseTo(5 + 50 / 18 + 5 + 50 / 18);
    expect(t[3].x - (t[2].x + 5)).toBeCloseTo(40 / 18);
    // m:sty = p：正体
    expect(texts(math(r('sin', 'p'))).every((o) => !o.italic)).toBe(true);
  });

  it('分数：分子在分数线上、分母在下，居中；有分数线', () => {
    const b = layoutMath(math(`<m:f><m:num>${r('1')}</m:num><m:den>${r('22')}</m:den></m:f>`), 10, '宋体', fake)!;
    const [num, d1] = b.ops.filter((o) => o.t === 'text') as { x: number; y: number }[];
    const bar = b.ops.find((o) => o.t === 'line') as { y1: number };
    expect(num.y).toBeLessThan(bar.y1);
    expect(d1.y).toBeGreaterThan(bar.y1);
    expect(num.x).toBeGreaterThan(d1.x); // 分子比分母短，居中
    expect(b.asc).toBeGreaterThan(10);
    expect(b.desc).toBeGreaterThan(2);
  });

  it('上下标缩小到 70%，上标在基线上方、下标在下方', () => {
    const t = texts(math(`<m:sSubSup><m:e>${r('x')}</m:e><m:sub>${r('i')}</m:sub><m:sup>${r('2')}</m:sup></m:sSubSup>`));
    const [x, i, two] = ['x', 'i', '2'].map((c) => t.find((o) => o.s === c)!);
    expect(i.size).toBeCloseTo(7);
    expect(two.y).toBeLessThan(0);
    expect(i.y).toBeGreaterThan(0);
    expect(i.x).toBeGreaterThan(x.x);
  });

  it('括号（m:d）：缺省圆括号，按内容高度放大', () => {
    const t = texts(math(`<m:d><m:e>${r('a')}</m:e></m:d>`));
    expect(t.map((o) => o.s).join('')).toBe('(a)');
    const tall = texts(math(`<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e><m:f><m:num>${r('1')}</m:num><m:den>${r('2')}</m:den></m:f></m:e></m:d>`));
    expect(tall[0].s).toBe('[');
    expect(tall[0].size).toBeGreaterThan(10);
  });

  it('排进段落：公式占一个字符的位置、宽度按排版结果；单独成段的公式居中；分数让行变高', () => {
    const d = newDocument();
    const frac = math(`<m:f><m:num>${r('1')}</m:num><m:den>${r('2')}</m:den></m:f>`, true);
    const p = newParagraph('');
    p.runs = [{ text: FIELD_CHAR, props: d.defaults, opaque: { xml: frac, level: 'para', kind: 'math', w: 10, h: 10 } }];
    d.body = [p, newParagraph('国')];
    const lines = layoutDocument(d, fake).pages[0].body;
    const g = lines[0].line.glyphs[0];
    expect(g.math).toBeDefined();
    const tw = d.setup.width - d.setup.marginLeft - d.setup.marginRight;
    expect(g.x).toBeCloseTo((tw - g.math!.w) / 2);
    expect(lines[0].line.height).toBeGreaterThan(lines[1].line.height);
  });
});

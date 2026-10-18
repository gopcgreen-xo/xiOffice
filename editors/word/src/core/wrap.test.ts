import { describe, expect, it } from 'vitest';
import { layoutDocument, type Measurer } from './layout';
import { FIELD_CHAR, newDocument, newParagraph, newTable, textWidth, type Doc, type Paragraph, type TextBox } from './model';
import { isCJK } from './text';

/** 假字体：汉字 1 em，西文 0.5 em */
const fake: Measurer = {
  width: (ch) => (isCJK(ch) ? 1 : 0.5),
  metrics: () => ({ ascent: 0.88, descent: 0.12 }),
};

function box(wrap: TextBox['wrap'], over: Partial<TextBox> = {}): TextBox {
  return {
    w: 100,
    h: 60,
    posH: { rel: 'margin', offset: 0 },
    posV: { rel: 'paragraph', offset: 0 },
    wrap,
    dist: { t: 0, b: 0, l: 9, r: 9 },
    behind: false,
    fill: null,
    line: null,
    inset: { l: 0, t: 0, r: 0, b: 0 },
    vAlign: 'top',
    paragraphs: [],
    ...over,
  };
}

/** 段首锚定一个浮动对象、后面是一长段文字 */
function withFloat(b: TextBox): Doc {
  const d = newDocument();
  const p = newParagraph('国'.repeat(200));
  p.runs.unshift({ text: FIELD_CHAR, props: p.runs[0].props, box: b });
  d.body = [p];
  return d;
}

const lines = (d: Doc) => layoutDocument(d, fake).pages[0].body;

describe('浮动对象的文字环绕（wp:anchor 的 wrap）', () => {
  it('四周型：和浮动对象同高的行排在它旁边（加上环绕距离），下面的行恢复整栏', () => {
    const d = withFloat(box('square'));
    const ml = d.setup.marginLeft;
    const ls = lines(d);
    const beside = ls.filter((l) => l.y < d.setup.marginTop + 60 - 0.01);
    const below = ls.filter((l) => l.y >= d.setup.marginTop + 60);
    expect(beside.length).toBeGreaterThan(0);
    for (const l of beside) expect(l.x).toBeCloseTo(ml + 100 + 9);
    // 旁边的行按剩下的宽度断行
    const n = (l: (typeof ls)[number]) => l.line.end - l.line.start;
    expect(n(beside[1])).toBe(Math.floor((textWidth(d.setup) - 109 + 0.01) / 10.5));
    expect(below[0].x).toBeCloseTo(ml);
    expect(n(below[1])).toBe(Math.floor((textWidth(d.setup) + 0.01) / 10.5));
    // 文字前后接续，没有丢字
    for (let i = 1; i < ls.length; i++) expect(ls[i].line.start).toBe(ls[i - 1].line.end);
  });

  it('上下型：浮动对象所在的高度不放文字，文字从它下面开始', () => {
    const d = withFloat(box('topAndBottom'));
    const ls = lines(d);
    expect(ls[0].y).toBeCloseTo(d.setup.marginTop + 60);
    expect(ls[0].x).toBeCloseTo(d.setup.marginLeft);
  });

  it('浮于文字上方 / 衬于文字下方不影响文字', () => {
    for (const b of [box('none'), box('square', { behind: true })]) {
      const ls = lines(withFloat(b));
      const d = withFloat(b);
      expect(ls[0].x).toBeCloseTo(d.setup.marginLeft);
      expect(ls[0].y).toBeCloseTo(d.setup.marginTop);
    }
  });

  it('排版时算出的位置就是画浮动对象的位置', () => {
    const d = withFloat(box('square', { posH: { rel: 'page', offset: 300 }, posV: { rel: 'page', offset: 200 } }));
    const f = layoutDocument(d, fake).pages[0].floats[0];
    expect(f.x).toBeCloseTo(300);
    expect(f.y).toBeCloseTo(200);
  });
});

describe('浮动表格（w:tblpPr）', () => {
  function docWithFloatTable(): { d: Doc; after: Paragraph } {
    const d = newDocument();
    const t = newTable(2, 2, 200, d.defaults);
    t.float = { hAnchor: 'margin', vAnchor: 'page', x: 0, y: 300, dist: { l: 9, r: 9, t: 0, b: 0 } };
    const after = newParagraph('国'.repeat(600));
    d.body = [newParagraph('前'), t, after];
    return { d, after };
  }

  it('表格放在指定位置、不占正文高度，后面的文字绕开它', () => {
    const { d } = docWithFloatTable();
    const ls = lines(d);
    const cells = ls.filter((l) => l.path.length > 1);
    const text = ls.filter((l) => l.path.length === 1 && l.path[0] === 2);
    expect(Math.min(...cells.map((l) => l.y))).toBeCloseTo(300 + 0.5); // 上框线 0.5 磅之下
    // 下一段紧接上一段，不被表格推下去
    expect(text[0].y).toBeCloseTo(d.setup.marginTop + 15.6);
    // 与表格同高的行在表格右边
    const beside = text.filter((l) => l.y + l.line.height > 300 + 0.01 && l.y < 300 + 20);
    expect(beside.length).toBeGreaterThan(0);
    for (const l of beside) expect(l.x).toBeCloseTo(d.setup.marginLeft + 200 + 9);
  });
});

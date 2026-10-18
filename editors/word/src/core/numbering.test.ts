// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { layoutDocument, type Measurer } from './layout';
import { newDocument, newParagraph, paraText, type Doc } from './model';
import { chineseNumber, computeLabels, ensureList, formatNumber } from './numbering';
import { applyNumbering, backspace, changeNumLevel, insertSectionBreak } from './ops';
import { readDocx, writeDocx } from '../ooxml/docx';
import { isCJK } from './text';

const fake: Measurer = {
  width: (ch) => (isCJK(ch) ? 1 : 0.5),
  metrics: () => ({ ascent: 0.88, descent: 0.12 }),
};
const at = (i: number, offset = 0) => ({ story: 'body' as const, path: [i], offset });

describe('编号格式', () => {
  it('中文计数', () => {
    expect([1, 10, 11, 20, 21, 100, 105, 110].map(chineseNumber)).toEqual(['一', '十', '十一', '二十', '二十一', '一百', '一百零五', '一百一十']);
    expect(formatNumber(4, 'upperRoman')).toBe('IV');
    expect(formatNumber(28, 'lowerLetter')).toBe('bb');
    expect(formatNumber(3, 'decimalEnclosedCircle')).toBe('③');
  });

  it('公文多级编号：一、（一）1.（1），下级随上级重新开始', () => {
    const d = newDocument();
    d.body = ['总则', '目的', '范围', '职责', '采购部'].map((t) => newParagraph(t));
    applyNumbering(d, at(0), at(4), 'gongwen');
    changeNumLevel(d, at(1), at(2), 1);
    changeNumLevel(d, at(4), at(4), 1);
    const labels = computeLabels(d);
    expect(d.body.map((b) => labels.get(b as never))).toEqual(['一、', '（一）', '（二）', '二、', '（一）']);
  });

  it('编号文字参与排版，悬挂缩进的正文对齐到缩进位置', () => {
    const d = newDocument();
    d.body = [newParagraph('国'.repeat(50))];
    applyNumbering(d, at(0), at(0), 'decimal');
    const lines = layoutDocument(d, fake).pages[0].body;
    const [l0, l1] = lines.map((l) => l.line);
    expect(l0.glyphs[0].ch).toBe('1');
    // 首行编号在 0，正文从左缩进 21 磅开始；第二行也从 21 磅开始
    expect(l0.glyphs[0].x).toBeCloseTo(0);
    expect(l0.glyphs.find((g) => g.ch === '国')!.x).toBeCloseTo(21);
    expect(l1.glyphs[0].x).toBeCloseTo(21);
  });

  it('段首 Backspace 去掉编号', () => {
    const d = newDocument();
    d.body = [newParagraph('甲')];
    applyNumbering(d, at(0), at(0), 'bullet');
    backspace(d, at(0));
    expect(d.body[0].kind === 'p' && d.body[0].props.num).toBeNull();
  });
});

describe('首页不同 / 奇偶页不同', () => {
  it('各页使用对应的页眉', () => {
    const d = newDocument();
    d.setup.titlePg = true;
    d.setup.evenAndOdd = true;
    d.header[0].runs = [{ text: '奇', props: d.defaults }];
    d.headerFirst[0].runs = [{ text: '首', props: d.defaults }];
    d.headerEven[0].runs = [{ text: '偶', props: d.defaults }];
    d.body = [newParagraph('一\f二\f三\f四')];
    const pages = layoutDocument(d, fake).pages;
    expect(pages.map((p) => p.header.map((l) => l.line.glyphs.map((g) => g.ch).join('')).join(''))).toEqual(['首', '偶', '奇', '偶']);
    expect(pages[1].header[0].story).toBe('headerEven');
  });
});

describe('分节符', () => {
  const doc = (type: 'nextPage' | 'continuous' | 'oddPage') => {
    const d = newDocument();
    d.body = [newParagraph('甲乙')];
    insertSectionBreak(d, at(0, 1), type);
    return d;
  };
  it('下一页分节符换页，连续分节符不换页', () => {
    const d = doc('nextPage');
    expect(d.body.map((b) => paraText(b as never))).toEqual(['甲', '乙']);
    expect(layoutDocument(d, fake).pages).toHaveLength(2);
    expect(layoutDocument(doc('continuous'), fake).pages).toHaveLength(1);
  });
  it('奇数页分节符：从第 1 页到第 3 页，中间插入空白页', () => {
    const pages = layoutDocument(doc('oddPage'), fake).pages;
    expect(pages).toHaveLength(3);
    expect(pages[1].body).toHaveLength(0);
    expect(pages[0].body.at(-1)!.sectionBreak).toBe('oddPage');
  });
});

describe('docx 读写', () => {
  it('编号、分节、首页/偶数页页眉页脚往返一致', () => {
    const d: Doc = newDocument();
    d.setup.titlePg = true;
    d.setup.evenAndOdd = true;
    d.headerFirst[0].runs = [{ text: '首页页眉', props: d.defaults }];
    d.footerEven[0].runs = [{ text: '偶数页页脚', props: d.defaults }];
    d.body = [newParagraph('总则'), newParagraph('目的'), newParagraph('正文')];
    applyNumbering(d, at(0), at(1), 'gongwen');
    changeNumLevel(d, at(1), at(1), 1);
    insertSectionBreak(d, at(2, 2), 'continuous');
    ensureList(d, 'bullet');

    const back = readDocx(writeDocx(d));
    expect(back.lossy).toEqual([]);
    expect(back.setup.titlePg).toBe(true);
    expect(back.setup.evenAndOdd).toBe(true);
    expect(paraText(back.headerFirst[0])).toBe('首页页眉');
    expect(paraText(back.footerEven[0])).toBe('偶数页页脚');
    const labels = computeLabels(back);
    expect(back.body.slice(0, 2).map((b) => labels.get(b as never))).toEqual(['一、', '（一）']);
    expect(back.body.map((b) => (b.kind === 'p' ? b.props.sectionBreak : null))).toEqual([null, null, 'continuous', null]);
    const lvl = back.lists.find((l) => l.id === (back.body[0] as never as { props: { num: { id: number } } }).props.num.id)!.levels[0];
    expect(lvl).toMatchObject({ fmt: 'chineseCounting', text: '%1、', suffix: 'nothing' });
  });
});

describe('编号自己的字符格式', () => {
  it('编号级别写了加粗 / 字号时，编号按它排版，正文不变', () => {
    const d = newDocument();
    d.body = [newParagraph('条款')];
    applyNumbering(d, at(0), at(0), 'gongwen');
    const glyphs = () => layoutDocument(d, fake).pages[0].body[0].line.glyphs;
    const plain = glyphs();
    expect(plain[0].props.bold).toBe(false);
    d.lists[0].levels[0].label = { bold: true, size: 21 };
    const bold = glyphs();
    expect(bold[0].props).toMatchObject({ bold: true, size: 21 });
    expect(bold[0].w).toBeCloseTo(plain[0].w * 2);
    // 正文的字不受影响
    expect(bold[bold.length - 1].props).toMatchObject({ bold: false, size: 10.5 });
  });
});

import { describe, expect, it } from 'vitest';
import { layoutDocument, caretRect, hitTest, rowEnds, selectionRects, type Measurer } from './layout';
import { FIELD_CHAR, newDocument, newHFParagraph, newParagraph, newTable, textWidth, type Doc, type Paragraph, type TextBox } from './model';
import { isCJK } from './text';
import { linePixels, snapPixel } from './render';

/** 假字体：汉字与全角标点 1 em，西文 0.5 em，上伸 0.88 下伸 0.12。 */
const fake: Measurer = {
  width: (ch) => (isCJK(ch) ? 1 : 0.5),
  metrics: () => ({ ascent: 0.88, descent: 0.12 }),
};

const han = (n: number) => '国'.repeat(n);

function docWith(...paras: Paragraph[]): Doc {
  const d = newDocument();
  d.body = paras;
  return d;
}

function bodyLines(d: Doc) {
  return layoutDocument(d, fake).pages.map((p) => p.body.map((l) => l.line));
}

describe('中文默认页面', () => {
  it('A4 五号字每行 39 字（与 Word 中文默认一致）', () => {
    const d = docWith(newParagraph(han(100)));
    const lines = bodyLines(d)[0];
    expect(lines.map((l) => l.end - l.start)).toEqual([39, 39, 22]);
    expect(textWidth(d.setup)).toBeCloseTo(415.3, 1);
  });

  it('行网格：每页 44 行，行高 15.6 磅', () => {
    const d = docWith(...Array.from({ length: 100 }, () => newParagraph('国')));
    const pages = bodyLines(d);
    expect(pages.map((p) => p.length)).toEqual([44, 44, 12]);
    expect(pages[0][0].height).toBeCloseTo(15.6);
  });

  it('1.5 倍行距 = 1.5 个网格行；大字号占两格', () => {
    const p = newParagraph('国', { line: { rule: 'auto', value: 1.5 } });
    expect(bodyLines(docWith(p))[0][0].height).toBeCloseTo(23.4);
    const big = newParagraph('标题', undefined, { ...newDocument().defaults, size: 22 });
    expect(bodyLines(docWith(big))[0][0].height).toBeCloseTo(31.2);
  });

  it('Word 实测：行网格下多倍行距 = max(倍数 × 网格行距, 取整到网格的自然高度)，不是两者相乘', () => {
    const run = { ...newDocument().defaults, size: 22 }; // 自然高度 22 磅 → 两格 31.2
    const h = (v: number) => bodyLines(docWith(newParagraph('标题', { line: { rule: 'auto', value: v } }, run)))[0][0].height;
    expect(h(1.5)).toBeCloseTo(31.2);
    expect(h(2.408)).toBeCloseTo(2.408 * 15.6);
    expect(h(3)).toBeCloseTo(46.8);
  });

  it('固定值行距不受网格影响', () => {
    const p = newParagraph('国', { line: { rule: 'exact', value: 28 } });
    expect(bodyLines(docWith(p))[0][0].height).toBe(28);
  });
});

describe('断行规则', () => {
  it('行末逗号溢出边界（标点溢出）', () => {
    const lines = bodyLines(docWith(newParagraph(han(39) + '，' + han(5))))[0];
    expect(lines[0].end - lines[0].start).toBe(40);
  });

  it('左对齐 / 居中 / 右对齐的段落标点不溢出：逗号放不下时连同前一个字换行（与 Word 一致）', () => {
    for (const align of ['left', 'center', 'right'] as const) {
      const lines = bodyLines(docWith(newParagraph(han(39) + '，' + han(5), { align })))[0];
      expect(lines[0].end - lines[0].start).toBe(38);
      expect(lines[1].glyphs.slice(0, 2).map((g) => g.ch).join('')).toBe('国，');
    }
    // 放得下的逗号仍留在行尾
    const fits = bodyLines(docWith(newParagraph(han(38) + '，' + han(5), { align: 'left' })))[0];
    expect(fits[0].end - fits[0].start).toBe(39);
  });

  it('行首禁则：右括号不能出现在行首，前一个字一起换行', () => {
    const lines = bodyLines(docWith(newParagraph(han(39) + '」' + han(5))))[0];
    expect(lines[0].end - lines[0].start).toBe(38);
  });

  it('行尾禁则：左括号不能留在行尾', () => {
    const lines = bodyLines(docWith(newParagraph(han(38) + '《' + han(5))))[0];
    expect(lines[0].end - lines[0].start).toBe(38);
  });

  it('西文单词不拆开，在空格处换行', () => {
    const words = Array.from({ length: 40 }, () => 'liner').join(' ');
    const d = docWith(newParagraph(words));
    const text = words;
    for (const l of bodyLines(d)[0].slice(0, -1)) {
      expect(text[l.end - 1]).toBe(' ');
      expect(/[a-z]/.test(text[l.end])).toBe(true);
    }
  });

  it('中西文之间自动加 1/4 字宽间距', () => {
    const l = bodyLines(docWith(newParagraph('中A', { align: 'left' })))[0][0];
    expect(l.xs[1] - l.xs[0]).toBeCloseTo(10.5);
    expect(l.glyphs[1].x - l.glyphs[0].x).toBeCloseTo(10.5 + 10.5 * 0.25);
  });

  it('Word 实测（打印成 PDF：五号 2.64、小四 2.98 磅）：西文→中文的自动间距也是 1/4 字宽；有空格时不加', () => {
    const l = bodyLines(docWith(newParagraph('中A中 A 中', { align: 'left' })))[0][0];
    const x = l.glyphs.map((g) => g.x);
    expect(x[2] - x[1]).toBeCloseTo(5.25 + 10.5 * 0.25);
    expect(x[3] - x[2]).toBeCloseTo(10.5 + 5.25); // 中、空格，没有自动间距
    expect(x[4] - x[3]).toBeCloseTo(5.25 + 5.25);
  });

  it('Word 实测：压缩标点时两端对齐的行可多放一个字，标点合计最多压掉半个字宽；不压缩标点的文档不压', () => {
    // 五号字版心 415.3 = 39.55 个字
    const first = (text: string, compress: boolean) => {
      const d = docWith(newParagraph(text + han(20), { align: 'both' }));
      d.setup.compressPunct = compress;
      return bodyLines(d)[0][0].end;
    };
    const one = han(19) + '，' + han(20); // 第 20 个字是逗号，前 40 个字超出 4.7 磅
    expect(first(one, true)).toBe(40);
    expect(first(one, false)).toBe(39);
    // 前 41 个字里有 8 个逗号，超出 15.2 磅：超过上限（见“挤进 / 推出”），仍是 40 个字
    const many = Array.from({ length: 41 }, (_, i) => (i % 5 === 4 && i < 40 ? '，' : '国')).join('');
    expect(first(many, true)).toBe(40);
  });

  it('Word 实测：两端对齐时空格和中西文间距最多压掉 1/4 来多放一点，左对齐不压', () => {
    const d0 = newDocument();
    // 可用宽度 104：九个汉字 94.5 + 空格 5.25 + a 5.25 = 105，超出 1 磅（空格可压 1.31）
    const indentRight = textWidth(d0.setup) - 104;
    const both = bodyLines(docWith(newParagraph('国'.repeat(9) + ' a 国国国', { align: 'both', indentRight })))[0][0];
    const left = bodyLines(docWith(newParagraph('国'.repeat(9) + ' a 国国国', { align: 'left', indentRight })))[0][0];
    expect(both.end).toBe(12); // “a ” 留在第一行
    expect(both.right).toBeLessThanOrEqual(104 + 0.01);
    expect(left.end).toBe(10); // “a” 换到下一行
  });

  it('Word 实测：全角标点和字母数字之间不加自动间距（：1、（A）', () => {
    const x = bodyLines(docWith(newParagraph('中：1、中（A）', { align: 'left' })))[0][0].glyphs.map((g) => g.x);
    expect(x[2] - x[1]).toBeCloseTo(10.5); // ：→1
    expect(x[3] - x[2]).toBeCloseTo(5.25); // 1→、
    expect(x[7] - x[6]).toBeCloseTo(5.25); // A→）
    // 全角空格与数字之间也不加（33.docx：“　　2、”）
    const y = bodyLines(docWith(newParagraph('　　2、', { align: 'left' })))[0][0].glyphs;
    expect(y.find((g) => g.ch === '2')!.x).toBeCloseTo(21);
  });

  it('Word 实测：相邻标点挤压——结束类标点后面紧跟全角标点时只占半宽，两个开始类标点相连时后一个只占半宽，！？不挤压', () => {
    const xs = (t: string) => bodyLines(docWith(newParagraph(t, { align: 'left' })))[0][0].glyphs.map((g) => g.x);
    expect(xs('中）：中').map((v) => +v.toFixed(2))).toEqual([0, 10.5, 15.75, 26.25]);
    expect(xs('中。（中').map((v) => +v.toFixed(2))).toEqual([0, 10.5, 15.75, 26.25]);
    // 后一个开始括号去掉左边的空白：字形往左画半个字
    expect(xs('中（（中').map((v) => +v.toFixed(2))).toEqual([0, 10.5, 15.75, 26.25]);
    expect(xs('中！？中').map((v) => +v.toFixed(2))).toEqual([0, 10.5, 21, 31.5]);
  });

  describe('挤进 / 推出（行末的字靠压缩标点放进本行，还是换到下一行）', () => {
    // N 个字（第 3、15 个是逗号）排在 N·10.5 - o 宽的行里：第 N 个字超出 o 磅
    const fits = (n: number, o: number, opts: { compat?: number; after?: string; align?: 'both' | 'left' } = {}) => {
      const d0 = newDocument();
      const chars = Array.from({ length: n }, (_, i) => (i === 3 || i === 15 ? '，' : '国')).join('');
      const d = docWith(newParagraph(chars + (opts.after ?? '') + han(20), { align: opts.align ?? 'both', indentRight: textWidth(d0.setup) - (n * 10.5 - o) }));
      d.setup.compressPunct = true;
      if (opts.compat) d.setup.compatMode = opts.compat;
      return bodyLines(d)[0][0].end >= n;
    };
    it('兼容模式 15：上限 = 字宽 × 33K / (W/字宽 + 33K)（30 字、2 个逗号时 7.29 磅），字数越多上限越小', () => {
      expect(fits(30, 7.1)).toBe(true);
      expect(fits(30, 7.5)).toBe(false);
      expect(fits(49, 5.9)).toBe(true); // 10.5 × 66 / 114 = 6.08
      expect(fits(49, 6.3)).toBe(false);
    });
    it('兼容模式 15：该字后面紧跟要悬挂的行尾标点时更倾向推出（a ≈ 21）', () => {
      expect(fits(30, 6.0, { after: '。' })).toBe(true); // 10.5 × 42 / 71 = 6.21
      expect(fits(30, 6.5, { after: '。' })).toBe(false);
    });
    it('兼容模式 15：左对齐不压', () => {
      expect(fits(30, 1, { align: 'left' })).toBe(false);
    });
    it('兼容模式 11–14：任何对齐都压，但行内标点合计最多一个半宽', () => {
      for (const align of ['both', 'left'] as const) {
        expect(fits(30, 5.0, { compat: 11, align })).toBe(true);
        expect(fits(30, 5.5, { compat: 11, align })).toBe(false);
      }
    });
  });

  it('兼容模式 11 + useAltKinsoku：行尾逗号不悬挂；剩余 ≥ 半字宽时压成半宽留在本行（行内标点照样不压）', () => {
    const run = (head: string, rem: number) => {
      const d0 = newDocument();
      const d = docWith(newParagraph(head + '，' + han(20), { align: 'left', indentRight: textWidth(d0.setup) - (20 * 10.5 + rem) }));
      Object.assign(d.setup, { compressPunct: true, compatMode: 11, altKinsoku: true });
      return bodyLines(d)[0][0];
    };
    expect(run(han(20), 5.5).end).toBe(21);
    expect(run(han(20), 4.5).end).toBe(19); // 不悬挂：“国，”一起换行
    // 行尾逗号先去掉自己的半宽，够了就不再压行内的标点
    const mid = run(han(5) + '，' + han(14), 5.5);
    expect(mid.end).toBe(21);
    expect(mid.glyphs[6].x - mid.glyphs[5].x).toBeCloseTo(10.5);
  });

  it('Word 实测：字符间距（w:spacing）对汉字和全角标点按两倍生效，西文按一倍', () => {
    const p = newParagraph('国国ab国');
    p.runs[0].props = { ...p.runs[0].props, spacing: -0.5 };
    const x = bodyLines(docWith(p))[0][0].glyphs.map((g) => g.x);
    expect(x[1] - x[0]).toBeCloseTo(9.5);
    expect(x[3] - x[2]).toBeCloseTo(4.75);
  });

  it('Word 实测：空格不撑高行（20 磅的空格不改变五号字行的行高）', () => {
    const d = newDocument();
    const p = newParagraph('国');
    p.runs.push({ text: '    ', props: { ...d.defaults, size: 20 } }, { text: '国', props: { ...d.defaults } });
    const q = newParagraph('国');
    q.runs.push({ text: 'XX', props: { ...d.defaults, size: 20 } });
    const [a, b] = bodyLines(docWith(p, q))[0];
    expect(a.height).toBeCloseTo(15.6);
    expect(b.height).toBeCloseTo(31.2);
  });

  it('两端对齐：以手动换行符结束的行默认拉开；有 doNotExpandShiftReturn 时不拉开（33.docx）', () => {
    const run = (noExpand: boolean) => {
      const d = docWith(newParagraph('国国国\n国', { align: 'both' }));
      d.setup.noExpandShiftReturn = noExpand;
      return bodyLines(d)[0][0].glyphs.map((g) => g.x);
    };
    expect(run(true)).toEqual([0, 10.5, 21]);
    expect(run(false)[2]).toBeGreaterThan(100);
  });

  it('手动换行与分页符', () => {
    const pages = bodyLines(docWith(newParagraph('甲\n乙\f丙')));
    expect(pages[0].map((l) => l.brk)).toEqual(['line', 'page']);
    expect(pages[1][0].start).toBe(4);
  });
});

describe('对齐', () => {
  it('两端对齐：除末行外撑满整行，末行不拉伸', () => {
    const d = docWith(newParagraph(han(38) + 'ab ' + han(20)));
    const [l1, l2] = bodyLines(d)[0];
    expect(l1.right).toBeCloseTo(textWidth(d.setup), 1);
    expect(l2.right).toBeLessThan(textWidth(d.setup) - 10);
  });

  it('居中与右对齐', () => {
    const w = textWidth(newDocument().setup);
    const c = bodyLines(docWith(newParagraph('国国', { align: 'center' })))[0][0];
    expect(c.xs[0]).toBeCloseTo((w - 21) / 2);
    const r = bodyLines(docWith(newParagraph('国国', { align: 'right' })))[0][0];
    expect(r.right).toBeCloseTo(w);
  });

  it('首行缩进 2 字符', () => {
    const l = bodyLines(docWith(newParagraph(han(50), { firstLineChars: 2 })))[0];
    expect(l[0].xs[0]).toBeCloseTo(21);
    expect(l[0].end - l[0].start).toBe(37);
    expect(l[1].xs[0]).toBe(0);
  });
});

describe('字符网格', () => {
  it('指定每行 40 字（网格窄于字号）时每个汉字仍占一格', () => {
    const d = docWith(newParagraph(han(90)));
    d.setup.grid = { type: 'linesAndChars', linePitch: 15.6, charSpace: Math.round((textWidth(d.setup) / 40 - 10.5) * 4096) };
    const lines = bodyLines(d)[0];
    expect(lines[0].end - lines[0].start).toBe(40);
  });

  it('公文网格（每行 28 字）：三号正文每字一格；二号标题每字 = 字号 + 同样的字距差（不占两格）', () => {
    const d = newDocument();
    d.defaults = { ...d.defaults, size: 16 };
    const tw = textWidth(d.setup);
    d.setup.grid = { type: 'linesAndChars', linePitch: 28.95, charSpace: Math.round((tw / 28 - 16) * 4096) };
    d.body = [newParagraph(han(60), undefined, { ...d.defaults, size: 16 }), newParagraph(han(30), undefined, { ...d.defaults, size: 22 })];
    const lines = bodyLines(d)[0];
    const [p1] = lines;
    const p2 = lines.find((l, i) => i > 0 && l.glyphs[0]?.props.size === 22)!;
    expect(p1.end - p1.start).toBe(28);
    // 每字 22 + (网格 − 16) 磅
    const delta = d.setup.grid.charSpace / 4096;
    expect(p2.end - p2.start).toBe(Math.floor((tw + 0.01) / (22 + delta)));
  });

  it('Word 实测：linesAndChars 网格下每个字符（中西文、任何字号）都加 charSpace/4096 磅', () => {
    // 排版对照样例：版心 442.25 磅、charSpace = -844（-0.206 磅）；Word 里三号字每字 15.79、二号 21.79、小四 11.79
    const d = newDocument();
    d.defaults = { ...d.defaults, size: 16 };
    d.setup.grid = { type: 'linesAndChars', linePitch: 29, charSpace: -844 };
    d.body = [newParagraph('国国', undefined, { ...d.defaults, size: 16 }), newParagraph('国国', undefined, { ...d.defaults, size: 22 }), newParagraph('国国', undefined, { ...d.defaults, size: 12 }), newParagraph('AA', undefined, { ...d.defaults, size: 22 })];
    const adv = bodyLines(d)[0].map((l) => l.xs[1] - l.xs[0]);
    expect(adv[0]).toBeCloseTo(15.794, 2);
    expect(adv[1]).toBeCloseTo(21.794, 2);
    expect(adv[2]).toBeCloseTo(11.794, 2);
    expect(adv[3]).toBeCloseTo(11 - 0.206, 2); // 西文（假测量器 0.5em）同样减
  });
});

describe('分页', () => {
  it('Word 实测：网格行在页底只要文字底部（行距与文字高的平均）在版心内就放得下', () => {
    // 行距 29、三号字（假测量器文字高 16）：页底剩 (29 + 16) / 2 = 22.5 磅就能再放一行
    const pages = (rest: number) => {
      const d = newDocument();
      d.defaults = { ...d.defaults, size: 16 };
      d.setup.grid = { type: 'linesAndChars', linePitch: 29, charSpace: 0 };
      d.setup.height = d.setup.marginTop + d.setup.marginBottom + 21 * 29 + rest;
      d.body = Array.from({ length: 30 }, () => newParagraph('国', undefined, { ...d.defaults, size: 16 }));
      return bodyLines(d)[0].length;
    };
    expect(pages(28.8)).toBe(22);
    expect(pages(22.6)).toBe(22);
    expect(pages(22.4)).toBe(21);
  });

  it('孤行控制：段首一行不单独留在页底', () => {
    const filler = Array.from({ length: 43 }, () => newParagraph('国'));
    const tail = newParagraph(han(100), { widowControl: true });
    const pages = bodyLines(docWith(...filler, tail));
    expect(pages[0]).toHaveLength(43);
    expect(pages[1]).toHaveLength(3);
  });

  it('不开孤行控制时段落在页底拆开（Word 中文默认）', () => {
    const filler = Array.from({ length: 43 }, () => newParagraph('国'));
    const pages = bodyLines(docWith(...filler, newParagraph(han(100))));
    expect(pages[0]).toHaveLength(44);
  });

  it('段前分页', () => {
    const pages = bodyLines(docWith(newParagraph('一'), newParagraph('二', { pageBreakBefore: true })));
    expect(pages).toHaveLength(2);
  });
});

describe('段距', () => {
  const ys = (d: Doc) => layoutDocument(d, fake).pages[0].body.map((l) => l.y);
  const noGrid = (...ps: Paragraph[]) => {
    const d = docWith(...ps);
    d.setup.grid = { type: 'none', linePitch: 15.6, charSpace: 0 };
    return d;
  };
  const exact = { line: { rule: 'exact' as const, value: 15 } };

  it('Word 实测：上一段段后与下一段段前相邻时取较大的一个，不相加', () => {
    const y = ys(noGrid(newParagraph('甲', { ...exact, spaceAfter: 5 }), newParagraph('乙', { ...exact, spaceBefore: 12 }), newParagraph('丙', { ...exact, spaceBefore: 2 })));
    expect(y[1] - y[0]).toBeCloseTo(15 + 12);
    expect(y[2] - y[1]).toBeCloseTo(15 + 2);
  });

  it('表格单元格里的段落同样取较大的一个', () => {
    const d = noGrid(newParagraph(''));
    const t = newTable(1, 1, 200, d.defaults);
    t.rows[0].cells[0].paragraphs = [newParagraph('甲', { ...exact, spaceAfter: 8 }), newParagraph('乙', { ...exact, spaceBefore: 8 })];
    d.body = [t, newParagraph('')];
    const [a, b] = layoutDocument(d, fake).pages[0].body.filter((l) => l.path.length > 1).map((l) => l.y);
    expect(b - a).toBeCloseTo(15 + 8);
  });

  it('Word 实测：段前“自动”是 14 磅（同样与上一段段后取大），页面开头的第一段不加', () => {
    const y = ys(noGrid(newParagraph('甲', { ...exact, spaceBeforeAuto: true }), newParagraph('乙', { ...exact, spaceAfter: 6 }), newParagraph('丙', { ...exact, spaceBeforeAuto: true })));
    expect(y[0]).toBeCloseTo(newDocument().setup.marginTop);
    expect(y[2] - y[1]).toBeCloseTo(15 + 14);
  });
});

describe('表格行尾标记', () => {
  it('每行一个：在这一行最右边的格线外，基线与该行第一行文字相同，颜色取那一行的文字', () => {
    const d = docWith(newParagraph('前'));
    const t = newTable(2, 2, 200, d.defaults);
    t.rows[1].cells[0].paragraphs = [newParagraph('第一行\n第二行', {}, { ...d.defaults, color: 'FF0000' })];
    d.body = [newParagraph('前'), t, newParagraph('后')];
    const page = layoutDocument(d, fake).pages[0];
    const ends = rowEnds(page);
    expect(ends.map((e) => [e.bi, e.r])).toEqual([[1, 0], [1, 1]]);
    const right = Math.max(...page.cells!.map((c) => c.x + c.w));
    for (const e of ends) expect(e.x).toBeCloseTo(right);
    const first = page.body.filter((l) => l.path[0] === 1 && l.path[1] === 1).sort((a, b) => a.y - b.y)[0];
    expect(ends[1].baseline).toBeCloseTo(first.y + first.line.baseline);
    expect(ends[1].line?.glyphs[0].props.color).toBe('FF0000');
  });
});

describe('页眉页脚占的位置', () => {
  const bottomOf = (d: Doc) => {
    const pg = layoutDocument(d, fake).pages[0];
    const last = pg.body[pg.body.length - 1];
    return last.y + last.line.height;
  };
  const longDoc = (setup: Partial<Doc['setup']>) => {
    const d = docWith(...Array.from({ length: 80 }, () => newParagraph('国', { line: { rule: 'exact', value: 10 } })));
    d.setup.grid = { type: 'none', linePitch: 15.6, charSpace: 0 };
    Object.assign(d.setup, setup);
    return d;
  };

  it('原文件没有页脚时正文排到下边距（页脚距边界比下边距大也不占位置）', () => {
    const d = longDoc({ marginBottom: 12, footerDistance: 50 });
    d.footer = [{ ...newHFParagraph(), hfNone: true }];
    expect(bottomOf(d)).toBeGreaterThan(d.setup.height - 12 - 10 - 0.01);
  });

  it('存在的页脚即使只有一个空段落也占位置', () => {
    const d = longDoc({ marginBottom: 12, footerDistance: 50 });
    d.footer = [newHFParagraph()];
    expect(bottomOf(d)).toBeLessThanOrEqual(d.setup.height - 50 + 0.01);
  });

  it('页脚里的图文框段落不占页脚高度', () => {
    const d = longDoc({ marginBottom: 12, footerDistance: 50 });
    d.footer = [newParagraph('', { frame: { hAnchor: 'margin', vAnchor: 'text', xAlign: 'center', x: 0, y: 0 } }), newHFParagraph()];
    const one = longDoc({ marginBottom: 12, footerDistance: 50 });
    one.footer = [newHFParagraph()];
    expect(bottomOf(d)).toBeCloseTo(bottomOf(one));
  });
});

describe('表格', () => {
  it('列宽、行高与跨页', () => {
    const d = newDocument();
    const t = newTable(3, 90, textWidth(d.setup), d.defaults);
    t.rows[0].height = 40;
    t.rows[0].heightRule = 'exact';
    d.body = [t, newParagraph('')];
    const L = layoutDocument(d, fake);
    const first = L.pages[0];
    // 第一行固定 40 磅；其余行按字体行高，不对齐文档的行网格（Word 实测：表格里不用行网格）
    const ys = [...new Set(first.body.map((l) => Math.round(l.y * 10) / 10))];
    expect(ys[1] - ys[0]).toBeCloseTo(40);
    expect(ys[2] - ys[1]).toBeLessThan(15.6);
    expect(L.pages.length).toBe(2);
    // 行不拆开：每页的行顶必须落在版心内
    for (const pg of L.pages) for (const l of pg.body) expect(l.y + l.line.height).toBeLessThanOrEqual(d.setup.height - d.setup.marginBottom + 0.01);
    // 单元格文字区左边 = 单元格左边 + 左边距
    const xs = [...new Set(first.body.filter((l) => l.path.length === 4).map((l) => Math.round(l.x * 10) / 10))];
    expect(xs).toHaveLength(3);
    expect(xs[1] - xs[0]).toBeCloseTo(textWidth(d.setup) / 3, 1);
  });
});

describe('页眉页脚与光标', () => {
  it('页码域按页显示', () => {
    const d = docWith(newParagraph('一\f二'));
    d.footer = [newParagraph('', { align: 'center' })];
    d.footer[0].runs = [{ text: '￼', props: d.defaults, field: 'PAGE' }];
    const L = layoutDocument(d, fake);
    expect(L.pages.map((p) => p.footer[0].line.glyphs[0].ch)).toEqual(['1', '2']);
  });

  it('点击定位与光标位置互逆', () => {
    const d = docWith(newParagraph(han(50)));
    const L = layoutDocument(d, fake);
    const c = caretRect(L, { story: 'body', path: [0], offset: 45 }, 0)!;
    const hit = hitTest(L, c.page, c.x + 1, c.y + 5, 'body')!;
    expect(hit.offset).toBe(45);
    const rects = selectionRects(L, { story: 'body', path: [0], offset: 10 }, { story: 'body', path: [0], offset: 45 }, 0);
    expect(rects).toHaveLength(2);
  });
});

describe('嵌入式图片', () => {
  const withImage = (w: number, h: number, props?: Parameters<typeof newParagraph>[1]) => {
    const p = newParagraph('国国', props);
    p.runs.push({ text: FIELD_CHAR, props: p.runs[0].props, image: { id: 'img-test', w, h } });
    return p;
  };

  it('图片占一个字符位置，底边在基线上并把行撑高', () => {
    const d = docWith(withImage(60, 40, { snapToGrid: false }));
    const line = bodyLines(d)[0][0];
    expect(line.end - line.start).toBe(3);
    const g = line.glyphs.find((x) => x.image)!;
    expect(g.image).toEqual({ id: 'img-test', w: 60, h: 40 });
    expect(g.w).toBe(60);
    // 五号字 10.5 磅：上伸换成图片高度 40，下伸 0.12 em
    expect(line.baseline).toBeCloseTo(40);
    expect(line.height).toBeCloseTo(40 + 10.5 * 0.12);
    // 图片排在两个字后面：光标越过图片要走过它的宽度
    expect(line.xs[3] - line.xs[2]).toBeCloseTo(60);
  });

  it('比版心宽的图片按比例缩小到一行宽', () => {
    const d = docWith(withImage(1000, 500, { snapToGrid: false }));
    const width = textWidth(d.setup);
    const g = bodyLines(d)[0].flatMap((l) => l.glyphs).find((x) => x.image)!;
    expect(g.image!.w).toBeCloseTo(width);
    expect(g.image!.h).toBeCloseTo(width / 2);
  });

  it('行网格下图片所在行占整数个网格行', () => {
    const line = bodyLines(docWith(withImage(60, 40)))[0][0];
    expect(line.height).toBeCloseTo(15.6 * 3);
  });
});

describe('表格边框', () => {
  it('边框线带上表格的边框粗细（缺省 0.5 磅）', () => {
    const d = newDocument();
    const t = newTable(1, 1, textWidth(d.setup), d.defaults);
    const thick = newTable(1, 1, textWidth(d.setup), d.defaults);
    thick.borderSize = 1.5;
    d.body = [t, newParagraph(''), thick];
    const ws = layoutDocument(d, fake).pages[0].borders.map((b) => b.w);
    expect(ws).toEqual([0.5, 0.5, 0.5, 0.5, 1.5, 1.5, 1.5, 1.5]);
  });
});

describe('文本框', () => {
  const box = (over: Partial<TextBox> = {}): TextBox => ({
    w: 200,
    h: 100,
    posH: { rel: 'column', offset: 100 },
    posV: { rel: 'paragraph', offset: 2 },
    wrap: 'none',
    behind: false,
    fill: 'FFFFFF',
    line: { color: '000000', w: 0.75 },
    inset: { l: 7.2, t: 3.6, r: 7.2, b: 3.6 },
    vAlign: 'top',
    paragraphs: [newParagraph('合同号', { snapToGrid: false }), newParagraph('船号', { snapToGrid: false })],
    ...over,
  });
  const withBox = (b: TextBox, text = '包装标识：') => {
    const p = newParagraph(text);
    p.runs.push({ text: FIELD_CHAR, props: p.runs[0].props, box: b });
    return p;
  };

  it('锚点不占宽度、不影响正文排版；文本框按栏和段落定位，文字在内边距里面', () => {
    const d = docWith(newParagraph('第一段'), withBox(box()));
    const page = layoutDocument(d, fake).pages[0];
    const plain = layoutDocument(docWith(newParagraph('第一段'), newParagraph('包装标识：')), fake).pages[0];
    expect(page.body.map((l) => [l.y, l.line.right])).toEqual(plain.body.map((l) => [l.y, l.line.right]));
    const anchorLine = page.body[1];
    expect(anchorLine.line.end - anchorLine.line.start).toBe(6);

    expect(page.floats).toHaveLength(1);
    const f = page.floats[0];
    expect(f.x).toBeCloseTo(d.setup.marginLeft + 100);
    expect(f.y).toBeCloseTo(anchorLine.y + 2);
    expect(f.lines.map((l) => l.line.glyphs.map((g) => g.ch).join(''))).toEqual(['合同号', '船号']);
    expect(f.lines[0].x).toBeCloseTo(f.x + 7.2);
    expect(f.lines[0].y).toBeCloseTo(f.y + 3.6);
    expect(f.lines[1].y).toBeCloseTo(f.lines[0].y + f.lines[0].line.height);
  });

  it('相对页面 / 页边距定位，对齐方式，垂直居中', () => {
    const d = docWith(
      withBox(box({ posH: { rel: 'page', offset: 30 }, posV: { rel: 'page', offset: 40 } })),
      withBox(box({ posH: { rel: 'margin', offset: 0, align: 'right' }, posV: { rel: 'margin', offset: 0, align: 'center' }, vAlign: 'center' })),
    );
    const s = d.setup;
    const [a, b] = layoutDocument(d, fake).pages[0].floats;
    expect([a.x, a.y]).toEqual([30, 40]);
    expect(b.x).toBeCloseTo(s.width - s.marginRight - 200);
    expect(b.y).toBeCloseTo(s.marginTop + (s.height - s.marginTop - s.marginBottom - 100) / 2);
    // 两行文字在框内垂直居中
    const textH = b.lines[0].line.height + b.lines[1].line.height;
    expect(b.lines[0].y).toBeCloseTo(b.y + 3.6 + (100 - 7.2 - textH) / 2);
  });

  it('表格单元格里的文本框相对单元格定位', () => {
    const d = newDocument();
    const t = newTable(2, 1, textWidth(d.setup), d.defaults);
    t.rows[0].cells[1].paragraphs = [withBox(box({ posH: { rel: 'column', offset: 10 } }), '格')];
    d.body = [t];
    const page = layoutDocument(d, fake).pages[0];
    const cellLine = page.body.find((l) => l.path.length === 4 && l.path[2] === 1)!;
    expect(page.floats[0].x).toBeCloseTo(cellLine.x + 10);
  });

  it('页眉里的文本框 / 形状：每页都放，按页眉段落定位，标记为页眉页脚层（不能在正文里选中）', () => {
    const d = docWith(newParagraph('正文'), newParagraph('第二页', { pageBreakBefore: true }));
    const line = box({ geom: 'line', paragraphs: [], h: 0, posH: { rel: 'margin', offset: 0 }, posV: { rel: 'paragraph', offset: 20 } });
    d.header = [withBox(line, '页眉')];
    const pages = layoutDocument(d, fake).pages;
    expect(pages).toHaveLength(2);
    for (const pg of pages) {
      expect(pg.floats).toHaveLength(1);
      const f = pg.floats[0];
      expect(f.hf).toBe(true);
      expect(f.x).toBeCloseTo(d.setup.marginLeft);
      expect(f.y).toBeCloseTo(pg.header[0].y + 20);
    }
  });
});

describe('线宽换算成像素', () => {
  it('细线至少 1 像素，零头不足半个像素时舍去', () => {
    // 0.5 磅的表格线、0.75 磅的文本框边框：在 1 倍和 2 倍缩放下都是 1 像素
    expect([0.5, 0.75].flatMap((w) => [1, 1.333, 2].map((s) => linePixels(w, s)))).toEqual([1, 1, 1, 1, 1, 1]);
    expect(linePixels(0.75, 3)).toBe(2);
    expect(linePixels(1.5, 2)).toBe(3);
    expect(linePixels(2.25, 1)).toBe(2);
    // 1 磅在 125% 缩放（1.667 像素 / 磅）下与 0.5 磅一样画 1 像素（Word 屏幕显示）
    expect(linePixels(1, 1.667)).toBe(1);
    expect(linePixels(1, 2)).toBe(2);
  });
});

describe('下划线', () => {
  it('带下划线的空格也画下划线（填空用的横线），行尾的空格不画', () => {
    const d = newDocument();
    const p = newParagraph('持续', { snapToGrid: false });
    const u = { ...d.defaults, underline: true };
    p.runs.push({ text: '   ', props: u }, { text: '周', props: d.defaults }, { text: '  ', props: u });
    const line = bodyLines(docWith(p))[0][0];
    const under = line.glyphs.filter((g) => g.props.underline);
    expect(under.map((g) => g.ch)).toEqual([' ', ' ', ' ']);
    // 三个空格连成一段：紧接在“持续”后面，每个 0.5 em
    expect(under.map((g) => g.x - line.xs[0])).toEqual([21, 26.25, 31.5]);
    expect(under.every((g) => g.w === 5.25)).toBe(true);
  });
});

describe('表格的公共边', () => {
  it('相邻单元格的公共竖线坐标完全相同（不会因浮点误差画成两条）', () => {
    const d = newDocument();
    const t = newTable(3, 2, textWidth(d.setup), d.defaults);
    // 列宽取带小数的值，累加时容易出现浮点误差
    t.cols = [70.85, 99.2, 42.55];
    d.body = [t];
    const xs = layoutDocument(d, fake).pages[0].borders.filter((b) => b.x1 === b.x2).map((b) => b.x1);
    // 4 条竖线位置 × 2 行，中间两条各被左右两个单元格画一次：不同的坐标值只有 4 个
    expect(new Set(xs).size).toBe(4);
  });

  it('取整到像素时抹掉浮点误差', () => {
    expect(snapPixel(472.4999999)).toBe(snapPixel(472.5000001));
    expect(snapPixel(472.4)).toBe(472);
    expect(snapPixel(472.6)).toBe(473);
  });
});

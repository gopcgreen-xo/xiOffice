import { afterEach, describe, expect, it } from 'vitest';
import { layoutDocument, type Measurer } from './layout';
import { FIELD_CHAR, newDocument, newParagraph, type Run } from './model';
import { setShowCode } from './fieldView';
import { isCJK } from './text';

/** 假字体：汉字 1 em，西文 0.5 em */
const fake: Measurer = { width: (ch) => (isCJK(ch) ? 1 : 0.5), metrics: () => ({ ascent: 0.88, descent: 0.12 }) };

function lineOf(runs: (doc: ReturnType<typeof newDocument>) => Run[]) {
  const d = newDocument();
  const p = newParagraph('');
  p.runs = runs(d);
  d.body = [p];
  return layoutDocument(d, fake).pages[0].body[0].line;
}

describe('域的排版', () => {
  afterEach(() => setShowCode(['f1'], false));
  const runs = (d: ReturnType<typeof newDocument>): Run[] => {
    const pr = { ...d.defaults };
    return [
      { text: '表', props: pr },
      { text: FIELD_CHAR, props: pr, fc: { kind: 'begin', id: 'f1', instr: 'SEQ 表' } },
      { text: '12', props: pr },
      { text: FIELD_CHAR, props: pr, fc: { kind: 'end', id: 'f1' } },
      { text: '页', props: pr },
    ];
  };

  it('开始 / 结束标记不占宽度，结果照常排', () => {
    const line = lineOf(runs);
    const size = newDocument().defaults.size;
    const shown = line.glyphs.filter((g) => g.w > 0).map((g) => g.ch);
    expect(shown).toEqual(['表', '1', '2', '页']);
    // 每个字符位置都有光标 x：表 | 开始 | 1 2 | 结束 | 页
    expect(line.xs).toHaveLength(line.end - line.start + 1);
    expect(line.xs[2] - line.xs[1]).toBe(0);
    expect(line.xs[6] - line.xs[0]).toBeCloseTo(size * 3);
  });

  it('显示域代码：开始标记显示 { 域代码 }，结果和结束标记宽度为 0（光标位置还在）', () => {
    setShowCode(['f1'], true);
    const line = lineOf(runs);
    const size = newDocument().defaults.size;
    expect(line.glyphs.filter((g) => g.w > 0).map((g) => g.ch)).toEqual(['表', '{ SEQ 表 }', '页']);
    const code = line.glyphs.find((g) => g.ch.startsWith('{'))!;
    // “{ SEQ ” 6 个西文 + “表” + “ }” 2 个西文
    expect(code.w).toBeCloseTo(size * (8 * 0.5 + 1));
    expect(line.xs).toHaveLength(line.end - line.start + 1);
    expect(line.xs[5]).toBeCloseTo(line.xs[2]);
  });
});

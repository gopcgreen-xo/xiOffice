import { describe, expect, it } from 'vitest';
import { newDocument, newParagraph, paraText, type Doc } from './model';
import {
  applyAiText,
  applyRunProps,
  backspace,
  deleteRange,
  deleteRow,
  getPara,
  History,
  insertColumn,
  insertRaw,
  insertRow,
  insertTable,
  insertText,
  selectedText,
  splitParagraph,
  wordCount,
} from './ops';
import type { Pos } from './layout';

const at = (path: number[], offset: number): Pos => ({ story: 'body', path, offset });
const texts = (d: Doc) => d.body.map((b) => (b.kind === 'p' ? paraText(b) : '[表格]'));

describe('编辑操作', () => {
  it('输入、拆段、退格合并', () => {
    const d = newDocument();
    let p = insertText(d, at([0], 0), '办公用品');
    p = splitParagraph(d, at([0], 2));
    expect(texts(d)).toEqual(['办公', '用品']);
    expect(p).toEqual(at([1], 0));
    p = backspace(d, p)!;
    expect(texts(d)).toEqual(['办公用品']);
    expect(p.offset).toBe(2);
  });

  it('AI 结果写回：替换用选区首字格式，插入到后面另起一段，跨表格不执行', () => {
    const d = newDocument();
    d.body = [newParagraph('请于月底报价。其它')];
    applyRunProps(d, at([0], 0), at([0], 7), { bold: true });
    // 反向选区（从后往前选）也一样
    const p = applyAiText(d, at([0], 7), at([0], 0), '请于9月30日前报价。\n谢谢。', 'replace')!;
    expect(texts(d)).toEqual(['请于9月30日前报价。', '谢谢。其它']);
    expect(p).toEqual(at([1], 3));
    expect(getPara(d, at([0], 0)).runs[0]!.props.bold).toBe(true);
    applyAiText(d, at([1], 0), at([1], 3), 'Thanks.', 'insert-after');
    expect(texts(d)).toEqual(['请于9月30日前报价。', '谢谢。', 'Thanks.其它']);
    const t = newDocument();
    t.body = [newParagraph('甲')];
    insertTable(t, at([0], 0), 2, 2);
    expect(applyAiText(t, at([1], 0), at([0, 0, 0, 0], 0), 'x', 'replace')).toBeNull();
  });

  it('粘贴多行文字生成多个段落', () => {
    const d = newDocument();
    insertText(d, at([0], 0), '第一段\r\n第二段\n第三段');
    expect(texts(d)).toEqual(['第一段', '第二段', '第三段']);
  });

  it('跨段删除并合并', () => {
    const d = newDocument();
    d.body = [newParagraph('甲甲甲'), newParagraph('乙乙'), newParagraph('丙丙丙')];
    deleteRange(d, at([0], 1), at([2], 2));
    expect(texts(d)).toEqual(['甲丙']);
  });

  it('跨容器删除（正文到表格内）不执行', () => {
    const d = newDocument();
    d.body = [newParagraph('甲')];
    insertTable(d, at([0], 0), 2, 2);
    expect(deleteRange(d, at([1], 0), at([0, 0, 0, 0], 0))).toBeNull();
  });

  it('设置字符格式只影响选区并合并相同格式', () => {
    const d = newDocument();
    d.body = [newParagraph('物资采购计划')];
    applyRunProps(d, at([0], 2), at([0], 4), { bold: true });
    const p = getPara(d, at([0], 0));
    expect(p.runs.map((r) => [r.text, r.props.bold])).toEqual([
      ['物资', false],
      ['采购', true],
      ['计划', false],
    ]);
    applyRunProps(d, at([0], 0), at([0], 6), { bold: true });
    expect(p.runs).toHaveLength(1);
  });

  it('带格式插入（光标处先设粗体再输入）', () => {
    const d = newDocument();
    d.body = [newParagraph('甲乙')];
    insertRaw(d, at([0], 1), '中', { ...d.defaults, bold: true });
    expect(getPara(d, at([0], 0)).runs.map((r) => [r.text, r.props.bold])).toEqual([
      ['甲', false],
      ['中', true],
      ['乙', false],
    ]);
  });

  it('表格插入、增删行列', () => {
    const d = newDocument();
    d.body = [newParagraph('前文后文')];
    const cell = insertTable(d, at([0], 2), 2, 3)!;
    expect(texts(d)).toEqual(['前文', '[表格]', '后文']);
    expect(cell).toEqual(at([1, 0, 0, 0], 0));
    insertText(d, cell, '单元格');
    insertRow(d, cell, true);
    insertColumn(d, cell);
    const t = d.body[1];
    if (t.kind !== 'table') throw new Error();
    expect(t.rows).toHaveLength(3);
    expect(t.cols).toHaveLength(4);
    expect(t.cols.reduce((a, b) => a + b, 0)).toBeCloseTo(415.3, 1);
    deleteRow(d, at([1, 0, 0, 0], 0));
    expect(t.rows).toHaveLength(2);
  });

  it('选中文本、字数统计', () => {
    const d = newDocument();
    d.body = [newParagraph('采购 Laptop charger 4件，'), newParagraph('第二段')];
    expect(selectedText(d, at([0], 0), at([1], 2))).toBe('采购 Laptop charger 4件，\r\n第二');
    // 汉字 2+1+3=6，西文单词 Laptop charger 4 = 3
    expect(wordCount(d)).toBe(9);
  });

  it('撤销重做，连续输入合并为一步', () => {
    const d = newDocument();
    const h = new History<Pos>();
    h.record(d, at([0], 0), 'type');
    insertText(d, at([0], 0), '甲');
    h.record(d, at([0], 1), 'type');
    insertText(d, at([0], 1), '乙');
    const s = h.undo(d, at([0], 2))!;
    expect(paraText(s.doc.body[0] as never)).toBe('');
    const r = h.redo(s.doc, s.sel)!;
    expect(paraText(r.doc.body[0] as never)).toBe('甲乙');
  });
});

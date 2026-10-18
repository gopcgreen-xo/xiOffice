// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { cssColor, cssLength, fragmentToHtml, fragmentToText, htmlToFragment } from './clipboard';
import { copyFragment, insertFragment } from '../core/ops';
import { defaultRunProps, newDocument, newParagraph, newTable, paraText, type Paragraph, type Table } from '../core/model';

const base = defaultRunProps();

describe('HTML → 片段', () => {
  it('Word 的段落：行内字体、字号、加粗、颜色、对齐、首行缩进；&nbsp; 空段落保留', () => {
    const html = `<html><body><!--StartFragment--><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:16.0pt;font-family:黑体'>采购合同</span></b></p>
<p class=MsoNormal style='text-indent:21.0pt'><span style='font-family:宋体;mso-ascii-font-family:"Times New Roman"'>甲方：</span><span lang=EN-US style='color:red'>ABC<o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal>乙方</p><!--EndFragment--></body></html>`;
    const f = htmlToFragment(html, base, 415)! as Paragraph[];
    expect(f.length).toBe(4);
    expect(f[0].props.align).toBe('center');
    expect(f[0].runs[0]).toMatchObject({ text: '采购合同', props: { bold: true, size: 16, fontEA: '黑体' } });
    expect(f[1].props.firstLine).toBe(21);
    expect(f[1].runs.map((r) => [r.text, r.props.color])).toEqual([
      ['甲方：', 'auto'],
      ['ABC', 'FF0000'],
    ]);
    expect(f[2].runs).toEqual([]);
  });

  it('表格：colspan / rowspan → 合并单元格，列宽按 td 宽度', () => {
    const html = `<table border=1><tr><td colspan=2 style="width:100pt">名称</td><td rowspan=2>备注</td></tr><tr><td width=40>A</td><td>B</td></tr></table>`;
    const [t] = htmlToFragment(html, base, 400)! as Table[];
    expect(t.kind).toBe('table');
    expect(t.borders).toBe(true);
    expect(t.rows[0].cells.map((c) => [c.gridSpan, c.vMerge, paraText(c.paragraphs[0])])).toEqual([
      [2, 'none', '名称'],
      [1, 'restart', '备注'],
    ]);
    expect(t.rows[1].cells.map((c) => [c.gridSpan, c.vMerge, paraText(c.paragraphs[0])])).toEqual([
      [1, 'none', 'A'],
      [1, 'none', 'B'],
      [1, 'continue', ''],
    ]);
    expect(t.cols[0]).toBe(30);
    expect(t.cols.reduce((a, b) => a + b, 0)).toBeCloseTo(400);
  });

  it('Excel 的表格默认有边框；网页列表加编号；看不懂的标签取文字', () => {
    const excel = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><!--StartFragment--><table><tr><td>1</td><td>2</td></tr></table><!--EndFragment--></body></html>`;
    expect((htmlToFragment(excel, base, 400)![0] as Table).borders).toBe(true);
    const web = `<ol><li>第一</li><li><custom-tag>第二</custom-tag></li></ol><ul><li>x</li></ul>`;
    expect((htmlToFragment(web, base, 400) as Paragraph[]).map(paraText)).toEqual(['1. 第一', '2. 第二', '• x']);
  });

  it('样式小工具', () => {
    expect([cssColor('#f00'), cssColor('rgb(0, 128, 0)'), cssColor('windowtext'), cssColor('black')]).toEqual(['FF0000', '008000', 'auto', 'auto']);
    expect([cssLength('16px'), cssLength('1cm')?.toFixed(2), cssLength('2em', 12)]).toEqual([12, '28.35', 24]);
  });
});

describe('复制 / 粘贴片段', () => {
  function sample() {
    const d = newDocument();
    d.body = [newParagraph('第一段'), newTable(2, 1, 400, base), newParagraph('第三段')];
    ((d.body[1] as Table).rows[0].cells[0].paragraphs[0].runs = [{ text: '格', props: base }]);
    return d;
  }

  it('跨表格的选区连表格一起复制；转成 HTML 和纯文本', () => {
    const d = sample();
    const f = copyFragment(d, { story: 'body', path: [0], offset: 2 }, { story: 'body', path: [2], offset: 2 });
    expect(f.map((b) => b.kind)).toEqual(['p', 'table', 'p']);
    expect(paraText(f[0] as Paragraph)).toBe('段');
    expect(fragmentToText(f)).toBe('段\r\n格\t\r\n第三');
    const html = fragmentToHtml(f);
    expect(html).toContain('<table');
    // HTML 再读回来结构一致
    expect(htmlToFragment(html, base, 400)!.map((b) => b.kind)).toEqual(['p', 'table', 'p']);
  });

  it('粘贴一段：带格式插在光标处；多段：首尾接上原来的文字，中间整段插入', () => {
    const d = newDocument();
    d.body = [newParagraph('AB')];
    const bold = { ...base, bold: true };
    let end = insertFragment(d, { story: 'body', path: [0], offset: 1 }, [{ ...newParagraph(), runs: [{ text: 'x', props: bold }] }]);
    expect(end.offset).toBe(2);
    expect((d.body[0] as Paragraph).runs.map((r) => [r.text, r.props.bold])).toEqual([
      ['A', false],
      ['x', true],
      ['B', false],
    ]);
    end = insertFragment(d, { story: 'body', path: [0], offset: 1 }, [newParagraph('1'), newTable(2, 2, 400, base), newParagraph('3')]);
    expect(d.body.map((b) => (b.kind === 'p' ? paraText(b) : 'T'))).toEqual(['A1', 'T', '3xB']);
    expect(end).toEqual({ story: 'body', path: [2], offset: 1 });
  });

  it('表格里粘贴表格：转成制表符分隔的段落', () => {
    const d = sample();
    const t = newTable(2, 1, 400, base);
    t.rows[0].cells[0].paragraphs[0].runs = [{ text: 'a', props: base }];
    t.rows[0].cells[1].paragraphs[0].runs = [{ text: 'b', props: base }];
    insertFragment(d, { story: 'body', path: [1, 0, 1, 0], offset: 0 }, [t]);
    expect(paraText((d.body[1] as Table).rows[0].cells[1].paragraphs[0])).toBe('a\tb');
  });
});

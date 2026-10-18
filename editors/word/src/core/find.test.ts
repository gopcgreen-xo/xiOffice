import { describe, expect, it } from 'vitest';
import { findAll, findNext, replaceAll, replaceMatch } from './find';
import { defaultRunProps, newDocument, newParagraph, newTable, paraText, type Paragraph, type Table } from './model';

function sample() {
  const d = newDocument();
  const t = newTable(1, 1, 400, defaultRunProps());
  t.rows[0].cells[0].paragraphs[0].runs = [{ text: '甲方盖章', props: defaultRunProps() }];
  d.body = [newParagraph('甲方：ABC 公司，甲方代表'), t, newParagraph('Abc abcd')];
  d.footer = [newParagraph('甲方（盖章）')];
  return d;
}

describe('Word 查找替换', () => {
  it('正文、表格、页脚都能找到；区分大小写；全字匹配', () => {
    const d = sample();
    expect(findAll(d, '甲方').map((m) => [m.story, m.path.join('.'), m.offset])).toEqual([
      ['body', '0', 0],
      ['body', '0', 10],
      ['body', '1.0.0.0', 0],
      ['footer', '0', 0],
    ]);
    expect(findAll(d, 'abc').length).toBe(3);
    expect(findAll(d, 'abc', { matchCase: true }).length).toBe(1);
    expect(findAll(d, 'abc', { whole: true }).length).toBe(2);
  });

  it('下一个按文档顺序，到末尾从头开始；往前找', () => {
    const d = sample();
    const all = findAll(d, '甲方');
    expect(findNext(all, { story: 'body', path: [0], offset: 0 })?.offset).toBe(10);
    expect(findNext(all, { story: 'footer', path: [0], offset: 0 })?.path).toEqual([0]);
    expect(findNext(all, { story: 'body', path: [0], offset: 0 }, true)?.story).toBe('footer');
  });

  it('替换保留原文字的格式；全部替换', () => {
    const d = sample();
    const p = d.body[0] as Paragraph;
    p.runs = [
      { text: '甲方', props: { ...defaultRunProps(), bold: true } },
      { text: '：ABC 公司，甲方代表', props: defaultRunProps() },
    ];
    const end = replaceMatch(d, findAll(d, '甲方')[0], '某地某公司');
    expect(end.offset).toBe(5);
    expect(p.runs[0]).toMatchObject({ text: '某地某公司', props: { bold: true } });
    expect(replaceAll(d, '甲方', '买方')).toBe(3);
    expect(paraText(p)).toBe('某地某公司：ABC 公司，买方代表');
    expect(paraText((d.body[1] as Table).rows[0].cells[0].paragraphs[0])).toBe('买方盖章');
    expect(paraText(d.footer[0])).toBe('买方（盖章）');
  });
});

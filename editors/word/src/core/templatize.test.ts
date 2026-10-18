import { afterEach, describe, expect, it } from 'vitest';
import { newDocument, newParagraph, paraText, type Doc, type Paragraph } from './model';
import { tracking } from './ops';
import { applyTemplateHits, templateParagraphs } from './templatize';

afterEach(() => {
  tracking.on = false;
});

const FIELDS = new Set(['公司1', '公司2', '采购合同号', '总金额', '总金额U', '交货地点']);

function docOf(...texts: string[]): Doc {
  const d = newDocument();
  d.body = texts.map((t) => newParagraph(t));
  return d;
}
const text = (d: Doc, i: number) => paraText(d.body[i] as Paragraph);

describe('合同 → 模版（按建议把数据换成 {{字段}}）', () => {
  it('不开修订：指定段落替换、标签后插入、没给段落时全文替换', () => {
    const d = docOf('需方：bb有限公司     合同编号： ', '供方：XX有限公司', '交货地点：bb有限公司', '合计 709967.88');
    const { refs, texts } = templateParagraphs(d);
    expect(texts).toEqual(['需方：bb有限公司     合同编号： ', '供方：XX有限公司', '交货地点：bb有限公司', '合计 709967.88']);
    const r = applyTemplateHits(
      d,
      refs,
      [
        { text: 'bb有限公司', field: '公司1' },
        { p: 0, after: '合同编号：', field: '采购合同号' },
        { p: 1, text: 'XX有限公司', field: '公司2' },
        { p: 3, text: '709967.88', field: '总金额' },
      ],
      FIELDS,
    );
    expect(r.skipped).toEqual([]);
    expect(r.applied).toBe(5);
    expect(text(d, 0)).toBe('需方：{{公司1}}     合同编号：{{采购合同号}} ');
    expect(text(d, 1)).toBe('供方：{{公司2}}');
    expect(text(d, 2)).toBe('交货地点：{{公司1}}');
    expect(text(d, 3)).toBe('合计 {{总金额}}');
  });

  it('修订打开：原文标为删除、占位符记为插入；再次读取时看不到已删除的字', () => {
    const d = docOf('供方：XX有限公司');
    const { refs } = templateParagraphs(d);
    tracking.on = true;
    tracking.author = '审核';
    applyTemplateHits(d, refs, [{ p: 0, text: 'XX有限公司', field: '公司2' }], FIELDS);
    const p = d.body[0] as Paragraph;
    expect(p.runs.map((r) => [r.text, r.props.rev?.kind ?? ''])).toEqual([
      ['供方：', ''],
      ['{{公司2}}', 'ins'],
      ['XX有限公司', 'del'],
    ]);
    expect(templateParagraphs(d).texts).toEqual(['供方：{{公司2}}']);
    // 再跑一次同样的建议：已删除的字不再匹配
    const again = applyTemplateHits(d, templateParagraphs(d).refs, [{ p: 0, text: 'XX有限公司', field: '公司2' }], FIELDS);
    expect(again.applied).toBe(0);
  });

  it('跳过：未知字段、纯数字不给段落、找不到、已有占位符里面；重叠时长的优先', () => {
    const d = docOf('总价款：柒拾万元整（含税）', '{{公司1}} 地址', '税率 13%');
    const { refs } = templateParagraphs(d);
    const r = applyTemplateHits(
      d,
      refs,
      [
        { p: 0, text: '柒拾万元整', field: '总金额U' },
        { p: 0, text: '柒拾万', field: '总金额' },
        { p: 1, text: '公司1', field: '公司2' },
        { text: '13%', field: '总金额' },
        { p: 2, text: '不存在', field: '总金额' },
        { p: 2, text: '13%', field: '税率' },
        { p: 9, text: '13%', field: '总金额' },
      ],
      FIELDS,
    );
    expect(r.applied).toBe(1);
    expect(text(d, 0)).toBe('总价款：{{总金额U}}（含税）');
    expect(text(d, 1)).toBe('{{公司1}} 地址');
    expect(r.skipped.map((s) => s.reason)).toEqual(['原文里没找到这段文字', '没有指明段落', '原文里没找到这段文字', '没有这个字段：税率', '段落序号不对：9']);
  });

  it('after 用在已经有值的标签上：换掉这个值（到两个空格 / 段尾为止，不含末尾句号）；紧跟下一个标签时照常插入', () => {
    const d = docOf('工程编号：H1702', '项目工程号：N2345-01      船号： H1702', '交货时间：按甲方通知送货。', '电话：      传真：', '合同编号：WC2026-0158');
    const { refs } = templateParagraphs(d);
    const r = applyTemplateHits(
      d,
      refs,
      [
        { p: 0, after: '工程编号：', field: '船号' },
        { p: 1, after: '项目工程号：', field: '船号' },
        { p: 1, after: '船号：', field: '船号' },
        { p: 2, after: '交货时间：', field: '交货日期' },
        { p: 3, after: '电话：', field: '公司电话1' },
        // AI 同时给了 text 和 after 指向同一个值：只换一次
        { p: 4, text: 'WC2026-0158', field: '采购合同号' },
        { p: 4, after: '合同编号：', field: '采购合同号' },
      ],
      new Set([...FIELDS, '船号', '交货日期', '公司电话1']),
    );
    expect(r.skipped).toEqual([]);
    expect(templateParagraphs(d).texts).toEqual(['工程编号：{{船号}}', '项目工程号：{{船号}}      船号： {{船号}}', '交货时间：{{交货日期}}。', '电话：{{公司电话1}}      传真：', '合同编号：{{采购合同号}}']);
    expect(r.applied).toBe(6);
  });

  it('表格“标签一格、值一格”：标签在格末、右边那格空着 → 填进右边那格；右边有字或标签后还有字时照常插在标签后', () => {
    const cell = (t: string, vMerge = 'none') => ({ paragraphs: [newParagraph(t)], vMerge });
    const d = docOf('标题');
    d.body.push({
      kind: 'table',
      cols: [100, 100, 100],
      rows: [
        { cells: [cell('合同编号：'), cell(''), cell('')] },
        { cells: [cell('工程编号：'), cell('已有'), cell('')] },
        { cells: [cell('船号：  '), cell('', 'continue'), cell('')] },
        { cells: [cell('电话：'), cell('传真：'), cell('')] },
      ],
    } as unknown as Doc['body'][number]);
    const { refs, texts } = templateParagraphs(d);
    const at = (t: string) => texts.indexOf(t);
    const r = applyTemplateHits(
      d,
      refs,
      [
        { p: at('合同编号：'), after: '合同编号：', field: '采购合同号' },
        { p: at('工程编号：'), after: '工程编号：', field: '船号' },
        { p: at('船号：  '), after: '船号：', field: '船号' },
        { p: at('电话：'), after: '电话：', field: '公司电话1' },
      ],
      new Set([...FIELDS, '船号', '公司电话1']),
    );
    expect(r.skipped).toEqual([]);
    const t = d.body[1] as unknown as { rows: { cells: { paragraphs: Paragraph[] }[] }[] };
    const grid = t.rows.map((row) => row.cells.map((c) => paraText(c.paragraphs[0])));
    expect(grid).toEqual([
      ['合同编号：', '{{采购合同号}}', ''],
      ['工程编号：{{船号}}', '已有', ''],
      ['船号：{{船号}}  ', '', ''],
      ['电话：{{公司电话1}}', '传真：', ''],
    ]);
  });

  it('标签后已经有占位符就不再插入；表格单元格里的段落也有序号', () => {
    const d = docOf('合同编号：{{采购合同号}}');
    d.body.push({
      kind: 'table',
      cols: [100],
      rows: [{ cells: [{ paragraphs: [newParagraph('交货地点：青岛')] }] }],
    } as unknown as Doc['body'][number]);
    const { refs, texts } = templateParagraphs(d);
    expect(texts).toEqual(['合同编号：{{采购合同号}}', '交货地点：青岛']);
    expect(refs[1]).toEqual({ story: 'body', path: [1, 0, 0, 0] });
    const r = applyTemplateHits(
      d,
      refs,
      [
        { p: 0, after: '合同编号：', field: '采购合同号' },
        { p: 1, text: '青岛', field: '交货地点' },
      ],
      FIELDS,
    );
    expect(r.applied).toBe(1);
    expect(r.skipped.map((s) => s.reason)).toEqual(['没找到标签文字']);
  });
});

// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { strToU8, unzipSync, zipSync, strFromU8 } from 'fflate';
import { readDocx, writeDocx } from '../ooxml/docx';
import { allRevisions, addComment, backspace, commentRanges, deleteComment, deleteRange, hasRevisions, insertRaw, nextRevision, replyComment, resolveRevisions, tracking } from './ops';
import { defaultRunProps, newDocument, newParagraph, paraText, type Paragraph } from './model';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const at = (offset: number, path = [0]) => ({ story: 'body' as const, path, offset });

afterEach(() => {
  tracking.on = false;
});

describe('修订', () => {
  it('读取 Word 的插入 / 删除：文字都在，带修订人；接受全部、拒绝全部', () => {
    const doc = `<w:document ${W}><w:body><w:p><w:r><w:t>单价</w:t></w:r><w:del w:id="1" w:author="张三" w:date="2026-09-30T10:00:00Z"><w:r><w:delText>100</w:delText></w:r></w:del><w:ins w:id="2" w:author="张三"><w:r><w:t>98</w:t></w:r></w:ins><w:r><w:t>元</w:t></w:r></w:p></w:body></w:document>`;
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc) }));
    expect(d.lossy).toEqual([]);
    const p = d.body[0] as Paragraph;
    expect(paraText(p)).toBe('单价10098元');
    expect(p.runs.map((r) => r.props.rev?.kind ?? '')).toEqual(['', 'del', 'ins', '']);
    expect(p.runs[1].props.rev).toEqual({ kind: 'del', author: '张三', date: '2026-09-30T10:00:00Z' });
    // 保存后修订还在
    const xml = strFromU8(unzipSync(writeDocx(d))['word/document.xml']);
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="张三" w:date="2026-09-30T10:00:00Z"><w:r>.*<w:delText xml:space="preserve">100<\/w:delText><\/w:r><\/w:del>/);
    expect(xml).toContain('<w:t xml:space="preserve">98</w:t></w:r></w:ins>');
    const a = structuredClone(d);
    expect(resolveRevisions(a, true)).toBe(2);
    expect(paraText(a.body[0] as Paragraph)).toBe('单价98元');
    expect(hasRevisions(a)).toBe(false);
    const r = structuredClone(d);
    resolveRevisions(r, false);
    expect(paraText(r.body[0] as Paragraph)).toBe('单价100元');
  });

  it('修订模式：输入记为插入；删除标为删除；删掉自己的插入就真删；Backspace 跳过已删除的字', () => {
    const d = newDocument();
    d.body = [newParagraph('ABCD')];
    tracking.on = true;
    tracking.author = '李四';
    insertRaw(d, at(2), 'xy');
    const p = d.body[0] as Paragraph;
    expect(p.runs.map((r) => [r.text, r.props.rev?.kind ?? ''])).toEqual([
      ['AB', ''],
      ['xy', 'ins'],
      ['CD', ''],
    ]);
    deleteRange(d, at(1), at(5));
    expect(paraText(p)).toBe('ABCD');
    expect(p.runs.map((r) => [r.text, r.props.rev?.kind ?? ''])).toEqual([
      ['A', ''],
      ['BC', 'del'],
      ['D', ''],
    ]);
    // 光标在 D 前：Backspace 跳过 BC，删掉 A
    const pos = backspace(d, at(3));
    expect(pos?.offset).toBe(0);
    expect(p.runs.map((r) => [r.text, r.props.rev?.kind ?? ''])).toEqual([
      ['ABC', 'del'],
      ['D', ''],
    ]);
    expect(nextRevision(d, at(0))?.[0].offset).toBe(0);
    tracking.on = false;
    // 关闭修订后输入：不带修订标记（不沿用删除标记）
    insertRaw(d, at(1), 'z');
    expect(p.runs.find((r) => r.text === 'z')?.props.rev).toBeUndefined();
  });
});

describe('修订合并', () => {
  const rev = (kind: 'ins' | 'del', author: string) => ({ ...defaultRunProps(), rev: { kind, author } }) as never;
  const build = (runs: [string, 'ins' | 'del' | '', string?][]) => {
    const d = newDocument();
    const p = newParagraph('');
    p.runs = runs.map(([text, kind, author]) => ({ text, props: kind ? rev(kind, author ?? '张三') : defaultRunProps() }));
    d.body = [p];
    return d;
  };
  it('同一人紧挨着的删除 + 插入算一处；接受 / 拒绝一次处理完', () => {
    const d = build([['单价', ''], ['100', 'del'], ['98', 'ins'], ['元', '']]);
    expect(allRevisions(d)).toHaveLength(1);
    const accepted = structuredClone(d);
    const [a, b] = allRevisions(accepted)[0];
    resolveRevisions(accepted, true, [a, b]);
    expect(paraText(accepted.body[0] as Paragraph)).toBe('单价98元');
    resolveRevisions(d, false, allRevisions(d)[0]);
    expect(paraText(d.body[0] as Paragraph)).toBe('单价100元');
  });
  it('不同人、或中间隔着正常文字，仍是不同的修订', () => {
    expect(allRevisions(build([['a', 'del'], ['b', 'ins', '李四']]))).toHaveLength(2);
    expect(allRevisions(build([['a', 'del'], ['x', ''], ['b', 'ins']]))).toHaveLength(2);
  });
});

describe('批注', () => {
  it('加批注、回复、保存、读回：范围、作者、回复关系；删除', () => {
    const d = newDocument();
    d.body = [newParagraph('合同金额以实际结算为准')];
    const id = addComment(d, at(0), at(4), '金额要写大写', '王五')!;
    replyComment(d, id, '已改', '赵六');
    expect(commentRanges(d).get(id)?.map((p) => p.offset)).toEqual([1, 5]);
    expect(paraText(d.body[0] as Paragraph).replace(/￼/g, '')).toBe('合同金额以实际结算为准');
    const out = writeDocx(d);
    const files = unzipSync(out);
    expect(strFromU8(files['word/comments.xml'])).toContain('w:author="王五"');
    expect(strFromU8(files['[Content_Types].xml'])).toContain('/word/comments.xml');
    const doc = strFromU8(files['word/document.xml']);
    // 回复跟着被回复的那条写范围标记
    expect(doc.match(/commentRangeStart/g)?.length).toBe(2);
    expect(doc.match(/commentReference/g)?.length).toBe(2);
    const back = readDocx(out);
    expect(back.comments?.map((c) => [c.author, c.text])).toEqual([
      ['王五', '金额要写大写'],
      ['赵六', '已改'],
    ]);
    deleteComment(back, back.comments![0].id);
    expect(back.comments).toEqual([]);
    expect((back.body[0] as Paragraph).runs.some((r) => r.mark)).toBe(false);
  });
});

describe('文字格式：删除线、上下标、突出显示', () => {
  it('读写往返', () => {
    const d = newDocument();
    const p = newParagraph('');
    const base = defaultRunProps();
    p.runs = [
      { text: 'm', props: base },
      { text: '2', props: { ...base, vertAlign: 'superscript' } },
      { text: '作废', props: { ...base, strike: true } },
      { text: '重点', props: { ...base, highlight: 'FFFF00' } },
      { text: '底纹', props: { ...base, highlight: 'D9D9D9' } },
    ];
    d.body = [p];
    const out = writeDocx(d);
    const xml = strFromU8(unzipSync(out)['word/document.xml']);
    expect(xml).toContain('<w:highlight w:val="yellow"/>');
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>');
    const back = readDocx(out);
    expect(back.lossy).toEqual([]);
    expect((back.body[0] as Paragraph).runs.map((r) => [r.text, r.props.vertAlign ?? '', r.props.strike ?? false, r.props.highlight ?? ''])).toEqual([
      ['m', '', false, ''],
      ['2', 'superscript', false, ''],
      ['作废', '', true, ''],
      ['重点', '', false, 'FFFF00'],
      ['底纹', '', false, 'D9D9D9'],
    ]);
  });
});

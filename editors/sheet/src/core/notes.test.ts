import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { deleteNote, setNote } from './notes';
import { relTarget } from './parts';
import { applyEntry, insertDelete, setInput, Tx } from './ops';
import { saveXlsx } from './save';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';
import { readDrawings } from './drawing';

const edit = (b: Workbook, fn: (tx: Tx) => unknown) => {
  const tx = new Tx(b, 'x');
  fn(tx);
  return tx.commit();
};
const parts = (bytes: Uint8Array) => unzipSync(bytes);
const txt = (f: Record<string, Uint8Array>, p: string) => strFromU8(f[p]);
const reopen = (bytes: Uint8Array) => {
  const b = readWorkbook(bytes);
  b.source = bytes;
  return b;
};

describe('批注：修改', () => {
  it('新建、修改、删除，可撤销；文字为空等于删除', () => {
    const b = blankWorkbook();
    const e1 = edit(b, (tx) => setNote(tx, 0, 1, 2, '第一条', '小王'))!;
    expect(b.sheets[0].notes).toEqual([{ r: 1, c: 2, author: '小王', text: '第一条' }]);
    edit(b, (tx) => setNote(tx, 0, 0, 0, '前面', '小李'));
    expect(b.sheets[0].notes!.map((n) => [n.r, n.c])).toEqual([[0, 0], [1, 2]]);
    edit(b, (tx) => setNote(tx, 0, 1, 2, '改过', '别人'));
    expect(b.sheets[0].notes![1]).toMatchObject({ text: '改过', author: '小王' });
    edit(b, (tx) => setNote(tx, 0, 0, 0, '   ', ''));
    expect(b.sheets[0].notes!.length).toBe(1);
    edit(b, (tx) => deleteNote(tx, 0, 1, 2));
    expect(b.sheets[0].notes).toBeUndefined();
    applyEntry(b, e1, -1);
    expect(b.sheets[0].notes).toBeUndefined();
  });

  it('线程批注只读', () => {
    const b = blankWorkbook();
    b.sheets[0].notes = [{ r: 0, c: 0, author: 'tc={x}', text: '[Threaded comment]\n你的 Excel 版本……' }];
    const tx = new Tx(b, 'x');
    expect(setNote(tx, 0, 0, 0, '改', 'a')).toBe(false);
    expect(b.sheets[0].notes![0].text).toMatch(/^\[Threaded comment\]/);
  });
});

describe('批注：写出 xlsx', () => {
  it('新建批注：comments + VML + 关系 + 内容类型 + legacyDrawing，读回来一致', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setInput(tx, 0, 0, 0, '数据'));
    edit(b, (tx) => setNote(tx, 0, 0, 0, '这里要核对 & <确认>', '小王'));
    edit(b, (tx) => setNote(tx, 0, 3, 1, '第二条', '小李'));
    const out = saveXlsx(b);
    const f = parts(out);
    expect(txt(f, 'xl/comments1.xml')).toContain('<author>小王</author><author>小李</author>');
    expect(txt(f, 'xl/comments1.xml')).toContain('<comment ref="A1" authorId="0"><text><t xml:space="preserve">这里要核对 &amp; &lt;确认&gt;</t></text></comment>');
    expect(txt(f, 'xl/comments1.xml')).toContain('<comment ref="B4" authorId="1">');
    const vml = txt(f, 'xl/drawings/vmlDrawing1.vml');
    expect(vml).toContain('<x:Row>0</x:Row><x:Column>0</x:Column>');
    expect(vml).toContain('<x:Row>3</x:Row><x:Column>1</x:Column>');
    expect(vml).toContain('<o:idmap v:ext="edit" data="1"/>');
    expect(txt(f, 'xl/worksheets/_rels/sheet1.xml.rels')).toMatch(/Type="[^"]*\/comments" Target="\.\.\/comments1\.xml"/);
    expect(txt(f, 'xl/worksheets/_rels/sheet1.xml.rels')).toMatch(/Type="[^"]*\/vmlDrawing" Target="\.\.\/drawings\/vmlDrawing1\.vml"/);
    const ct = txt(f, '[Content_Types].xml');
    expect(ct).toContain('<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>');
    expect(ct).toMatch(/<Default Extension="vml" ContentType="application\/vnd.openxmlformats-officedocument.vmlDrawing"\/>/);
    const sheet = txt(f, 'xl/worksheets/sheet1.xml');
    expect(sheet).toMatch(/<legacyDrawing r:id="rId2"\/>/);
    expect(sheet.indexOf('<pageMargins')).toBeLessThan(sheet.indexOf('<legacyDrawing'));
    const back = reopen(out);
    expect(back.sheets[0].notes).toEqual([
      { r: 0, c: 0, author: '小王', text: '这里要核对 & <确认>' },
      { r: 3, c: 1, author: '小李', text: '第二条' },
    ]);
  });

  it('没改批注：字节不变；改一条：其它批注保留原文（含富文本），VML 里没改的框原样保留', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setNote(tx, 0, 0, 0, '一', '甲'));
    edit(b, (tx) => setNote(tx, 0, 1, 0, '二', '乙'));
    let out = saveXlsx(b);
    // 模拟 Excel 写的：第一条是带格式的富文本（作者名加粗）
    const f = parts(out);
    f['xl/comments1.xml'] = strToU8(txt(f, 'xl/comments1.xml').replace(/<comment ref="A1"[\s\S]*?<\/comment>/, '<comment ref="A1" authorId="0" shapeId="0"><text><r><rPr><b/><sz val="9"/></rPr><t>甲:</t></r><r><rPr><sz val="9"/></rPr><t xml:space="preserve">\n一</t></r></text></comment>'));
    out = zipSync(f);
    const b2 = reopen(out);
    expect(b2.sheets[0].notes![0].text).toBe('甲:\n一');
    expect(Array.from(saveXlsx(b2))).toEqual(Array.from(out));
    edit(b2, (tx) => setNote(tx, 0, 1, 0, '二（改）', ''));
    const f2 = parts(saveXlsx(b2));
    const c = txt(f2, 'xl/comments1.xml');
    expect(c).toContain('<r><rPr><b/><sz val="9"/></rPr><t>甲:</t></r>');
    expect(c).toContain('二（改）');
    expect(c).not.toContain('<t xml:space="preserve">二</t>');
    expect(txt(f2, 'xl/drawings/vmlDrawing1.vml').match(/<v:shape\b/g)!.length).toBe(2);
    expect(txt(f2, 'xl/drawings/vmlDrawing1.vml')).toContain(txt(f, 'xl/drawings/vmlDrawing1.vml').match(/<v:shape\b[\s\S]*?<\/v:shape>/)![0]);
  });

  it('删除最后一条：comments、VML、关系、legacyDrawing 全部去掉', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setNote(tx, 0, 0, 0, '一', '甲'));
    const out = saveXlsx(b);
    const b2 = reopen(out);
    edit(b2, (tx) => deleteNote(tx, 0, 0, 0));
    const f = parts(saveXlsx(b2));
    expect(Object.keys(f).filter((k) => /comments|vmlDrawing/.test(k))).toEqual([]);
    expect(f['xl/worksheets/_rels/sheet1.xml.rels']).toBeUndefined();
    expect(txt(f, 'xl/worksheets/sheet1.xml')).not.toContain('legacyDrawing');
    expect(txt(f, '[Content_Types].xml')).not.toContain('comments1.xml');
  });

  it('插入行后批注跟着移动，再保存也一致', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setNote(tx, 0, 2, 0, '三行', '甲'));
    const out = saveXlsx(b);
    const b2 = reopen(out);
    edit(b2, (tx) => insertDelete(tx, 0, 'row', 0, 2));
    const f = parts(saveXlsx(b2));
    expect(txt(f, 'xl/comments1.xml')).toContain('ref="A5"');
    expect(txt(f, 'xl/drawings/vmlDrawing1.vml')).toContain('<x:Row>4</x:Row>');
    expect(readDrawings(f, 'xl/worksheets/sheet1.xml').notes[0]).toMatchObject({ r: 4, c: 0 });
  });

  it('第二张表的批注用新的 idmap；Target 相对路径', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setNote(tx, 0, 0, 0, 'a', '甲'));
    let out = saveXlsx(b);
    const b2 = reopen(out);
    edit(b2, (tx) => setNote(tx, 0, 1, 0, 'b', '甲'));
    out = saveXlsx(b2);
    expect(txt(parts(out), 'xl/drawings/vmlDrawing1.vml').match(/_x0000_s(\d+)/g)).toEqual(['_x0000_s1025', '_x0000_s1026']);
    expect(relTarget('xl/worksheets/sheet1.xml', 'xl/comments1.xml')).toBe('../comments1.xml');
    expect(relTarget('xl/worksheets/sheet1.xml', 'xl/drawings/drawing1.xml')).toBe('../drawings/drawing1.xml');
    expect(relTarget('xl/drawings/drawing1.xml', 'xl/charts/chart1.xml')).toBe('../charts/chart1.xml');
    expect(relTarget('xl/workbook.xml', 'xl/pivotCache/pivotCacheDefinition1.xml')).toBe('pivotCache/pivotCacheDefinition1.xml');
  });
});

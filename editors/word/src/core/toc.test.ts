// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readDocx, writeDocx } from '../ooxml/docx';
import { layoutDocument, type Measurer } from './layout';
import { newDocument, newParagraph, paraText, textWidth, type Paragraph } from './model';
import { applyParaStyle, blockPages, headings, insertToc, tocRange, updateToc } from './toc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const fake: Measurer = { width: (ch) => (/[一-鿿]/.test(ch) ? 1 : 0.5), metrics: () => ({ ascent: 0.8, descent: 0.2 }) };
const at = (bi: number, offset = 0) => ({ story: 'body' as const, path: [bi], offset });

describe('样式与目录', () => {
  it('读取：中文 Word 的样式 ID 是数字，按样式名认出标题；大纲级别来自样式', () => {
    const styles = `<w:styles ${W}><w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 2"/></w:style></w:styles>`;
    const doc = `<w:document ${W}><w:body><w:p><w:pPr><w:pStyle w:val="2"/></w:pPr><w:r><w:t>第一章</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="3"/></w:pPr><w:r><w:t>一节</w:t></w:r></w:p><w:p><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>直接设的级别</w:t></w:r></w:p><w:p><w:r><w:t>正文</w:t></w:r></w:p></w:body></w:document>`;
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc), 'word/styles.xml': strToU8(styles) }));
    expect(d.body.map((b) => [(b as Paragraph).props.style, (b as Paragraph).props.outline])).toEqual([
      ['Heading1', 0],
      ['Heading2', 1], // 样式里没写 outlineLvl：内置标题样式本身带级别
      [undefined, 2],
      [undefined, undefined],
    ]);
    expect(headings(d).map((h) => [h.level, h.text])).toEqual([[1, '第一章'], [2, '一节'], [3, '直接设的级别']]);
  });

  it('制表位：样式和段落合并、clear 清除；右对齐制表位让页码右边对齐', () => {
    const styles = `<w:styles ${W}><w:style w:type="paragraph" w:styleId="T"><w:name w:val="x"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="840"/><w:tab w:val="left" w:pos="1680"/></w:tabs></w:pPr></w:style></w:styles>`;
    const doc = `<w:document ${W}><w:body><w:p><w:pPr><w:pStyle w:val="T"/><w:tabs><w:tab w:val="clear" w:pos="840"/><w:tab w:val="right" w:leader="dot" w:pos="4000"/></w:tabs></w:pPr><w:r><w:t>甲</w:t></w:r><w:r><w:tab/><w:t>乙</w:t></w:r><w:r><w:tab/><w:t>12</w:t></w:r></w:p></w:body></w:document>`;
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc), 'word/styles.xml': strToU8(styles) }));
    const p = d.body[0] as Paragraph;
    expect(p.props.tabs).toEqual([{ pos: 84, align: 'left', leader: 'none' }, { pos: 200, align: 'right', leader: 'dot' }]);
    const L = layoutDocument(d, fake);
    const line = L.pages[0].body[0].line;
    const size = p.runs[0].props.size;
    // 乙 从 84 磅开始；“12” 右边对齐到 200 磅
    expect(line.xs[2]).toBeCloseTo(84);
    expect(line.right).toBeCloseTo(200);
    expect(line.xs[4]).toBeCloseTo(200 - size);
    const dots = line.glyphs.find((g) => g.leader);
    expect(dots?.leader).toBe('.');
    // 写回
    const xml = strFromU8(unzipSync(writeDocx(d))['word/document.xml']);
    expect(xml).toContain('<w:tabs><w:tab w:val="left" w:leader="none" w:pos="1680"/><w:tab w:val="right" w:leader="dot" w:pos="4000"/></w:tabs>');
  });

  it('套用标题样式、插入并更新目录；保存后样式和大纲级别还在', () => {
    const d = newDocument();
    d.body = [newParagraph(''), newParagraph('概述'), newParagraph('正文内容'), newParagraph('范围'), newParagraph('附录')];
    applyParaStyle(d, at(1), at(1), 'Heading1');
    applyParaStyle(d, at(3), at(3), 'Heading2');
    applyParaStyle(d, at(4), at(4), 'Heading1');
    const h1 = d.body[1] as Paragraph;
    expect(h1.props).toMatchObject({ style: 'Heading1', outline: 0, keepNext: true });
    expect(h1.runs[0].props).toMatchObject({ size: 22, bold: true });
    applyParaStyle(d, at(4), at(4), 'Normal');
    expect((d.body[4] as Paragraph).props.outline).toBeUndefined();

    const end = insertToc(d, at(0), blockPages(layoutDocument(d, fake)));
    expect(end).toEqual(at(2)); // 两条目录 + 空段落
    expect(tocRange(d)).toEqual([0, 2]);
    const e = d.body[1] as Paragraph;
    expect(paraText(e)).toBe('范围\t1');
    expect(e.props).toMatchObject({ style: 'TOC2', tabs: [{ pos: textWidth(d.setup), align: 'right', leader: 'dot' }] });
    // 标题改名后更新目录
    (d.body[3] as Paragraph).runs[0].text = '总述';
    expect(updateToc(d, blockPages(layoutDocument(d, fake)))).toBe(true);
    expect(paraText(d.body[0] as Paragraph)).toBe('总述\t1');

    const z = unzipSync(writeDocx(d));
    expect(strFromU8(z['word/styles.xml'])).toContain('<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>');
    expect(strFromU8(z['word/styles.xml'])).toContain('w:styleId="TOC2"><w:name w:val="toc 2"/>');
    const back = readDocx(writeDocx(d));
    expect(headings(back).map((h) => h.text)).toEqual(['总述', '范围']);
    expect(tocRange(back)).toEqual([0, 2]);
  });
});

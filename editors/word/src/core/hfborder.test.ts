// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { readDocx, writeDocx } from '../ooxml/docx';
import { layoutDocument, type Measurer } from './layout';
import { newDocument, newParagraph, textWidth, type Paragraph } from './model';
import { isCJK } from './text';

const fake: Measurer = { width: (ch) => (isCJK(ch) ? 1 : 0.5), metrics: () => ({ ascent: 0.88, descent: 0.12 }) };

describe('段落边框（w:pBdr）与文本框架（w:framePr）', () => {
  it('页眉下边框：画在文字下方 space 磅处，横跨版心；边框的距离和线宽算进段落高度', () => {
    const d = newDocument();
    const h = newParagraph('页眉', { borders: { bottom: { w: 0.75, space: 1, color: 'auto' } } });
    d.header = [h];
    d.body = [newParagraph('正文')];
    const page = layoutDocument(d, fake).pages[0];
    const line = page.header[0];
    const seg = page.borders.find((b) => b.y1 === b.y2)!;
    expect(seg.y1).toBeCloseTo(line.y + line.line.height + 1 + 0.375);
    expect(seg.x1).toBeCloseTo(d.setup.marginLeft);
    expect(seg.x2).toBeCloseTo(d.setup.marginLeft + textWidth(d.setup));
  });

  it('文本框架：页脚里居中的页码按 xAlign 居中，不占高度（下一段和它在同一高度）', () => {
    const d = newDocument();
    const num = newParagraph('1', { frame: { hAnchor: 'margin', vAnchor: 'text', xAlign: 'center', x: 0, y: 0 } });
    d.footer = [num, newParagraph('')];
    d.body = [newParagraph('正文')];
    const page = layoutDocument(d, fake).pages[0];
    const [f0, f1] = page.footer;
    expect(f0.y).toBeCloseTo(f1.y);
    const g = f0.line.glyphs[0];
    expect(f0.x + g.x + g.w / 2).toBeCloseTo(d.setup.marginLeft + textWidth(d.setup) / 2);
  });

  it('边框和框架写出后读回一致', () => {
    const d = newDocument();
    d.body = [
      newParagraph('甲', { borders: { top: { w: 0.5, space: 1, color: 'FF0000' }, bottom: { w: 0.75, space: 4, color: 'auto' } } }),
      newParagraph('乙', { frame: { hAnchor: 'margin', vAnchor: 'text', xAlign: 'right', x: 0, y: 0 } }),
    ];
    const [a, b] = readDocx(writeDocx(d)).body as Paragraph[];
    expect(a.props.borders).toEqual(d.body[0].kind === 'p' ? (d.body[0] as Paragraph).props.borders : null);
    expect(b.props.frame).toMatchObject({ hAnchor: 'margin', xAlign: 'right' });
  });
});

describe('页眉页脚里的表格和图片（W18）', () => {
  it('页眉里的表格读成只读占位并按表格排；比上边距高时正文往下让；保存时写回原表格', async () => {
    const { zipSync, strToU8, unzipSync, strFromU8 } = await import('fflate');
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const row = (t: string) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc></w:tr>`;
    const tbl = `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${[1, 2, 3, 4, 5, 6].map((i) => row(`第${i}行`)).join('')}</w:tbl>`;
    const bytes = zipSync({
      'word/document.xml': strToU8(`<w:document ${W}><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="h1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800" w:header="851" w:footer="992"/></w:sectPr></w:body></w:document>`),
      'word/_rels/document.xml.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="h1" Target="header1.xml"/></Relationships>'),
      'word/header1.xml': strToU8(`<w:hdr ${W}>${tbl}<w:p/></w:hdr>`),
    });
    const d = readDocx(bytes);
    expect(d.lossy).not.toContain('页眉页脚中的表格');
    expect(d.header[0].hfTable?.table.rows).toHaveLength(6);
    const page = layoutDocument(d, fake).pages[0];
    const cellLines = page.header.filter((l) => l.ro);
    expect(cellLines).toHaveLength(6);
    // 六行表格 + 一个空段落高过上边距 72 磅：正文从页眉底下开始
    const bottom = Math.max(...page.header.map((l) => l.y + l.line.height));
    expect(page.body[0].y).toBeGreaterThanOrEqual(bottom - 0.01);
    expect(page.body[0].y).toBeGreaterThan(d.setup.marginTop);
    // 整篇重写时表格照写
    expect(strFromU8(unzipSync(writeDocx(d))['word/header1.xml'] ?? new Uint8Array())).toContain('<w:tbl>');
  });
});

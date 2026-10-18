// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { readDocx } from '../ooxml/docx';
import { layoutDocument, type Measurer } from './layout';
import { isCJK } from './text';

/** 假字体：汉字 1 em，西文 0.5 em */
const fake: Measurer = {
  width: (ch) => (isCJK(ch) ? 1 : 0.5),
  metrics: () => ({ ascent: 0.88, descent: 0.12 }),
};

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** n 段正文，第 at 段末尾引用脚注 1、第 at2 段引用脚注 2 */
function docx(n: number, at: number, at2 = -1): Uint8Array {
  const ref = (id: number) => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
  const body = Array.from({ length: n }, (_, i) => `<w:p><w:r><w:t>第${i + 1}段</w:t></w:r>${i === at ? ref(1) : ''}${i === at2 ? ref(2) : ''}</w:p>`).join('');
  const note = (id: number, t: string) => `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${t}</w:t></w:r></w:p></w:footnote>`;
  return zipSync({
    'word/document.xml': strToU8(`<w:document ${NS}><w:body>${body}</w:body></w:document>`),
    'word/_rels/document.xml.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rF" Target="footnotes.xml"/></Relationships>'),
    'word/footnotes.xml': strToU8(
      `<w:footnotes ${NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>${note(2, '第二条脚注')}${note(1, '第一条脚注')}</w:footnotes>`,
    ),
  });
}

describe('脚注（footnotes.xml）', () => {
  it('引用显示为按出现顺序的上标编号，脚注排在引用所在页的页底、上面有分隔线', () => {
    const d = readDocx(docx(3, 0, 2));
    expect(Object.keys(d.footnotes ?? {}).sort()).toEqual(['1', '2']);
    const page = layoutDocument(d, fake).pages[0];
    const ref = page.body[0].line.glyphs[page.body[0].line.glyphs.length - 1];
    expect(ref.ch).toBe('1');
    expect(ref.props.vertAlign).toBe('superscript');
    expect(page.noteIds).toEqual(['1', '2']);
    const texts = page.notes!.map((l) => l.line.glyphs.map((g) => g.ch).join(''));
    expect(texts).toEqual(['1第一条脚注', '2第二条脚注']);
    // 脚注区贴着版心底边
    const last = page.notes![page.notes!.length - 1];
    expect(last.y + last.line.height).toBeCloseTo(d.setup.height - d.setup.marginBottom);
    // 分隔线在第一条脚注上面
    expect(page.borders.some((b) => b.y1 === b.y2 && b.y1 < page.notes![0].y)).toBe(true);
  });

  it('脚注区占掉的高度从正文里扣：同样多的段落，有脚注时第一页放的行更少', () => {
    const count = (bytes: Uint8Array) => layoutDocument(readDocx(bytes), fake).pages[0].body.length;
    expect(count(docx(100, 0))).toBeLessThan(count(docx(100, -1)));
  });
});

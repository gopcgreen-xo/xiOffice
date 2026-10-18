// @vitest-environment jsdom
/**
 * 字节级保留保存（W12–W14）：不改动 = 原文件字节；只改一段 = 其它部件和其它段落的字节不变；
 * 改过的段落引用原样式、保留公式 / 书签 / 超链接；新图片、新编号、批注写到对应部件。
 */
import { readdirSync, readFileSync, statSync } from 'fs';
import { join } from 'path';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { openDocx, patchDocx } from './patch';
import { readDocx } from './docx';
import { addComment, applyCellBorders, applyNumbering, getPara, insertImage, insertRaw, paragraphPaths, setCellFormat } from '../core/ops';
import { paraText, type Paragraph } from '../core/model';
import { registerImage } from '../core/images';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

/** 最小 docx：document.xml + styles.xml（中文 Word 风格的数字样式 id）+ 关系 */
function pkg(body: string, extra: Record<string, string> = {}): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>',
    ),
    'word/styles.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>` +
        '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style></w:styles>',
    ),
    'word/document.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document ${W} ${R} ${M}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    ),
  };
  for (const [k, v] of Object.entries(extra)) files[k] = strToU8(v);
  return zipSync(files);
}

const docXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['word/document.xml']);
const first = (xml: string, re: RegExp) => xml.match(re)?.[0] ?? '';

describe('字节级保留保存', () => {
  it('不改动：保存结果就是原文件的字节', async () => {
    const bytes = pkg('<w:p><w:r><w:t>甲</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    expect(origin).not.toBeNull();
    const out = await patchDocx(doc, origin!);
    expect(out).toEqual(bytes);
  });

  it('只改一段：其它段落、其它部件逐字节不变；改过的段落沿用原 pPr 和样式引用', async () => {
    const p1 = '<w:p w14:paraId="1A2B3C4D" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:pPr><w:pStyle w:val="1"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr></w:pPr><w:r><w:t>第一章</w:t></w:r></w:p>';
    // 第二段故意写成“非常规”的样子：属性顺序、实体、自闭合形式，未改时必须原样保留
    const p2 = '<w:p><w:pPr><w:jc  w:val="center" /></w:pPr><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">A &amp; B &#x4E2D;</w:t></w:r></w:p>';
    const bytes = pkg(p1 + p2);
    const { doc, origin } = await openDocx(bytes);
    insertRaw(doc, { story: 'body', path: [0], offset: 3 }, '：总则');
    const out = await patchDocx(doc, origin!);
    const a = unzipSync(bytes);
    const b = unzipSync(out);
    for (const name of Object.keys(a)) if (name !== 'word/document.xml' && !name.startsWith('docProps/')) expect(strFromU8(b[name]), name).toBe(strFromU8(a[name]));
    const xml = docXml(out);
    expect(xml).toContain(p2);
    const edited = first(xml, /<w:p w14:paraId[\s\S]*?<\/w:p>/);
    // 原 pPr（样式引用、段落边框）照抄；文字的格式与样式相同，不写 rPr
    expect(edited).toContain('<w:pPr><w:pStyle w:val="1"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr></w:pPr>');
    expect(edited).toContain('第一章：总则');
    expect(edited).not.toContain('<w:rPr>');
    expect(paraText(readDocx(out).body[0] as Paragraph)).toBe('第一章：总则');
  });

  it('改格式：只写与样式不同的项；原 pPr 里不认识的子元素保留，并按规定顺序排列', async () => {
    const bytes = pkg('<w:p><w:pPr><w:pStyle w:val="1"/><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:textDirection w:val="lrTb"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    const p = doc.body[0] as Paragraph;
    p.props.align = 'center';
    p.runs[0].props = { ...p.runs[0].props, italic: true };
    const xml = docXml(await patchDocx(doc, origin!));
    const pPr = first(xml, /<w:pPr>[\s\S]*?<\/w:pPr>/);
    expect(pPr).toBe('<w:pPr><w:pStyle w:val="1"/><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:jc w:val="center"/><w:textDirection w:val="lrTb"/></w:pPr>');
    // 加粗、字号来自样式：不重复写；只写斜体
    expect(first(xml, /<w:r>[\s\S]*?<\/w:r>/)).toBe('<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">标题</w:t></w:r>');
  });

  it('改过的段落里的公式、书签、超链接原样保留', async () => {
    const math = '<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>';
    const bytes = pkg(`<w:p><w:bookmarkStart w:id="0" w:name="甲"/><w:r><w:t>见</w:t></w:r><w:hyperlink r:id="rId9"><w:r><w:t>链接</w:t></w:r></w:hyperlink>${math}<w:bookmarkEnd w:id="0"/></w:p>`);
    const { doc, origin } = await openDocx(bytes);
    const p = doc.body[0] as Paragraph;
    expect(p.runs.some((r) => r.opaque?.kind === 'math')).toBe(true);
    // 位置 0 是书签开始标记，1 是“见”
    insertRaw(doc, { story: 'body', path: [0], offset: 2 }, '下面');
    const xml = docXml(await patchDocx(doc, origin!));
    expect(xml).toMatch(/<w:bookmarkStart w:id="0" w:name="甲"\/>/);
    expect(xml).toMatch(/<w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">链接<\/w:t><\/w:r><\/w:hyperlink>/);
    expect(xml).toMatch(/<m:oMath[^>]*><m:r><m:t>x=1<\/m:t><\/m:r><\/m:oMath>/);
    expect(xml).toMatch(/<w:bookmarkEnd w:id="0"\/>/);
    expect(xml).toContain('见下面');
  });

  it('表格结构没变：只替换改过的单元格，表格属性和其它单元格照抄', async () => {
    const tbl =
      '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>数量</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const bytes = pkg(tbl + '<w:p/>');
    const { doc, origin } = await openDocx(bytes);
    insertRaw(doc, { story: 'body', path: [0, 0, 1, 0], offset: 2 }, '（件）');
    const xml = docXml(await patchDocx(doc, origin!));
    expect(xml).toContain('<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr>');
    expect(xml).toContain('<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc>');
    expect(xml).toContain('<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">数量（件）</w:t></w:r></w:p></w:tc>');
  });

  it('只改单元格格式（垂直对齐、框线、底纹）：只换 tcPr 里对应的子元素，tcPr 其它子元素、单元格内容和其它单元格逐字节不变', async () => {
    const tbl =
      '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/><w:noWrap/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>名称</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>数量</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const bytes = pkg(tbl + '<w:p><w:r><w:t>后</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    const c00 = { story: 'body' as const, path: [0, 0, 0, 0], offset: 0 };
    const c01 = { story: 'body' as const, path: [0, 0, 1, 0], offset: 0 };
    setCellFormat(doc, c00, c00, { vAlign: 'center' });
    applyCellBorders(doc, c01, c01, 'bottom', { w: 1, color: 'FF0000', style: 'dashed' });
    const xml = docXml(await patchDocx(doc, origin!));
    expect(xml).toContain(
      '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/><w:noWrap/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>名称</w:t></w:r></w:p></w:tc>',
    );
    expect(xml).toContain('<w:tc><w:tcPr><w:tcBorders><w:bottom w:val="dashed" w:sz="8" w:space="0" w:color="FF0000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>数量</w:t></w:r></w:p></w:tc>');
    expect(xml).toContain('<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>');
    expect(xml).toContain('<w:p><w:r><w:t>后</w:t></w:r></w:p>');
    // 去掉底纹：shd 删掉，其余不动
    setCellFormat(doc, c00, c00, { shading: null });
    const xml2 = docXml(await patchDocx(doc, origin!));
    expect(xml2).toContain('<w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:noWrap/><w:vAlign w:val="center"/></w:tcPr>');
  });

  it('新图片、新编号、新批注：补上媒体文件、关系、编号定义和批注部件', async () => {
    const bytes = pkg('<w:p><w:r><w:t>一</w:t></w:r></w:p><w:p><w:r><w:t>二</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    const png = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 15, 4, 0, 9, 251, 3, 253, 227, 85, 242, 156, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);
    insertImage(doc, { story: 'body', path: [0], offset: 1 }, { id: registerImage(png, 'image/png'), w: 20, h: 20 });
    applyNumbering(doc, { story: 'body', path: [1], offset: 0 }, { story: 'body', path: [1], offset: 0 }, 'decimal');
    addComment(doc, { story: 'body', path: [1], offset: 0 }, { story: 'body', path: [1], offset: 1 }, '确认数量', '张三');
    const out = await patchDocx(doc, origin!);
    const z = unzipSync(out);
    const rels = strFromU8(z['word/_rels/document.xml.rels']);
    const media = Object.keys(z).filter((n) => n.startsWith('word/media/'));
    expect(media.length).toBe(1);
    expect(rels).toContain(`Target="${media[0].slice(5)}"`);
    expect(strFromU8(z['[Content_Types].xml'])).toMatch(/Extension="png"/);
    expect(strFromU8(z['word/numbering.xml'])).toMatch(/<w:abstractNum w:abstractNumId="0">[\s\S]*<w:num w:numId="\d+"><w:abstractNumId w:val="0"\/><\/w:num>/);
    expect(strFromU8(z['word/comments.xml'])).toContain('确认数量');
    // 原来的超链接关系还在
    expect(rels).toContain('Id="rId9"');
    const back = readDocx(out);
    expect(back.comments?.[0]?.text).toBe('确认数量');
    expect((back.body[1] as Paragraph).props.num).not.toBeNull();
  });

  describe('样本：打开即保存 = 原文件字节；改第一段后其它部件不变', () => {
    const dir = join(__dirname, '../../samples');
    const files = (d: string): string[] => readdirSync(d).flatMap((n) => (statSync(join(d, n)).isDirectory() ? files(join(d, n)) : /\.docx$/i.test(n) ? [join(d, n)] : []));
    it.each(files(dir))('%s', async (file) => {
      const bytes = new Uint8Array(readFileSync(file));
      const { doc, origin } = await openDocx(bytes);
      expect(origin).not.toBeNull();
      expect(await patchDocx(doc, origin!)).toEqual(bytes);
      const paths = paragraphPaths(doc, 'body');
      if (!paths.length) return;
      insertRaw(doc, { story: 'body', path: paths[0], offset: 0 }, '★');
      const out = await patchDocx(doc, origin!);
      const a = unzipSync(bytes);
      const b = unzipSync(out);
      for (const name of Object.keys(a)) if (name !== 'word/document.xml' && !name.startsWith('docProps/')) expect(strFromU8(b[name]), name).toBe(strFromU8(a[name]));
      const again = readDocx(out);
      expect(paraText(getPara(again, { story: 'body', path: paragraphPaths(again, 'body')[0] }))).toMatch(/^★/);
    });
  });
});

describe('墨迹（w14:contentPart，Word 带替代图片）', () => {
  it('用 Fallback 里的图片显示；不改动时字节不变，改了同段文字后墨迹原样写回', async () => {
    const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    const inline = (graphic: string) => `<w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><wp:docPr id="1" name="墨迹 1"/><a:graphic><a:graphicData uri="x">${graphic}</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const ink =
      `<mc:AlternateContent ${WP}><mc:Choice Requires="w14">${inline('<w14:contentPart r:id="rIdInk"/>')}</mc:Choice>` +
      `<mc:Fallback>${inline('<pic:pic><pic:blipFill><a:blip r:embed="rIdPng"/></pic:blipFill></pic:pic>')}</mc:Fallback></mc:AlternateContent>`;
    const rels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rIdInk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="ink/ink1.xml"/><Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>';
    const bytes = pkg(`<w:p><w:r><w:t>签字：</w:t></w:r><w:r>${ink}</w:r></w:p>`, {
      'word/_rels/document.xml.rels': rels,
      'word/ink/ink1.xml': '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:trace>0 0, 10 10</inkml:trace></inkml:ink>',
      'word/media/image1.png': 'png',
    });
    const { doc, origin } = await openDocx(bytes);
    const p = doc.body[0] as Paragraph;
    const img = p.runs.find((r) => r.image)?.image;
    expect(img).toMatchObject({ w: 100, h: 50 });
    expect(img?.raw).toContain('w14:contentPart');
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
    insertRaw(doc, { story: 'body', path: [0], offset: 0 }, '乙方');
    const out = await patchDocx(doc, origin!);
    const xml = docXml(out);
    expect(xml).toContain('乙方签字：');
    expect(xml).toContain('<w14:contentPart r:id="rIdInk"/>');
    expect(Object.keys(unzipSync(out))).toContain('word/ink/ink1.xml');
  });
});

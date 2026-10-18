// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { strToU8, unzipSync, zipSync, strFromU8 } from 'fflate';
import { readDocx, writeDocx } from './docx';
import { imageData, registerImage } from '../core/images';
import { FIELD_CHAR, newDocument, newParagraph, newTable, paraText, textWidth, type Paragraph } from '../core/model';

describe('docx 读写', () => {
  it('写出后再读回，版式属性保持一致', () => {
    const d = newDocument();
    d.setup.grid = { type: 'linesAndChars', linePitch: 15.6, charSpace: -1024 };
    d.setup.marginLeft = 80;
    const title = newParagraph('办公用品采购合同', { align: 'center', spaceAfterLines: 0.5 }, { ...d.defaults, fontEA: '黑体', size: 22, bold: true });
    const body = newParagraph('甲方：[采购单位]\t乙方\n第二行\f', { firstLineChars: 2, line: { rule: 'exact', value: 28 } });
    const t = newTable(3, 2, textWidth(d.setup), d.defaults);
    t.rows[0].height = 30;
    t.rows[0].heightRule = 'exact';
    t.rows[0].header = true;
    t.rows[1].cells[0].gridSpan = 2;
    t.rows[1].cells.splice(1, 1);
    t.rows[1].cells[0].paragraphs[0].runs = [{ text: '合并单元格 & <符号>', props: d.defaults }];
    t.rows[1].cells[1].borders = false;
    d.body = [title, body, t, newParagraph('')];
    d.footer[0].runs = [
      { text: '第 ', props: d.defaults },
      { text: FIELD_CHAR, props: d.defaults, field: 'PAGE' },
      { text: ' 页', props: d.defaults },
    ];

    const back = readDocx(writeDocx(d));
    expect(back.lossy).toEqual([]);
    expect(back.setup.grid).toEqual(d.setup.grid);
    expect(back.setup.marginLeft).toBeCloseTo(80);
    expect(back.setup.width).toBeCloseTo(595.3, 0);

    const [p0, p1, tb] = back.body;
    if (p0.kind !== 'p' || p1.kind !== 'p' || tb.kind !== 'table') throw new Error('结构不对');
    expect(paraText(p0)).toBe('办公用品采购合同');
    expect(p0.props.align).toBe('center');
    expect(p0.props.spaceAfterLines).toBe(0.5);
    expect(p0.runs[0].props).toMatchObject({ fontEA: '黑体', size: 22, bold: true });
    expect(paraText(p1)).toBe('甲方：[采购单位]\t乙方\n第二行\f');
    expect(p1.props.firstLineChars).toBe(2);
    expect(p1.props.line).toEqual({ rule: 'exact', value: 28 });
    expect(tb.cols.map((c) => Math.round(c))).toEqual(t.cols.map((c) => Math.round(c)));
    expect(tb.rows[0]).toMatchObject({ height: 30, heightRule: 'exact', header: true });
    expect(tb.rows[1].cells[0].gridSpan).toBe(2);
    expect(paraText(tb.rows[1].cells[0].paragraphs[0])).toBe('合并单元格 & <符号>');
    expect(tb.borders).toBe(true);
    expect(tb.rows[1].cells[1].borders).toBe(false);
    expect(tb.rows[0].cells[0].borders).toBeUndefined();
    expect(back.footer[0].runs.map((r) => r.field ?? r.text)).toEqual(['第 ', 'PAGE', ' 页']);
  });

  it('读取 Word 的样式继承、复杂域和不支持的内容', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const styles = `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="a"><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="1"><w:basedOn w:val="a"/><w:pPr><w:keepNext/><w:spacing w:before="340" w:after="330" w:line="578" w:lineRule="auto"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style></w:styles>`;
    const doc = `<w:document ${W}><w:body>
      <w:p><w:pPr><w:pStyle w:val="1"/></w:pPr><w:r><w:t>一、总则</w:t></w:r></w:p>
      <w:p><w:r><w:t xml:space="preserve">共 </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:drawing/></w:r></w:p>
      <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2098" w:right="1474" w:bottom="1984" w:left="1588" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="579"/></w:sectPr></w:body></w:document>`;
    const bytes = zipSync({ 'word/document.xml': strToU8(doc), 'word/styles.xml': strToU8(styles) });
    const d = readDocx(bytes);
    const [h, p] = d.body as Paragraph[];
    expect(h.props).toMatchObject({ keepNext: true, widowControl: false, align: 'both', spaceBefore: 17 });
    expect(h.props.line.value).toBeCloseTo(578 / 240);
    expect(h.runs[0].props).toMatchObject({ bold: true, size: 22, fontEA: '宋体' });
    // 空的 w:drawing 不认识：作为原样对象保留（不引用别的部件，整篇重写也能写回）
    expect(p.runs.filter((r) => !r.opaque).map((r) => r.field ?? r.text)).toEqual(['共 ', 'NUMPAGES']);
    expect(p.runs.at(-1)?.opaque?.kind).toBe('drawing');
    // 公文常用版心：上 37mm 左 28mm，行距 28.95 磅
    expect(d.setup.marginTop).toBeCloseTo(104.9);
    expect(d.setup.grid).toMatchObject({ type: 'lines', linePitch: 28.95 });
  });

  it('嵌入式图片：写出 media 文件和关系，读回后仍是同一张图片', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4, 5]);
    const id = registerImage(png, 'image/png');
    const d = newDocument();
    const p = newParagraph('图：');
    p.runs.push({ text: FIELD_CHAR, props: d.defaults, image: { id, w: 120, h: 45.5 } });
    // 同一张图片用两次只写一个文件
    const q = newParagraph('');
    q.runs.push({ text: FIELD_CHAR, props: d.defaults, image: { id, w: 60, h: 22.75 } });
    d.body = [p, q];

    const bytes = writeDocx(d);
    const files = unzipSync(bytes);
    expect(Object.keys(files).filter((f) => f.startsWith('word/media/'))).toEqual(['word/media/image1.png']);
    expect(Array.from(files['word/media/image1.png'])).toEqual(Array.from(png));
    expect(strFromU8(files['[Content_Types].xml'])).toContain('Extension="png" ContentType="image/png"');
    expect(strFromU8(files['word/_rels/document.xml.rels'])).toContain('Target="media/image1.png"');

    const back = readDocx(bytes);
    expect(back.lossy).toEqual([]);
    const [p0, p1] = back.body as Paragraph[];
    expect(paraText(p0)).toBe(`图：${FIELD_CHAR}`);
    expect(p0.runs[1].image).toMatchObject({ id, w: 120, h: 45.5 });
    expect(p1.runs[0].image).toMatchObject({ id, w: 60, h: 22.75 });
    // 读入的图片记下原始 XML（字节级保留保存时原样写回）
    expect(p0.runs[1].image?.raw).toMatch(/^<w:drawing/);
    expect(imageData(id)?.mime).toBe('image/png');
  });

  it('引用了其它部件、本编辑器又显示不了的图片：原样保留，只有整篇重写时才算会丢失', () => {
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const drawing = (kind: string, rid: string) =>
      `<w:p><w:r><w:drawing><wp:${kind}><wp:extent cx="1270000" cy="635000"/><a:graphic><a:graphicData><a:blip r:embed="${rid}"/></a:graphicData></a:graphic></wp:${kind}></w:drawing></w:r></w:p>`;
    const rels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Target="media/a.png"/><Relationship Id="r2" Target="media/b.emf"/></Relationships>';
    const read = (body: string) =>
      readDocx(
        zipSync({
          'word/document.xml': strToU8(`<w:document ${NS}><w:body>${body}</w:body></w:document>`),
          'word/_rels/document.xml.rels': strToU8(rels),
          'word/media/a.png': new Uint8Array([1, 2, 3]),
          'word/media/b.emf': new Uint8Array([4, 5, 6]),
        }),
      );
    const ok = read(drawing('inline', 'r1'));
    expect(ok.lossy).toEqual([]);
    expect((ok.body[0] as Paragraph).runs[0].image).toMatchObject({ w: 100, h: 50 });
    expect(read(drawing('anchor', 'r1')).lossy).toContain('图片、文本框或图形');
    const emf = read(drawing('inline', 'r2'));
    expect(emf.lossy).toContain('图片、文本框或图形');
    expect((emf.body[0] as Paragraph).runs[0].opaque).toMatchObject({ kind: 'drawing', level: 'run', w: 100, h: 50 });
  });

  it('浮动图片（公章）：位置、大小、衬于文字下方读进来，保存后再读一致', () => {
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const anchor =
      '<w:p><w:r><w:t>甲方（盖章）</w:t></w:r><w:r><w:drawing><wp:anchor behindDoc="1" relativeHeight="5"><wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="column"><wp:posOffset>635000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>-127000</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="1524000" cy="1524000"/><wp:wrapNone/><wp:docPr id="1" name="公章"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="r1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>';
    const png = new Uint8Array([137, 80, 78, 71, 9, 9, 9]);
    const d = readDocx(
      zipSync({
        'word/document.xml': strToU8(`<w:document ${NS}><w:body>${anchor}</w:body></w:document>`),
        'word/_rels/document.xml.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Target="media/seal.png"/></Relationships>'),
        'word/media/seal.png': png,
      }),
    );
    expect(d.lossy).toEqual([]);
    const box = (d.body[0] as Paragraph).runs[1].box!;
    expect(box).toMatchObject({ w: 120, h: 120, behind: true, wrap: 'none', posH: { rel: 'column', offset: 50 }, posV: { rel: 'paragraph', offset: -10 } });
    expect(Array.from(imageData(box.image!.id)!.data)).toEqual(Array.from(png));
    // 拖动后保存：新位置写进文件
    box.posH = { rel: 'page', offset: 300 };
    const out = writeDocx(d);
    const back = readDocx(out);
    expect((back.body[0] as Paragraph).runs[1].box).toMatchObject({ w: 120, h: 120, behind: true, posH: { rel: 'page', offset: 300 } });
    expect(Object.keys(unzipSync(out)).filter((f) => f.startsWith('word/media/'))).toEqual(['word/media/image1.png']);
  });

  it('编号段落的缩进：样式 → 编号级别 → 段落自己，逐层覆盖（Word 的规则）', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    // “列表段落”样式首行缩进 2 字符；编号级别左缩进 18 磅、悬挂 18 磅
    const styles = `<w:styles ${W}><w:style w:type="paragraph" w:styleId="af"><w:pPr><w:ind w:firstLineChars="200" w:firstLine="420"/></w:pPr></w:style></w:styles>`;
    const numbering =
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="2">` +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/></w:lvl>' +
      '</w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>';
    const para = (ind: string, level = 0) =>
      `<w:p><w:pPr><w:pStyle w:val="af"/><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="3"/></w:numPr>${ind}</w:pPr><w:r><w:t>条款</w:t></w:r></w:p>`;
    const doc =
      `<w:document ${W}><w:body>` +
      para('<w:ind w:firstLineChars="0"/>') + // Word 常见写法：取消样式的字符缩进，用编号级别的悬挂缩进
      para('') + // 没有自己的缩进：样式的“首行 2 字符”还在，左缩进来自编号级别
      para('<w:ind w:left="720" w:firstLine="0"/>') + // 段落自己写的优先
      para('<w:ind w:firstLineChars="0"/>', 1) + // 这个级别没写缩进：不改样式的缩进
      '<w:p><w:pPr><w:pStyle w:val="af"/><w:ind w:firstLineChars="0"/></w:pPr><w:r><w:t>无编号</w:t></w:r></w:p>' +
      '</w:body></w:document>';
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc), 'word/styles.xml': strToU8(styles), 'word/numbering.xml': strToU8(numbering) }));
    const ind = (d.body as Paragraph[]).map((p) => [p.props.indentLeft, p.props.firstLine, p.props.firstLineChars]);
    expect(ind).toEqual([
      [18, -18, null],
      [18, -18, 2],
      [36, 0, null],
      [0, 21, null],
      // 字符数为 0 = 不按字符数缩进，样式的 21 磅首行缩进仍然有效
      [0, 21, null],
    ]);
  });

  it('表格边框粗细：读取 w:sz（1/8 磅），写出后不变', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const tbl = (sz: number) =>
      `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="${sz}"/><w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`;
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(`<w:document ${W}><w:body>${tbl(12)}${tbl(4)}<w:p/></w:body></w:document>`) }));
    const sizes = (x: typeof d) => x.body.filter((b) => b.kind === 'table').map((b) => (b.kind === 'table' ? (b.borderSize ?? 0.5) : 0));
    expect(sizes(d)).toEqual([1.5, 0.5]);
    const bytes = writeDocx(d);
    expect(strFromU8(unzipSync(bytes)['word/document.xml'])).toContain('w:sz="12"');
    expect(sizes(readDocx(bytes))).toEqual([1.5, 0.5]);
  });

  it('符号（w:sym）保留符号字体，写出后仍是 w:sym；编号级别自带的字符格式', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const numbering =
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1．"/>` +
      '<w:rPr><w:rFonts w:hint="eastAsia"/><w:b/><w:sz w:val="24"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
    const doc =
      `<w:document ${W}><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
      '<w:r><w:rPr><w:rFonts w:ascii="SimSun" w:eastAsia="宋体"/></w:rPr><w:t>运输：</w:t></w:r>' +
      '<w:r><w:rPr><w:rFonts w:ascii="SimSun" w:eastAsia="宋体"/></w:rPr><w:sym w:font="Wingdings 2" w:char="00A3"/></w:r>' +
      '<w:r><w:rPr><w:rFonts w:ascii="SimSun" w:eastAsia="宋体"/></w:rPr><w:t>火车</w:t></w:r>' +
      '<w:r><w:sym w:font="Wingdings 2" w:char="F052"/></w:r>' +
      '</w:p></w:body></w:document>';
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc), 'word/numbering.xml': strToU8(numbering) }));
    const p = d.body[0] as Paragraph;
    // 字符放在私用区 U+F0xx，字体是符号字体；中文字体不变（接着输入汉字不受影响）
    expect(p.runs.map((r) => [r.text, r.props.fontLatin])).toEqual([
      ['运输：', 'SimSun'],
      [String.fromCharCode(0xf0a3), 'Wingdings 2'],
      ['火车', 'SimSun'],
      [String.fromCharCode(0xf052), 'Wingdings 2'],
    ]);
    expect(p.runs[1].props.fontEA).toBe('宋体');
    expect(d.lists[0].levels[0].label).toEqual({ bold: true, size: 12 });

    const files = unzipSync(writeDocx(d));
    const xml = strFromU8(files['word/document.xml']);
    expect(xml).toContain('<w:sym w:font="Wingdings 2" w:char="F0A3"/>');
    expect(xml).toContain('<w:sym w:font="Wingdings 2" w:char="F052"/>');
    expect(strFromU8(files['word/numbering.xml'])).toContain('<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>');
    const back = readDocx(writeDocx(d));
    expect((back.body[0] as Paragraph).runs.map((r) => r.text)).toEqual(p.runs.map((r) => r.text));
    expect(back.lists[0].levels[0].label).toEqual({ bold: true, size: 12 });
  });

  it('文本框（mc:AlternateContent 里的浮动绘图）：读取位置、边框和文字，保存后再读回不变', () => {
    const NS =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"';
    const drawing =
      '<w:drawing><wp:anchor behindDoc="0"><wp:positionH relativeFrom="column"><wp:posOffset>1270000</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>25400</wp:posOffset></wp:positionV><wp:extent cx="2540000" cy="1270000"/><wp:wrapNone/>' +
      '<a:graphic><a:graphicData><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:solidFill><a:srgbClr val="ffffff"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
      '<wps:txbx><w:txbxContent><w:p><w:r><w:t>合同号：A1</w:t></w:r></w:p><w:p><w:r><w:t>收货人：采购部</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
      '<wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>';
    const doc =
      `<w:document ${NS}><w:body><w:p><w:r><w:t>包装标识：</w:t></w:r>` +
      `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${drawing}</mc:Choice><mc:Fallback><w:pict><v:shape/></w:pict></mc:Fallback></mc:AlternateContent></w:r>` +
      '</w:p></w:body></w:document>';
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc) }));
    // 读了新格式就不再读旧格式（Fallback），也不再提示
    expect(d.lossy).toEqual([]);
    const p = d.body[0] as Paragraph;
    expect(paraText(p)).toBe(`包装标识：${FIELD_CHAR}`);
    const box = p.runs[1].box!;
    expect(box).toMatchObject({
      w: 200,
      h: 100,
      posH: { rel: 'column', offset: 100 },
      posV: { rel: 'paragraph', offset: 2 },
      wrap: 'none',
      behind: false,
      fill: 'FFFFFF',
      line: { color: '000000', w: 0.75 },
      inset: { l: 7.2, t: 3.6, r: 7.2, b: 3.6 },
      vAlign: 'top',
    });
    expect(box.paragraphs.map(paraText)).toEqual(['合同号：A1', '收货人：采购部']);

    const bytes = writeDocx(d);
    const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
    expect(xml).toContain('<mc:Choice Requires="wps">');
    expect(xml).toContain('<wp:posOffset>1270000</wp:posOffset>');
    const back = readDocx(bytes);
    expect(back.lossy).toEqual([]);
    const again = (back.body[0] as Paragraph).runs[1].box!;
    expect({ ...again, paragraphs: [], raw: undefined }).toEqual({ ...box, paragraphs: [], raw: undefined });
    expect(again.paragraphs.map(paraText)).toEqual(['合同号：A1', '收货人：采购部']);
  });

  it('没有文字内容的形状：原样保留，整篇重写时也写回', () => {
    const NS =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
    const doc = `<w:document ${NS}><w:body><w:p><w:r><w:drawing><wp:anchor><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData><wps:wsp><wps:spPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:body></w:document>`;
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc) }));
    expect(d.lossy).toEqual([]);
    expect((d.body[0] as Paragraph).runs[0].opaque?.kind).toBe('drawing');
    expect(strFromU8(unzipSync(writeDocx(d))['word/document.xml'])).toContain('<wps:wsp><wps:spPr/></wps:wsp>');
  });

  it('形状：外形、翻转、旋转；spPr 没写的填充 / 线条用形状样式（fillRef / lnRef）和主题颜色、主题线宽', () => {
    const NS =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
    const theme =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
      '<a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme><a:fontScheme name="x"/><a:fmtScheme name="x"><a:lnStyleLst><a:ln w="6350"/><a:ln w="12700"/><a:ln w="19050"/></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>';
    const shape = (spPr: string, style: string) =>
      `<w:r><w:drawing><wp:anchor behindDoc="1"><wp:positionH relativeFrom="page"><wp:posOffset>127000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>254000</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="1270000" cy="635000"/><wp:wrapNone/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>` +
      `<wps:spPr>${spPr}</wps:spPr><wps:style>${style}</wps:style><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const styleRefs = '<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>';
    const doc =
      `<w:document ${NS}><w:body><w:p>` +
      shape('<a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>', styleRefs) +
      // spPr 写了 noFill / 线条颜色：以 spPr 为准，线宽没写用主题线宽
      shape('<a:prstGeom prst="diamond"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>', styleRefs) +
      // 自由曲线：不认识，原样保留
      shape('<a:custGeom><a:pathLst/></a:custGeom>', styleRefs) +
      '</w:p></w:body></w:document>';
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc), 'word/theme/theme1.xml': strToU8(theme) }));
    expect(d.lossy).toEqual([]);
    const runs = (d.body[0] as Paragraph).runs;
    expect(runs[0].box).toMatchObject({ geom: 'ellipse', w: 100, h: 50, rot: 90, flipH: true, behind: true, fill: '4472C4', line: { color: '2F528F', w: 1 }, paragraphs: [] });
    expect(runs[1].box).toMatchObject({ geom: 'diamond', fill: null, line: { color: '000000', w: 1 } });
    expect(runs[2].opaque?.kind).toBe('drawing');

    // 整篇写出再读回：外形、翻转、旋转、颜色不变；没有文字的形状不写 wps:txbx
    const bytes = writeDocx(d);
    const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
    expect(xml).toContain('<a:prstGeom prst="ellipse">');
    expect(xml).toContain('<a:xfrm rot="5400000" flipH="1">');
    expect(xml).not.toContain('<wps:txbx>');
    const back = (readDocx(bytes).body[0] as Paragraph).runs;
    for (const i of [0, 1]) expect({ ...back[i].box, raw: undefined }).toEqual({ ...runs[i].box, raw: undefined, dist: back[i].box!.dist });
  });

  it('左缩进同时写了磅值和字符数时用磅值，只写字符数时按字号换算', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const para = (ind: string) => `<w:p><w:pPr><w:ind ${ind}/></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>条款</w:t></w:r></w:p>`;
    const doc = `<w:document ${W}><w:body>${para('w:leftChars="150" w:left="354" w:hangingChars="27" w:hanging="54"')}${para('w:leftChars="150"')}</w:body></w:document>`;
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc) }));
    const [a, b] = d.body as Paragraph[];
    expect(a.props.indentLeft).toBeCloseTo(17.7);
    expect(a.props.firstLineChars).toBeCloseTo(-0.27);
    expect(b.props.indentLeft).toBeCloseTo(15);
  });

  it('生成的包结构完整', () => {
    const files = unzipSync(writeDocx(newDocument()));
    expect(Object.keys(files).sort()).toEqual(
      ['[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels', 'word/document.xml', 'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml', 'word/header1.xml', 'word/header2.xml', 'word/header3.xml', 'word/settings.xml', 'word/styles.xml'].sort(),
    );
    expect(strFromU8(files['word/settings.xml'])).toContain('doNotCompress');
  });
});

describe('分栏 w:cols 读写', () => {
  it('各节的栏数、间距、分隔线、不等宽栏写出后读回一致；一栏的节没有 cols', () => {
    const d = newDocument();
    d.setup.cols = { num: 2, space: 21.25, sep: true };
    const end = newParagraph('第一节', { sectionBreak: 'continuous' });
    end.sectionCols = { num: 3, space: 18, sep: false, widths: [{ w: 100, space: 18 }, { w: 120, space: 18 }, { w: 150, space: 0 }] };
    const one = newParagraph('第二\v节', { sectionBreak: 'nextPage' });
    d.body = [end, one, newParagraph('最后一节')];
    const back = readDocx(writeDocx(d));
    expect(back.lossy).toEqual([]);
    expect(back.setup.cols).toEqual(d.setup.cols);
    const [b0, b1] = back.body as Paragraph[];
    expect(b0.sectionCols).toEqual(end.sectionCols);
    expect(b1.sectionCols).toBeUndefined();
    expect(paraText(b1)).toBe('第二\v节'); // 分栏符往返
    expect(strFromU8(unzipSync(writeDocx(d))['word/document.xml'])).toContain('<w:br w:type="column"/>');
  });
});

describe('浮动表格 w:tblpPr 读写', () => {
  it('位置、对齐、参照和环绕距离写出后读回一致', () => {
    const d = newDocument();
    const t = newTable(2, 1, 200, d.defaults);
    t.float = { hAnchor: 'margin', vAnchor: 'page', x: 0, y: 300, xAlign: 'center', dist: { l: 9, r: 9, t: 0, b: 6 } };
    d.body = [t, newParagraph('')];
    const back = readDocx(writeDocx(d));
    expect((back.body[0] as typeof t).float).toEqual(t.float);
    expect(back.lossy).toEqual([]);
  });
});

describe('兼容设置、字符间距、自动段距、缺省的页眉页脚', () => {
  it('兼容模式与 useAltKinsokuLineBreakRules 读进来、写回去；缺省兼容模式 15', () => {
    const d = newDocument();
    Object.assign(d.setup, { compatMode: 11, altKinsoku: true, compressPunct: true });
    const files = unzipSync(writeDocx(d));
    const settings = strFromU8(files['word/settings.xml']);
    expect(settings).toContain('<w:useAltKinsokuLineBreakRules/>');
    expect(settings).toMatch(/compatibilityMode" w:uri="[^"]+" w:val="11"/);
    const back = readDocx(writeDocx(d));
    expect([back.setup.compatMode, back.setup.altKinsoku, back.setup.compressPunct]).toEqual([11, true, true]);
    expect(readDocx(writeDocx(newDocument())).setup.compatMode).toBe(15);
  });

  it('字符间距 w:spacing（rPr）在 color 与 sz 之间写出，读回一致', () => {
    const d = newDocument();
    d.body = [newParagraph('紧缩一点', {}, { ...d.defaults, spacing: -0.5, color: 'FF0000' })];
    const xml = strFromU8(unzipSync(writeDocx(d))['word/document.xml']);
    expect(xml).toMatch(/<w:color w:val="FF0000"\/><w:spacing w:val="-10"\/><w:sz /);
    const p = readDocx(writeDocx(d)).body[0] as Paragraph;
    expect(p.runs[0].props.spacing).toBe(-0.5);
  });

  it('段前 / 段后自动（beforeAutospacing / afterAutospacing）读进来、写回去', () => {
    const d = newDocument();
    d.body = [newParagraph('自动', { spaceBeforeAuto: true, spaceAfterAuto: true })];
    const xml = strFromU8(unzipSync(writeDocx(d))['word/document.xml']);
    expect(xml).toContain('w:beforeAutospacing="1"');
    expect(xml).toContain('w:afterAutospacing="1"');
    const p = readDocx(writeDocx(d)).body[0] as Paragraph;
    expect([p.props.spaceBeforeAuto, p.props.spaceAfterAuto]).toEqual([true, true]);
  });

  it('原文件没有页脚时标记为“没有”（不占正文位置）；存在的空页眉不标', () => {
    const files = unzipSync(writeDocx(newDocument()));
    const doc = strFromU8(files['word/document.xml']).replace(/<w:footerReference [^>]*\/>/g, '');
    const back = readDocx(zipSync({ ...files, 'word/document.xml': strToU8(doc) }));
    expect(back.footer.map((p) => !!p.hfNone)).toEqual([true]);
    expect(back.header.some((p) => p.hfNone)).toBe(false);
    // 再整篇写出：没有的页脚仍不写引用，有内容后才写
    expect(strFromU8(unzipSync(writeDocx(back))['word/document.xml'])).not.toContain('w:footerReference');
    back.footer[0].runs = [{ text: '第 1 页', props: back.defaults }];
    expect(strFromU8(unzipSync(writeDocx(back))['word/document.xml'])).toContain('w:footerReference w:type="default"');
  });
});

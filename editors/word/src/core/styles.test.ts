// @vitest-environment jsdom
/** 样式管理：套用样式（直接格式保留）、新建 / 更新样式并写回 styles.xml、整篇重写时写出自定义样式 */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { applyCharacterStyle, applyParagraphStyle, catalogFromDoc, checkStyleName, createStyle, deleteStyle, redefineStyle, renameStyle, styleLabel, withEdits } from './styles';
import { applyRunProps, getPara, insertRaw, setParaProps } from './ops';
import { newDocument } from './model';
import { openDocx, patchDocx } from '../ooxml/patch';
import { readDocx, styleSource, writeDocx } from '../ooxml/docx';
import { pkg } from '../ooxml/testPackage';

const at = (offset: number, i = 0) => ({ story: 'body' as const, path: [i], offset });
const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Body2"><w:name w:val="Body Indent"/><w:basedOn w:val="a"/><w:qFormat/><w:pPr><w:ind w:firstLineChars="200" w:firstLine="420"/></w:pPr></w:style>' +
  '<w:style w:type="character" w:styleId="Red"><w:name w:val="Strong Red"/><w:qFormat/><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr></w:style>';
const P2 = '<w:p><w:r><w:t>第二段  保持原样</w:t></w:r></w:p>';
const open = async () => openDocx(pkg(`<w:p><w:r><w:t>第一段文字</w:t></w:r></w:p>${P2}`, { styles: STYLES }));
const stylesXml = (b: Uint8Array) => strFromU8(unzipSync(b)['word/styles.xml']);
const docXml = (b: Uint8Array) => strFromU8(unzipSync(b)['word/document.xml']);

describe('样式库', () => {
  it('列出原文件的样式；中文显示名；新建样式查重', async () => {
    const { origin, doc } = await open();
    const cat = origin!.styles;
    const names = cat.list().map((s) => [s.id, styleLabel(s), s.type]);
    expect(names).toEqual(expect.arrayContaining([['a', '正文', 'paragraph'], ['1', '标题 1', 'paragraph'], ['Body2', 'Body Indent', 'paragraph'], ['Red', 'Strong Red', 'character']]));
    expect(cat.defaultId()).toBe('a');
    expect(checkStyleName(cat, doc, '标题 1')).not.toBeNull();
    expect(checkStyleName(cat, doc, 'body indent')).not.toBeNull();
    expect(checkStyleName(cat, doc, '')).not.toBeNull();
    expect(checkStyleName(cat, doc, '公文正文')).toBeNull();
  });
});

describe('套用样式', () => {
  it('套用标题 1：格式按样式，保存只写 pStyle；其它段落字节不变；读回仍是标题', async () => {
    const { doc, origin } = await open();
    applyParagraphStyle(doc, origin!.styles, at(0), at(0), '1');
    const p = getPara(doc, at(0));
    expect(p.props).toMatchObject({ styleId: '1', style: 'Heading1', keepNext: true, outline: 0 });
    expect(p.runs[0].props).toMatchObject({ bold: true, size: 22 });
    const out = await patchDocx(doc, origin!);
    const xml = docXml(out);
    expect(xml).toContain(P2);
    const first = xml.match(/<w:p>.*?<\/w:p>/)![0];
    expect(first).toContain('<w:pStyle w:val="1"/>');
    expect(first).not.toMatch(/<w:b\/>|<w:sz /); // 与样式相同的格式不重复写
    expect(readDocx(out).body[0]).toMatchObject({ props: { styleId: '1', style: 'Heading1' } });
  });

  it('直接格式保留：先设成红色再套样式，颜色仍是红色', async () => {
    const { doc, origin } = await open();
    applyRunProps(doc, at(0), at(5), { color: 'FF0000' });
    applyParagraphStyle(doc, origin!.styles, at(0), at(0), '1');
    const out = await patchDocx(doc, origin!);
    expect(docXml(out).match(/<w:p>.*?<\/w:p>/)![0]).toContain('<w:color w:val="FF0000"/>');
    expect(getPara(readDocx(out), at(0)).runs[0].props).toMatchObject({ color: 'FF0000', bold: true });
  });

  it('换回正文：标题的加粗、字号、保持下一段都去掉', async () => {
    const { doc, origin } = await open();
    applyParagraphStyle(doc, origin!.styles, at(0), at(0), '1');
    applyParagraphStyle(doc, origin!.styles, at(0), at(0), 'a');
    const p = getPara(doc, at(0));
    expect(p.props.keepNext).toBe(false);
    expect(p.props.style).toBeUndefined();
    expect(p.runs[0].props).toMatchObject({ bold: false, size: 10.5 });
  });

  it('字符样式：套用后文字按样式；写 rStyle；清除后还原', async () => {
    const { doc, origin } = await open();
    applyCharacterStyle(doc, origin!.styles, at(1), at(3), 'Red');
    expect(getPara(doc, at(0)).runs.map((r) => [r.text, r.props.bold, r.props.color, r.props.rStyle])).toEqual([
      ['第', false, 'auto', undefined],
      ['一段', true, 'FF0000', 'Red'],
      ['文字', false, 'auto', undefined],
    ]);
    const out = await patchDocx(doc, origin!);
    expect(docXml(out)).toContain('<w:rStyle w:val="Red"/>');
    applyCharacterStyle(doc, origin!.styles, at(1), at(3), null);
    expect(getPara(doc, at(0)).runs).toHaveLength(1);
  });
});

describe('新建 / 更新样式', () => {
  it('从选中段落新建样式：写进 styles.xml，段落引用它且不重复写格式；读回样式表里有它', async () => {
    const { doc, origin } = await open();
    applyRunProps(doc, at(0), at(5), { bold: true, size: 16, fontEA: '黑体' });
    setParaProps(doc, at(0), at(0), { align: 'center', spaceAfter: 12, line: { rule: 'exact', value: 28.95 } });
    const id = createStyle(doc, origin!.styles, at(0), '公文标题', 'paragraph');
    expect(id).toBe('Wz1');
    expect(getPara(doc, at(0)).props.styleId).toBe('Wz1');
    const out = await patchDocx(doc, origin!);
    const sx = stylesXml(out);
    const st = sx.match(/<w:style [^>]*w:styleId="Wz1"[\s\S]*?<\/w:style>/)![0];
    expect(st).toContain('w:customStyle="1"');
    expect(st).toContain('<w:name w:val="公文标题"/>');
    expect(st).toContain('<w:basedOn w:val="a"/>');
    expect(st).toContain('<w:jc w:val="center"/>');
    expect(st).toMatch(/w:line="579" w:lineRule="exact"/);
    expect(st).toContain('<w:sz w:val="32"/>');
    expect(st).toContain('w:eastAsia="黑体"');
    // 原有样式不动
    expect(sx).toContain('w:styleId="1"');
    const first = docXml(out).match(/<w:p>.*?<\/w:p>/)![0];
    expect(first).toContain('<w:pStyle w:val="Wz1"/>');
    expect(first).not.toContain('<w:jc'); // 与样式相同，不再写
    expect(docXml(out)).toContain(P2);
    // 读回：样式表里有，段落的格式与保存前一致
    const back = readDocx(out);
    expect(getPara(back, at(0)).props).toMatchObject({ styleId: 'Wz1', align: 'center', spaceAfter: 12, line: { rule: 'exact', value: 28.95 } });
    expect(getPara(back, at(0)).runs[0].props).toMatchObject({ bold: true, size: 16, fontEA: '黑体' });
    expect(styleSource(out).list().find((s) => s.id === 'Wz1')?.name).toBe('公文标题');
  });

  it('用选中内容更新样式：用到它的段落跟着变，其它项保持；styles.xml 里是原位更新', async () => {
    const { doc, origin } = await open();
    // 第二段也用标题 1，再把第一段改成“蓝色加大”后用它更新标题 1
    applyParagraphStyle(doc, origin!.styles, at(0, 1), at(0, 1), '1');
    applyParagraphStyle(doc, origin!.styles, at(0, 0), at(0, 0), '1');
    applyRunProps(doc, at(0, 0), at(5, 0), { color: '0000FF', size: 30 });
    redefineStyle(doc, origin!.styles, '1', at(1, 0));
    const second = getPara(doc, at(0, 1));
    expect(second.runs[0].props).toMatchObject({ color: '0000FF', size: 30, bold: true });
    const out = await patchDocx(doc, origin!);
    const st = stylesXml(out).match(/<w:style [^>]*w:styleId="1"[\s\S]*?<\/w:style>/)![0];
    expect(st).toContain('<w:name w:val="heading 1"/>');
    expect(st).toContain('<w:qFormat/>');
    expect(st).toContain('<w:next w:val="a"/>');
    expect(st).toContain('<w:sz w:val="60"/>');
    expect(st).toContain('<w:color w:val="0000FF"/>');
    expect(st).not.toContain('customStyle');
    const back = readDocx(out);
    expect(getPara(back, at(0, 1)).runs[0].props).toMatchObject({ color: '0000FF', size: 30 });
    expect(getPara(back, at(0, 0)).runs[0].props).toMatchObject({ color: '0000FF', size: 30 });
  });

  it('字符样式：从选中文字新建，只记与段落样式不同的项', async () => {
    const { doc, origin } = await open();
    applyRunProps(doc, at(1), at(3), { italic: true, color: '008000' });
    const id = createStyle(doc, origin!.styles, at(1), '强调绿', 'character', at(3));
    const out = await patchDocx(doc, origin!);
    const st = stylesXml(out).match(new RegExp(`<w:style [^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`))![0];
    expect(st).toContain('w:type="character"');
    expect(st).toContain('<w:i/>');
    expect(st).toContain('<w:color w:val="008000"/>');
    expect(st).not.toContain('<w:sz');
    expect(docXml(out)).toContain(`<w:rStyle w:val="${id}"/>`);
  });

  it('没改样式时 styles.xml 逐字节不变', async () => {
    const { doc, origin } = await open();
    insertRaw(doc, at(0), '加');
    const out = await patchDocx(doc, origin!);
    expect(stylesXml(out)).toBe(strFromU8(unzipSync(pkg('', { styles: STYLES }))['word/styles.xml']));
  });
});

describe('新建文档（整篇重写）', () => {
  it('自定义样式写进 styles.xml，段落和文字引用它；读回一致', () => {
    const d = newDocument();
    insertRaw(d, at(0), '标题文字');
    applyRunProps(d, at(0), at(2), { bold: true, size: 22 });
    setParaProps(d, at(0), at(0), { align: 'center' });
    const cat = catalogFromDoc(d);
    const id = createStyle(d, cat, at(0), '大标题', 'paragraph');
    const bytes = writeDocx(d);
    const st = stylesXml(bytes).match(new RegExp(`<w:style [^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`))![0];
    expect(st).toContain('<w:name w:val="大标题"/>');
    expect(docXml(bytes)).toContain(`<w:pStyle w:val="${id}"/>`);
    const back = readDocx(bytes);
    expect(getPara(back, at(0)).props).toMatchObject({ styleId: id, align: 'center' });
    expect(withEdits(catalogFromDoc(back), undefined).list().map((x) => x.id)).toEqual(['Normal', 'Heading1', 'Heading2', 'Heading3']);
  });
});

describe('样式：基于关系、改名、删除、新建文档的样式表', () => {
  const OPEN_STYLES =
    '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Base1"><w:name w:val="公文正文"/><w:basedOn w:val="a"/><w:qFormat/><w:pPr><w:spacing w:after="100"/></w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Kid1"><w:name w:val="公文附件"/><w:basedOn w:val="Base1"/><w:qFormat/><w:rPr><w:b/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Kid2"><w:name w:val="公文备注"/><w:basedOn w:val="Base1"/><w:qFormat/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>';
  const open2 = async () =>
    openDocx(pkg('<w:p><w:pPr><w:pStyle w:val="Kid1"/></w:pPr><w:r><w:t>附件文字</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Kid2"/></w:pPr><w:r><w:t>备注文字</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Base1"/></w:pPr><w:r><w:t>正文文字</w:t></w:r></w:p>', { styles: OPEN_STYLES }));

  it('改上一级样式：基于它的样式里继承来的项跟着变，自己单独设的项不动', async () => {
    const { doc, origin } = await open2();
    expect(getPara(doc, at(0, 0)).runs[0].props).toMatchObject({ size: 14, bold: true });
    expect(getPara(doc, at(0, 1)).runs[0].props).toMatchObject({ size: 10 });
    // 把“公文正文”改成 16 磅、段后 12 磅
    applyRunProps(doc, at(0, 2), at(4, 2), { size: 16 });
    setParaProps(doc, at(0, 2), at(0, 2), { spaceAfter: 12 });
    redefineStyle(doc, origin!.styles, 'Base1', at(2, 2));
    // 附件样式没单独设字号，跟着变成 16（加粗仍在）；备注样式自己设了 10 磅，不动
    expect(getPara(doc, at(0, 0)).runs[0].props).toMatchObject({ size: 16, bold: true });
    expect(getPara(doc, at(0, 0)).props.spaceAfter).toBe(12);
    expect(getPara(doc, at(0, 1)).runs[0].props.size).toBe(10);
    expect(getPara(doc, at(0, 1)).props.spaceAfter).toBe(12);
    // 保存：三个段落只写样式引用，没有多余的直接格式；读回字号与保存前一致
    const out = await patchDocx(doc, origin!);
    const xml = docXml(out);
    expect(xml.match(/<w:sz /g)).toBeNull();
    const back = readDocx(out);
    expect(getPara(back, at(0, 0)).runs[0].props.size).toBe(16);
    expect(getPara(back, at(0, 1)).runs[0].props.size).toBe(10);
  });

  it('改名：自定义样式可以，内置样式不行；重名不行；styles.xml 里只改名字', async () => {
    const { doc, origin } = await open2();
    renameStyle(doc, origin!.styles, 'Kid1', '附件说明');
    expect(() => renameStyle(doc, origin!.styles, 'Kid2', '附件说明')).toThrow(/同名/);
    expect(() => renameStyle(doc, origin!.styles, 'a', '别的')).toThrow(/内置/);
    const sx = stylesXml(await patchDocx(doc, origin!));
    const st = sx.match(/<w:style [^>]*w:styleId="Kid1"[\s\S]*?<\/w:style>/)![0];
    expect(st).toContain('<w:name w:val="附件说明"/>');
    expect(st).toContain('<w:basedOn w:val="Base1"/>');
    expect(st).toContain('<w:b/>');
  });

  it('删除：用到它的段落改用上一级样式（直接格式保留），基于它的样式改成基于它的上一级，styles.xml 里去掉', async () => {
    const { doc, origin } = await open2();
    deleteStyle(doc, origin!.styles, 'Base1');
    // 附件样式基于已删的“公文正文”→ 改基于“正文”(a)；“公文正文”的 14 磅不再继承，附件回到正文字号
    expect(getPara(doc, at(0, 2)).props.styleId).toBe('a');
    expect(getPara(doc, at(0, 2)).runs[0].props.size).toBe(10.5);
    const out = await patchDocx(doc, origin!);
    const sx = stylesXml(out);
    expect(sx).not.toContain('w:styleId="Base1"');
    expect(sx).toMatch(/w:styleId="Kid1"[\s\S]*?<w:basedOn w:val="a"\/>/);
    expect(() => deleteStyle(doc, origin!.styles, 'a')).toThrow(/内置/);
    const back = readDocx(out);
    expect(getPara(back, at(0, 2)).props.styleId).toBe('a');
    expect(styleSource(out).list().some((s) => s.id === 'Base1')).toBe(false);
    // 本次新建的样式删除：直接丢掉，不进 styles.xml
    const again = await open2();
    const id = createStyle(again.doc, again.origin!.styles, at(0, 2), '临时样式', 'paragraph');
    deleteStyle(again.doc, again.origin!.styles, id);
    expect(stylesXml(await patchDocx(again.doc, again.origin!))).not.toContain('临时样式');
    expect(again.doc.styleEdits ?? {}).toEqual({});
  });

  it('新建文档的样式表：有正文和标题 1–3；套用标题后格式与“常用”页的标题一致，保存成标题样式', () => {
    const d = newDocument();
    insertRaw(d, at(0), '一级标题');
    const cat = catalogFromDoc(d);
    expect(cat.list().map((s) => s.id)).toEqual(['Normal', 'Heading1', 'Heading2', 'Heading3']);
    applyParagraphStyle(d, cat, at(0), at(0), 'Heading1');
    const p = getPara(d, at(0));
    expect(p.props).toMatchObject({ style: 'Heading1', outline: 0, keepNext: true });
    expect(p.runs[0].props).toMatchObject({ bold: true, size: 22 });
    const bytes = writeDocx(d);
    expect(docXml(bytes)).toContain('<w:pStyle w:val="Heading1"/>');
    expect(stylesXml(bytes)).toContain('w:styleId="Heading1"');
    applyParagraphStyle(d, cat, at(0), at(0), 'Normal');
    expect(getPara(d, at(0)).props.style).toBeUndefined();
    expect(getPara(d, at(0)).runs[0].props.bold).toBe(false);
  });
});

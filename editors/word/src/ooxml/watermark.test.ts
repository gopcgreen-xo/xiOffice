// @vitest-environment jsdom
/** 文字水印：读（Word 的 v:textpath 艺术字）、加 / 改 / 去、与页眉编辑共存、新建文档整篇写出 */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { openDocx, patchDocx } from './patch';
import { readDocx, writeDocx } from './docx';
import { pkg, W, R } from './testPackage';
import { insertRaw } from '../core/ops';
import { newDocument } from '../core/model';
import { watermarkParagraphXml, WATERMARK_NS } from '../../vendor/docx-engine/src/watermark';

const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HDR_TYPE = '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
const SECT = '<w:headerReference w:type="default" r:id="rId20"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';
const header = (inner: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr ${W} ${R}${WATERMARK_NS}>${inner}</w:hdr>`;
const withHeader = (inner: string) =>
  pkg('<w:p><w:r><w:t>正文</w:t></w:r></w:p>', {
    sect: SECT,
    rels: `<Relationship Id="rId20" Type="${OD}/header" Target="header1.xml"/>`,
    types: HDR_TYPE,
    extra: { 'word/header1.xml': header(inner) },
  });
const headerXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['word/header1.xml']);

describe('文字水印', () => {
  it('读：Word 的艺术字水印读成 doc.watermark（文字、颜色、透明度、斜 / 平）', async () => {
    const wm = watermarkParagraphXml({ text: '机密', colorHex: 'FF0000', opacity: 0.3, diagonal: true, bold: true });
    const { doc } = await openDocx(withHeader(`${wm}<w:p><w:r><w:t>公司</w:t></w:r></w:p>`));
    expect(doc.watermark).toMatchObject({ text: '机密', colorHex: 'FF0000', opacity: 0.3, diagonal: true, bold: true });
    const flat = (await openDocx(withHeader(watermarkParagraphXml({ text: '草稿', diagonal: false })))).doc;
    expect(flat.watermark).toMatchObject({ text: '草稿', diagonal: false });
    expect((await openDocx(withHeader('<w:p/>'))).doc.watermark).toBeUndefined();
  });

  it('没改：保存结果就是原文件（有水印的页眉也一样）', async () => {
    const bytes = withHeader(`${watermarkParagraphXml('机密')}<w:p><w:r><w:t>公司</w:t></w:r></w:p>`);
    const { doc, origin } = await openDocx(bytes);
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
  });

  it('给有页眉的文档加水印：页眉里多一个水印段落，原有文字还在，正文不变', async () => {
    const bytes = withHeader('<w:p><w:r><w:t>公司名称</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    doc.watermark = { text: '内部资料', colorHex: 'C0C0C0', opacity: 0.4, diagonal: true };
    const out = await patchDocx(doc, origin!);
    expect(headerXml(out)).toContain('<v:textpath');
    expect(headerXml(out)).toContain('string="内部资料"');
    expect(headerXml(out)).toContain('公司名称');
    expect(strFromU8(unzipSync(out)['word/document.xml'])).toBe(strFromU8(unzipSync(bytes)['word/document.xml']));
    expect(readDocx(out).watermark).toMatchObject({ text: '内部资料', diagonal: true });
  });

  it('给没有页眉的文档加水印：新建默认页眉并在节属性里引用', async () => {
    const bytes = pkg('<w:p><w:r><w:t>正文</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    doc.watermark = { text: '样品' };
    const out = await patchDocx(doc, origin!);
    const files = unzipSync(out);
    expect(strFromU8(files['word/document.xml'])).toMatch(/<w:headerReference w:type="default"/);
    expect((readDocx(out).watermark as { text?: string } | undefined)?.text).toBe('样品');
  });

  it('改文字、去掉水印', async () => {
    const bytes = withHeader(watermarkParagraphXml('机密'));
    const { doc, origin } = await openDocx(bytes);
    doc.watermark = { ...doc.watermark!, text: '绝密' };
    expect((readDocx(await patchDocx(doc, origin!)).watermark as { text?: string } | undefined)?.text).toBe('绝密');
    doc.watermark = undefined;
    const out = await patchDocx(doc, origin!);
    expect(headerXml(out)).not.toContain('<v:textpath');
    expect(readDocx(out).watermark).toBeUndefined();
  });

  it('水印和页眉文字一起改：两者都保存', async () => {
    const bytes = withHeader(`${watermarkParagraphXml('机密')}<w:p><w:r><w:t>公司</w:t></w:r></w:p>`);
    const { doc, origin } = await openDocx(bytes);
    const last = doc.header.length - 1;
    insertRaw(doc, { story: 'header', path: [last], offset: 0 }, '新·');
    const out = await patchDocx(doc, origin!);
    expect(headerXml(out)).toContain('string="机密"');
    expect(headerXml(out)).toContain('新·公司');
    expect((readDocx(out).watermark as { text?: string } | undefined)?.text).toBe('机密');
  });

  it('新建文档整篇写出：水印写进默认页眉，读回一致', () => {
    const d = newDocument();
    d.watermark = { text: '草稿', colorHex: 'FF0000', opacity: 0.25, diagonal: false };
    const bytes = writeDocx(d);
    const hdr = headerXml(bytes);
    expect(hdr).toContain('string="草稿"');
    expect(hdr).toContain('xmlns:v="urn:schemas-microsoft-com:vml"');
    expect(readDocx(bytes).watermark).toMatchObject({ text: '草稿', colorHex: 'FF0000', opacity: 0.25, diagonal: false });
    expect(readDocx(writeDocx(newDocument())).watermark).toBeUndefined();
  });
});

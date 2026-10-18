// @vitest-environment jsdom
/** 页眉页脚：插图片 / 链接（关系写进页眉自己的 .rels）、原来没有页眉时新建部件、新建文档整篇写出 */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { openDocx, patchDocx } from './patch';
import { readDocx, writeDocx } from './docx';
import { pkg, W, R } from './testPackage';
import { getPara, insertImage, insertRaw } from '../core/ops';
import { setLink } from '../core/links';
import { registerImage } from '../core/images';
import { newDocument, type HFStory } from '../core/model';

// 1×1 PNG
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HDR_TYPE = '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
const header = (inner: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr ${W} ${R}>${inner}</w:hdr>`;
const SECT_WITH_HEADER =
  '<w:headerReference w:type="default" r:id="rId20"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';
const withHeader = (inner: string) =>
  pkg('<w:p><w:r><w:t>正文第一段</w:t></w:r></w:p>', {
    sect: SECT_WITH_HEADER,
    rels: `<Relationship Id="rId20" Type="${OD}/header" Target="header1.xml"/>`,
    types: HDR_TYPE,
    extra: { 'word/header1.xml': header(inner) },
  });
const at = (story: HFStory, offset = 0) => ({ story, path: [0], offset });

describe('页眉页脚：图片和链接', () => {
  it('已有页眉里插图片：图片关系写进 header1.xml.rels，正文和文档关系不变', async () => {
    const bytes = withHeader('<w:p><w:r><w:t>公司名称</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    const id = registerImage(PNG, 'image/png');
    insertImage(doc, at('header', 0), { id, w: 60, h: 30 });
    const out = await patchDocx(doc, origin!);
    const a = unzipSync(bytes);
    const b = unzipSync(out);
    const hdr = strFromU8(b['word/header1.xml']);
    const rid = hdr.match(/r:embed="([^"]+)"/)?.[1];
    expect(rid).toBeTruthy();
    expect(hdr).toContain('公司名称');
    const rels = strFromU8(b['word/_rels/header1.xml.rels']);
    expect(rels).toMatch(new RegExp(`Id="${rid}"[^>]*Target="media/wz-image[^"]+\\.png"`));
    const media = rels.match(/Target="(media\/[^"]+)"/)![1];
    expect(Array.from(b[`word/${media}`])).toEqual(Array.from(PNG));
    expect(strFromU8(b['[Content_Types].xml'])).toMatch(/Extension="png"/);
    // 正文、文档关系逐字节不变
    expect(strFromU8(b['word/document.xml'])).toBe(strFromU8(a['word/document.xml']));
    expect(strFromU8(b['word/_rels/document.xml.rels'])).toBe(strFromU8(a['word/_rels/document.xml.rels']));
    // 读回：页眉里有这张图
    const back = readDocx(out);
    expect(back.header[0].runs.some((r) => !!r.image)).toBe(true);
  });

  it('页眉里加外部链接：链接关系写进页眉自己的关系文件', async () => {
    const bytes = withHeader('<w:p><w:r><w:t>网址 example</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    setLink(doc, at('header', 3), at('header', 10), { url: 'https://example.com' });
    const out = await patchDocx(doc, origin!);
    const b = unzipSync(out);
    expect(strFromU8(b['word/header1.xml'])).toMatch(/<w:hyperlink r:id="rIdWz\d+">/);
    expect(strFromU8(b['word/_rels/header1.xml.rels'])).toContain('Target="https://example.com" TargetMode="External"');
    expect(readDocx(out).header[0].runs.find((r) => r.props.link)?.props.link?.url).toBe('https://example.com');
  });

  it('页眉没改：什么部件都不变（字节一致）', async () => {
    const bytes = withHeader('<w:p><w:r><w:t>公司名称</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
  });

  it('新建文档整篇写出：页眉里的图片写进页眉自己的关系文件，正文图片不受影响', () => {
    const d = newDocument();
    const id = registerImage(PNG, 'image/png');
    insertRaw(d, { story: 'body', path: [0], offset: 0 }, '正文');
    insertImage(d, { story: 'body', path: [0], offset: 2 }, { id, w: 10, h: 10 });
    insertImage(d, at('header', 0), { id, w: 60, h: 30 });
    insertRaw(d, at('header', 1), '公司');
    const files = unzipSync(writeDocx(d));
    const hdr = strFromU8(files['word/header1.xml']);
    expect(hdr).toMatch(/r:embed="rIdImage1"/);
    expect(strFromU8(files['word/_rels/header1.xml.rels'])).toContain('Id="rIdImage1"');
    expect(strFromU8(files['word/document.xml'])).toMatch(/r:embed="rIdImage1"/);
    // 同一张图片只存一份
    expect(Object.keys(files).filter((k) => k.startsWith('word/media/'))).toHaveLength(1);
    const back = readDocx(writeDocx(d));
    expect(back.header[0].runs.some((r) => !!r.image)).toBe(true);
    expect(getPara(back, { story: 'body', path: [0] }).runs.some((r) => !!r.image)).toBe(true);
  });
});

describe('页眉页脚：原文件里没有的页眉 / 页脚', () => {
  it('没有页眉的文档里加页眉文字和图片：新建 header 部件、关系、内容类型，节属性里加引用，正文不变', async () => {
    const bytes = pkg('<w:p><w:r><w:t>正文第一段</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    expect(doc.header[0].hfNone).toBe(true);
    const id = registerImage(PNG, 'image/png');
    insertRaw(doc, at('header', 0), '公司名称');
    insertImage(doc, at('header', 4), { id, w: 60, h: 30 });
    insertRaw(doc, at('footer', 0), '内部资料');
    const out = await patchDocx(doc, origin!);
    const a = unzipSync(bytes);
    const b = unzipSync(out);
    const doc2 = strFromU8(b['word/document.xml']);
    const refs = [...doc2.matchAll(/<w:(header|footer)Reference w:type="default" r:id="([^"]+)"\/>/g)].map((m) => [m[1], m[2]]);
    expect(refs.map((r) => r[0]).sort()).toEqual(['footer', 'header']);
    const rels = strFromU8(b['word/_rels/document.xml.rels']);
    const hdrRel = rels.match(new RegExp(`Id="${refs.find((r) => r[0] === 'header')![1]}"[^>]*Target="([^"]+)"`))!;
    expect(strFromU8(b[`word/${hdrRel[1]}`])).toContain('公司名称');
    expect(strFromU8(b[`word/_rels/${hdrRel[1]}.rels`])).toContain('Target="media/');
    const ct = strFromU8(b['[Content_Types].xml']);
    expect(ct).toContain(`PartName="/word/${hdrRel[1]}"`);
    expect(ct).toContain('wordprocessingml.header+xml');
    expect(ct).toContain('wordprocessingml.footer+xml');
    // sectPr 里的引用在 pgSz 之前
    expect(doc2.indexOf('<w:headerReference')).toBeLessThan(doc2.indexOf('<w:pgSz'));
    expect(doc2).toContain('<w:p><w:r><w:t>正文第一段</w:t></w:r></w:p>');
    expect(strFromU8(a['word/styles.xml'])).toBe(strFromU8(b['word/styles.xml']));
    // 读回
    const back = readDocx(out);
    expect(back.header[0].runs.map((r) => r.text).join('')).toContain('公司名称');
    expect(back.header[0].hfNone).toBeFalsy();
    expect(back.footer[0].runs.map((r) => r.text).join('')).toBe('内部资料');
  });

  it('页眉里没写内容（只是占位）时不新建部件，保存结果就是原文件', async () => {
    const bytes = pkg('<w:p><w:r><w:t>正文</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
  });
});

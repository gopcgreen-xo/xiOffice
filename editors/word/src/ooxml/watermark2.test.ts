// @vitest-environment jsdom
/** 水印：首页 / 偶数页页眉也加、图片水印（读、加、去、新建文档） */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { openDocx, patchDocx } from './patch';
import { readDocx, writeDocx } from './docx';
import { pkg, W, R } from './testPackage';
import { registerImage } from '../core/images';
import { isPictureWatermark, newDocument, type Watermark } from '../core/model';
import { pictureWatermarkParagraphXml, watermarkParagraphXml, WATERMARK_NS } from '../../vendor/docx-engine/src/watermark';

const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HDR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const header = (inner: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr ${W} ${R}${WATERMARK_NS}>${inner}</w:hdr>`;
const para = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
const PG = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';
const files = (b: Uint8Array) => unzipSync(b);
const text = (b: Uint8Array, path: string) => strFromU8(files(b)[path]);
const kind = (w: Watermark | undefined) => (!w ? 'none' : isPictureWatermark(w) ? 'image' : 'text');

describe('水印：首页 / 偶数页页眉', () => {
  it('首页不同、奇偶页不同：已有的三个页眉部件都加上水印', async () => {
    const bytes = pkg(para('正文'), {
      sect: `<w:headerReference w:type="default" r:id="rId20"/><w:headerReference w:type="first" r:id="rId21"/><w:headerReference w:type="even" r:id="rId22"/><w:titlePg/>${PG}`,
      rels: ['20', '21', '22'].map((n, i) => `<Relationship Id="rId${n}" Type="${OD}/header" Target="header${i + 1}.xml"/>`).join(''),
      types: [1, 2, 3].map((n) => `<Override PartName="/word/header${n}.xml" ContentType="${HDR_CT}"/>`).join(''),
      extra: { 'word/header1.xml': header(para('默认')), 'word/header2.xml': header(para('首页')), 'word/header3.xml': header(para('偶数')) },
    });
    const { doc, origin } = await openDocx(bytes);
    // 打开 settings 里没有 evenAndOddHeaders：偶数页页眉不会显示，不加
    doc.watermark = { text: '机密' };
    const out = await patchDocx(doc, origin!);
    expect(text(out, 'word/header1.xml')).toContain('string="机密"');
    expect(text(out, 'word/header2.xml')).toContain('string="机密"');
    expect(text(out, 'word/header3.xml')).not.toContain('<v:textpath');
    // 去掉：都去掉，原有文字还在
    const { doc: d2, origin: o2 } = await openDocx(out);
    d2.watermark = undefined;
    const back = await patchDocx(d2, o2!);
    for (const n of [1, 2]) expect(text(back, `word/header${n}.xml`)).not.toContain('<v:textpath');
    expect(text(back, 'word/header2.xml')).toContain('首页');
  });

  it('首页不同但没有首页页眉：新建首页页眉（带水印）并引用', async () => {
    const bytes = pkg(para('正文'), {
      sect: `<w:headerReference w:type="default" r:id="rId20"/><w:titlePg/>${PG}`,
      rels: `<Relationship Id="rId20" Type="${OD}/header" Target="header1.xml"/>`,
      types: `<Override PartName="/word/header1.xml" ContentType="${HDR_CT}"/>`,
      extra: { 'word/header1.xml': header(para('默认')) },
    });
    const { doc, origin } = await openDocx(bytes);
    doc.watermark = { text: '草稿' };
    const out = await patchDocx(doc, origin!);
    const xml = text(out, 'word/document.xml');
    expect(xml).toMatch(/<w:headerReference w:type="first"/);
    const newPart = Object.keys(files(out)).find((k) => k.includes('wz-header'))!;
    expect(text(out, newPart)).toContain('string="草稿"');
    expect(text(out, 'word/header1.xml')).toContain('string="草稿"');
  });
});

describe('水印：图片', () => {
  const picWm = (): Watermark => ({ image: { id: registerImage(PNG, 'image/png'), widthPx: 200, heightPx: 100 }, scale: 100, washout: true });

  it('加到没有页眉的文档：新建页眉部件、图片放进它自己的关系文件；读回是图片水印', async () => {
    const bytes = pkg(para('正文'));
    const { doc, origin } = await openDocx(bytes);
    doc.watermark = picWm();
    const out = await patchDocx(doc, origin!);
    const f = files(out);
    const part = Object.keys(f).find((k) => /wz-header\d+\.xml$/.test(k))!;
    expect(strFromU8(f[part])).toContain('<v:imagedata');
    const rels = strFromU8(f[part.replace('word/', 'word/_rels/') + '.rels']);
    expect(rels).toMatch(/Target="media\/wz-image[^"]+\.png"/);
    const back = readDocx(out);
    expect(kind(back.watermark)).toBe('image');
    const w = back.watermark as Extract<Watermark, { image: unknown }>;
    expect(Array.from(unzipSync(out)[`word/${rels.match(/Target="(media\/[^"]+)"/)![1]}`])).toEqual(Array.from(PNG));
    expect(w.washout).toBe(true);
    // 按 pt 的大小写出：200×100 像素 = 150×75 pt
    expect(strFromU8(f[part])).toContain('width:150pt;height:75pt');
  });

  it('读 Word 的图片水印；没改字节不变；换成文字水印后图片水印去掉', async () => {
    const wm = pictureWatermarkParagraphXml({ image: { base64: '', mime: 'image/png', widthPx: 200, heightPx: 100 }, scale: 100 }, 'rId5', null);
    const bytes = pkg(para('正文'), {
      sect: `<w:headerReference w:type="default" r:id="rId20"/>${PG}`,
      rels: `<Relationship Id="rId20" Type="${OD}/header" Target="header1.xml"/>`,
      types: `<Override PartName="/word/header1.xml" ContentType="${HDR_CT}"/><Default Extension="png" ContentType="image/png"/>`,
      extra: {
        'word/header1.xml': header(wm),
        'word/_rels/header1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${OD}/image" Target="media/logo.png"/></Relationships>`,
      },
    });
    // 图片字节放进包里
    const zip = unzipSync(bytes);
    zip['word/media/logo.png'] = PNG;
    const { zipSync } = await import('fflate');
    const withMedia = zipSync(zip);
    const { doc, origin } = await openDocx(withMedia);
    expect(kind(doc.watermark)).toBe('image');
    expect(await patchDocx(doc, origin!)).toEqual(withMedia);
    doc.watermark = { text: '机密' };
    const out = await patchDocx(doc, origin!);
    expect(text(out, 'word/header1.xml')).not.toContain('<v:imagedata');
    expect(text(out, 'word/header1.xml')).toContain('string="机密"');
  });

  it('新建文档整篇写出：图片水印在页眉自己的关系里，读回一致', () => {
    const d = newDocument();
    d.watermark = picWm();
    const bytes = writeDocx(d);
    const f = files(bytes);
    expect(strFromU8(f['word/header1.xml'])).toContain('<v:imagedata');
    expect(strFromU8(f['word/_rels/header1.xml.rels'])).toContain('image');
    expect(kind(readDocx(bytes).watermark)).toBe('image');
    // 文字水印写到所有页眉部件里（默认 / 首页 / 偶数页）
    const t = newDocument();
    t.watermark = { text: '草稿' };
    const tb = unzipSync(writeDocx(t));
    for (const n of [1, 2, 3]) expect(strFromU8(tb[`word/header${n}.xml`])).toContain('string="草稿"');
  });

  it('文字水印文本读回（对比：文字水印不被当成图片）', async () => {
    const { doc } = await openDocx(pkg(para('x'), {
      sect: `<w:headerReference w:type="default" r:id="rId20"/>${PG}`,
      rels: `<Relationship Id="rId20" Type="${OD}/header" Target="header1.xml"/>`,
      types: `<Override PartName="/word/header1.xml" ContentType="${HDR_CT}"/>`,
      extra: { 'word/header1.xml': header(watermarkParagraphXml('机密')) },
    }));
    expect(kind(doc.watermark)).toBe('text');
  });
});

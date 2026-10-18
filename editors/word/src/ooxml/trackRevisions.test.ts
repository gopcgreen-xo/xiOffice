// @vitest-environment jsdom
/** 修订开关（审阅状态）存在 settings.xml 的 w:trackRevisions：读入、就地修补时只改这一处、整篇写出时也带上 */
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { openDocx, patchDocx, setTrackRevisionsXml } from './patch';
import { readDocx, writeDocx } from './docx';
import { insertRaw, tracking } from '../core/ops';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

function pkg(settings: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(`${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/_rels/document.xml.rels': strToU8(`${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`),
    'word/document.xml': strToU8(`${HEAD}<w:document ${W}><w:body><w:p><w:r><w:t>供方：XX有限公司</w:t></w:r></w:p><w:p><w:r><w:t>第二段</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`),
    'word/settings.xml': strToU8(`${HEAD}<w:settings ${W}>${settings}</w:settings>`),
  });
}

describe('修订开关 w:trackRevisions', () => {
  it('按 CT_Settings 的顺序插入（documentProtection / defaultTabStop 之前，扩展元素之前）；关闭时去掉', () => {
    expect(setTrackRevisionsXml(`<w:settings ${W}><w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/></w:settings>`, true)).toBe(`<w:settings ${W}><w:zoom w:percent="100"/><w:trackRevisions/><w:defaultTabStop w:val="420"/></w:settings>`);
    expect(setTrackRevisionsXml(`<w:settings ${W}><w:zoom/><w:documentProtection w:edit="readOnly"/></w:settings>`, true)).toContain('<w:trackRevisions/><w:documentProtection');
    expect(setTrackRevisionsXml(`<w:settings ${W}><w:zoom/><w15:chartTrackingRefBased/></w:settings>`, true)).toContain('<w:zoom/><w:trackRevisions/><w15:');
    expect(setTrackRevisionsXml(`<w:settings ${W}><w:zoom/></w:settings>`, true)).toBe(`<w:settings ${W}><w:zoom/><w:trackRevisions/></w:settings>`);
    expect(setTrackRevisionsXml(`<w:settings ${W}><w:zoom/><w:trackRevisions w:val="1"/><w:defaultTabStop/></w:settings>`, false)).toBe(`<w:settings ${W}><w:zoom/><w:defaultTabStop/></w:settings>`);
    // 已经有的不重复
    expect(setTrackRevisionsXml(`<w:settings ${W}><w:trackRevisions/><w:defaultTabStop/></w:settings>`, true).match(/trackRevisions/g)).toHaveLength(1);
  });

  it('读入：有 w:trackRevisions（值不是 0 / false）才算修订状态', () => {
    expect(readDocx(pkg('<w:trackRevisions/>')).trackRevisions).toBe(true);
    expect(readDocx(pkg('<w:trackRevisions w:val="false"/>')).trackRevisions).toBeUndefined();
    expect(readDocx(pkg('<w:zoom w:percent="100"/>')).trackRevisions).toBeUndefined();
  });

  it('就地修补：只打开修订开关时 document.xml 不变，settings.xml 只多了这一个元素；关掉又回到原样', async () => {
    const src = pkg('<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/>');
    const { doc, origin } = await openDocx(src);
    expect(await patchDocx(doc, origin!)).toBe(src);
    doc.trackRevisions = true;
    const on = unzipSync(await patchDocx(doc, origin!));
    const before = unzipSync(src);
    expect(strFromU8(on['word/document.xml'])).toBe(strFromU8(before['word/document.xml']));
    expect(strFromU8(on['word/settings.xml'])).toBe(strFromU8(before['word/settings.xml']).replace('<w:defaultTabStop', '<w:trackRevisions/><w:defaultTabStop'));
    // 保存后再打开仍是修订状态
    const again = await openDocx(zipSync(on));
    expect(again.doc.trackRevisions).toBe(true);
    again.doc.trackRevisions = false;
    expect(strFromU8(unzipSync(await patchDocx(again.doc, again.origin!))['word/settings.xml'])).toBe(strFromU8(before['word/settings.xml']));
  });

  it('修订中修改正文 + 保持修订开关：修订标记和开关都写进文件', async () => {
    const { doc, origin } = await openDocx(pkg('<w:zoom w:percent="100"/>'));
    tracking.on = true;
    tracking.author = '审核';
    try {
      insertRaw(doc, { story: 'body', path: [0], offset: 3 }, '{{公司2}}');
    } finally {
      tracking.on = false;
    }
    doc.trackRevisions = true;
    const files = unzipSync(await patchDocx(doc, origin!));
    expect(strFromU8(files['word/document.xml'])).toMatch(/<w:ins w:id="\d+" w:author="审核"[^>]*>.*\{\{公司2\}\}/);
    expect(strFromU8(files['word/settings.xml'])).toContain('<w:trackRevisions/>');
    // 整篇写出（新建文件）也带上
    expect(strFromU8(unzipSync(writeDocx(doc))['word/settings.xml'])).toContain('<w:trackRevisions/>');
  });
});

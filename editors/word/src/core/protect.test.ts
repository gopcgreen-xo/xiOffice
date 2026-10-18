// @vitest-environment jsdom
/** 文档保护：模式检查、密码、允许编辑的区域、读写 settings.xml / permStart */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { addPermRange, checkPassword, editClass, editError, inPerm, makeProtection, needsPassword, permRanges, removePermRange } from './protect';
import { insertRaw } from './ops';
import { newDocument, type Doc } from './model';
import { openDocx, patchDocx } from '../ooxml/patch';
import { readDocx, writeDocx } from '../ooxml/docx';
import { pkg } from '../ooxml/testPackage';

const at = (offset: number, i = 0) => ({ story: 'body' as const, path: [i], offset });
const doc1 = (): Doc => {
  const d = newDocument();
  insertRaw(d, at(0), '甲方：供应商名称，金额：100');
  return d;
};

describe('编辑检查', () => {
  it('没有保护 / 没启用：什么都放行', () => {
    const d = doc1();
    expect(editError(d, 'type', at(0), at(0))).toBeNull();
    d.protection = { edit: 'readOnly', enforced: false };
    expect(editError(d, 'type', at(0), at(0))).toBeNull();
  });

  it('只读：内容、批注、文档设置都不行；只在允许编辑的区域里可以改内容', () => {
    const d = doc1();
    d.protection = { edit: 'readOnly', enforced: true };
    for (const k of ['type', 'paste', 'delete', 'table', 'comment', 'review', 'style', 'watermark']) expect(editError(d, k, at(0), at(2)), k).toMatch(/受保护（只读）/);
    addPermRange(d, at(3), at(8));
    expect(permRanges(d)).toHaveLength(1);
    expect(inPerm(d, at(4), at(6))).toBe(true);
    expect(inPerm(d, at(4), at(12))).toBe(false); // 越出区域
    expect(inPerm(d, at(0), at(1))).toBe(false);
    expect(editError(d, 'type', at(4), at(4))).toBeNull();
    expect(editError(d, 'comment', at(4), at(6))).toBeNull();
    expect(editError(d, 'type', at(1), at(1))).not.toBeNull();
    expect(editError(d, 'style', at(4), at(4))).not.toBeNull(); // 文档级设置仍不行
  });

  it('只批注：任何位置能加批注，改内容只在允许的区域；只修订：都放行但不能接受 / 拒绝', () => {
    const d = doc1();
    d.protection = { edit: 'comments', enforced: true };
    expect(editError(d, 'comment', at(0), at(2))).toBeNull();
    expect(editError(d, 'type', at(0), at(0))).not.toBeNull();
    d.protection = { edit: 'trackedChanges', enforced: true };
    expect(editError(d, 'type', at(0), at(0))).toBeNull();
    expect(editError(d, 'style', at(0), at(0))).toBeNull();
    expect(editError(d, 'review', at(0), at(0))).toMatch(/修订/);
  });

  it('操作分类', () => {
    expect(editClass('comment')).toBe('comment');
    expect(editClass('review')).toBe('review');
    expect(editClass('watermark')).toBe('document');
    expect(editClass('type')).toBe('content');
  });

  it('允许编辑的区域：加、去；空选区不行；标记不占宽度', () => {
    const d = doc1();
    expect(() => addPermRange(d, at(2), at(2))).toThrow(/选中/);
    addPermRange(d, at(3), at(8));
    addPermRange(d, at(10), at(12));
    const r = permRanges(d);
    expect(r).toHaveLength(2);
    removePermRange(d, r[0].id);
    expect(permRanges(d)).toHaveLength(1);
  });
});

describe('密码', () => {
  it('不设密码：不用密码；设了：对的过、错的不过', async () => {
    const plain = await makeProtection('readOnly');
    expect(needsPassword(plain)).toBe(false);
    expect(await checkPassword(plain, '')).toBe(true);
    const p = await makeProtection('comments', '123456', 1000);
    expect(p).toMatchObject({ edit: 'comments', enforced: true, algorithmSid: 14 });
    expect(needsPassword(p)).toBe(true);
    expect(await checkPassword(p, '123456')).toBe(true);
    expect(await checkPassword(p, '654321')).toBe(false);
    expect(await checkPassword(p, '')).toBe(false);
  }, 30000);
});

describe('读写', () => {
  it('保存：加保护写进 settings.xml，读回一致；停止保护后去掉；没改不变', async () => {
    const bytes = pkg('<w:p><w:r><w:t>正文</w:t></w:r></w:p>', { extra: { 'word/settings.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>' } });
    const { doc, origin } = await openDocx(bytes);
    expect(doc.protection).toBeUndefined();
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
    doc.protection = await makeProtection('readOnly', 'abc', 1000);
    const out = await patchDocx(doc, origin!);
    const sx = strFromU8(unzipSync(out)['word/settings.xml']);
    expect(sx).toMatch(/<w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="rsaAES"/);
    const { doc: back, origin: o2 } = await openDocx(out);
    expect(back.protection).toMatchObject({ edit: 'readOnly', enforced: true });
    expect(await checkPassword(back.protection!, 'abc')).toBe(true);
    expect(await patchDocx(back, o2!)).toEqual(out);
    back.protection = undefined;
    expect(strFromU8(unzipSync(await patchDocx(back, o2!))['word/settings.xml'])).not.toContain('documentProtection');
  }, 30000);

  it('例外区域读写：permStart / permEnd 原样保存，读回还能识别', async () => {
    const bytes = pkg('<w:p><w:r><w:t>甲方：</w:t></w:r><w:permStart w:id="5" w:edGrp="everyone"/><w:r><w:t>供应商</w:t></w:r><w:permEnd w:id="5"/><w:r><w:t>（盖章）</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    expect(permRanges(doc)).toHaveLength(1);
    expect(inPerm(doc, at(4), at(6))).toBe(true);
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
    addPermRange(doc, at(0 + 8), at(0 + 9));
    const out = await patchDocx(doc, origin!);
    const xml = strFromU8(unzipSync(out)['word/document.xml']);
    expect(xml).toContain('<w:permStart w:id="5" w:edGrp="everyone"/>');
    expect(xml).toMatch(/<w:permStart w:id="20000" w:edGrp="everyone"\/>/);
    expect(permRanges(readDocx(out))).toHaveLength(2);
  });

  it('新建文档整篇写出：保护写进 settings', () => {
    const d = newDocument();
    d.protection = { edit: 'trackedChanges', enforced: true };
    const sx = strFromU8(unzipSync(writeDocx(d))['word/settings.xml']);
    expect(sx).toContain('<w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>');
    expect(sx.indexOf('documentProtection')).toBeLessThan(sx.indexOf('defaultTabStop'));
  });
});

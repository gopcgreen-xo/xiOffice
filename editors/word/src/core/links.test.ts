// @vitest-environment jsdom
/** 超链接与书签：模型操作、保存（新外部链接登记关系、书签原样写回）、读回 */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { addBookmark, checkBookmarkName, clearLink, linkAt, listBookmarks, normalizeUrl, removeBookmark, setLink } from './links';
import { getPara, insertRaw } from './ops';
import { newDocument, paraText, type Doc } from './model';
import { openDocx, patchDocx } from '../ooxml/patch';
import { readDocx, writeDocx } from '../ooxml/docx';
import { pkg } from '../ooxml/testPackage';

const at = (offset: number, i = 0) => ({ story: 'body' as const, path: [i], offset });
const doc3 = (): Doc => {
  const d = newDocument();
  insertRaw(d, at(0), '请参阅合同附件');
  return d;
};
const text = (d: Doc, i = 0) => paraText(getPara(d, at(0, i))).replaceAll('￼', '');

describe('超链接', () => {
  it('补全地址：没写协议补 https，邮箱补 mailto', () => {
    expect(normalizeUrl(' example.com/a ')).toBe('https://example.com/a');
    expect(normalizeUrl('http://x.cn')).toBe('http://x.cn');
    expect(normalizeUrl('a@b.com')).toBe('mailto:a@b.com');
    expect(normalizeUrl('')).toBe('');
  });

  it('加链接后文字变蓝加下划线；去掉链接后还原；光标在链接里能找到完整范围', () => {
    const d = doc3();
    setLink(d, at(1), at(3), { url: 'https://example.com' });
    const p = getPara(d, at(0));
    expect(p.runs.map((r) => [r.text, !!r.props.link, r.props.color, r.props.underline])).toEqual([
      ['请', false, 'auto', false],
      ['参阅', true, '0563C1', true],
      ['合同附件', false, 'auto', false],
    ]);
    const hit = linkAt(d, at(2))!;
    expect(hit.text).toBe('参阅');
    expect([hit.from.offset, hit.to.offset]).toEqual([1, 3]);
    expect(linkAt(d, at(5))).toBeNull();
    clearLink(d, hit.from, hit.to);
    expect(getPara(d, at(0)).runs).toHaveLength(1);
    expect(getPara(d, at(0)).runs[0].props.link).toBeUndefined();
  });

  it('保存：新外部链接登记为外部关系，原有链接保留原 rid；其它段落字节不变', async () => {
    const other = '<w:p><w:r><w:t>第二段  保持原样</w:t></w:r></w:p>';
    const bytes = pkg(
      `<w:p><w:r><w:t>请参阅合同附件</w:t></w:r></w:p>${other}<w:p><w:hyperlink r:id="rId9"><w:r><w:t>旧链接</w:t></w:r></w:hyperlink></w:p>`,
      { rels: '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://old.example" TargetMode="External"/>' },
    );
    const { doc, origin } = await openDocx(bytes);
    // 读入时 rid 对应的地址记在 url 里
    expect(getPara(doc, at(0, 2)).runs[0].props.link).toEqual({ rid: 'rId9', url: 'https://old.example' });
    setLink(doc, at(1, 0), at(3, 0), { url: 'https://example.com/a?x=1&y=2' });
    const out = await patchDocx(doc, origin!);
    const files = unzipSync(out);
    const xml = strFromU8(files['word/document.xml']);
    const rels = strFromU8(files['word/_rels/document.xml.rels']);
    const rid = xml.match(/<w:hyperlink r:id="(rIdWz\d+)"/)?.[1];
    expect(rid).toBeTruthy();
    expect(rels).toContain(`Id="${rid}"`);
    expect(rels).toContain('Target="https://example.com/a?x=1&amp;y=2" TargetMode="External"');
    expect(rels).toContain('Id="rId9"');
    expect(xml).toContain(other);
    expect(xml).toContain('<w:hyperlink r:id="rId9">');
    // 读回
    const back = readDocx(out);
    expect(getPara(back, at(0, 0)).runs.find((r) => r.props.link)?.props.link?.url).toBe('https://example.com/a?x=1&y=2');
  });

  it('新建文档整篇写出：链接写进关系，读回地址不变；文档内链接只写 anchor', () => {
    const d = doc3();
    setLink(d, at(0), at(2), { url: 'https://example.com' });
    setLink(d, at(4), at(7), { anchor: 'chapter1' });
    const files = unzipSync(writeDocx(d));
    const xml = strFromU8(files['word/document.xml']);
    expect(xml).toMatch(/<w:hyperlink r:id="rIdLink1">/);
    expect(xml).toContain('<w:hyperlink w:anchor="chapter1">');
    expect(strFromU8(files['word/_rels/document.xml.rels'])).toContain('Target="https://example.com" TargetMode="External"');
    const back = readDocx(writeDocx(d));
    const links = getPara(back, at(0)).runs.filter((r) => r.props.link).map((r) => r.props.link!);
    expect(links.map((l) => l.url ?? l.anchor)).toEqual(['https://example.com', 'chapter1']);
  });
});

describe('书签', () => {
  it('名字规则', () => {
    expect(checkBookmarkName('第一章')).toBeNull();
    expect(checkBookmarkName('a_1')).toBeNull();
    expect(checkBookmarkName('')).not.toBeNull();
    expect(checkBookmarkName('1a')).not.toBeNull();
    expect(checkBookmarkName('a b')).not.toBeNull();
    expect(checkBookmarkName('a-b')).not.toBeNull();
    expect(checkBookmarkName('a'.repeat(41))).not.toBeNull();
  });

  it('添加、列出、同名移动、删除；标记不占宽度也不改文字', () => {
    const d = doc3();
    addBookmark(d, at(1), at(3), '条款');
    expect(text(d)).toBe('请参阅合同附件');
    let list = listBookmarks(d);
    expect(list.map((b) => [b.name, b.text])).toEqual([['条款', '参阅']]);
    addBookmark(d, at(3), at(5), '条款'); // 同名：移动
    list = listBookmarks(d);
    expect(list).toHaveLength(1);
    expect(list[0].text).toBe('合同');
    addBookmark(d, at(0), at(0), '开头'); // 空选区也行
    expect(listBookmarks(d).map((b) => b.name)).toEqual(['开头', '条款']);
    expect(removeBookmark(d, '开头')).toBe(true);
    expect(removeBookmark(d, '不存在')).toBe(false);
    expect(listBookmarks(d).map((b) => b.name)).toEqual(['条款']);
    expect(() => addBookmark(d, at(0), at(1), '1x')).toThrow();
  });

  it('保存：书签原样写回（id 不与已有的冲突），其它段落不变；读回能列出', async () => {
    const other = '<w:p><w:r><w:t>第二段</w:t></w:r></w:p>';
    const bytes = pkg(`<w:p><w:bookmarkStart w:id="0" w:name="旧书签"/><w:r><w:t>请参阅合同</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>${other}`);
    const { doc, origin } = await openDocx(bytes);
    expect(listBookmarks(doc).map((b) => b.name)).toEqual(['旧书签']);
    addBookmark(doc, at(2, 0), at(4, 0), '新书签'); // 开头已有一个书签标记，占 1 个偏移
    const out = await patchDocx(doc, origin!);
    const xml = strFromU8(unzipSync(out)['word/document.xml']);
    expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="旧书签"/>');
    expect(xml).toMatch(/<w:bookmarkStart w:id="10000" w:name="新书签"\/>/);
    expect(xml).toContain('<w:bookmarkEnd w:id="10000"/>');
    expect(xml).toContain(other);
    const back = readDocx(out);
    expect(listBookmarks(back).map((b) => [b.name, b.text])).toEqual([
      ['旧书签', '请参阅合同'],
      ['新书签', '参阅'],
    ]);
  });
});

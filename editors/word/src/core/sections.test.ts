// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readDocx, writeDocx } from '../ooxml/docx';
import { layoutDocument, type Measurer, type PageLayout } from './layout';
import { newDocument, newParagraph, pageGeom, paraText, sectionList, type Doc, type Paragraph } from './model';
import { backspace, getPara, hfLinked, insertRaw, insertSectionBreak, setHFLinked, setSectionPage } from './ops';
import { isCJK } from './text';

const fake: Measurer = {
  width: (ch) => (isCJK(ch) ? 1 : 0.5),
  metrics: () => ({ ascent: 0.88, descent: 0.12 }),
};
const at = (i: number, offset = 0) => ({ story: 'body' as const, path: [i], offset });
const text = (lines: PageLayout['header']) => lines.map((l) => l.line.glyphs.map((g) => g.ch).join('')).join('');
const heads = (d: Doc) => layoutDocument(d, fake).pages.map((p) => text(p.header));
const feet = (d: Doc) => layoutDocument(d, fake).pages.map((p) => text(p.footer));

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const hdr = (tag: 'hdr' | 'ftr', t: string) => strToU8(`<w:${tag} ${W}><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:${tag}>`);
const ref = (kind: 'header' | 'footer', type: string, rid: string) => `<w:${kind}Reference w:type="${type}" r:id="${rid}"/>`;

/** 三节：封面（无页眉）、目录（第二节只改页眉，页脚沿用）、正文（首页不同） */
function threeSections(): Uint8Array {
  const p = (t: string, sect = '') => `<w:p>${sect ? `<w:pPr>${sect}</w:pPr>` : ''}<w:r><w:t>${t}</w:t></w:r></w:p>`;
  const doc =
    `<w:document ${W}><w:body>` +
    p('封面', `<w:sectPr>${ref('header', 'default', 'h1')}${ref('footer', 'default', 'f1')}</w:sectPr>`) +
    p('目录', `<w:sectPr>${ref('header', 'default', 'h2')}</w:sectPr>`) +
    p('正文') +
    `<w:sectPr>${ref('header', 'default', 'h3')}${ref('header', 'first', 'h4')}<w:titlePg/></w:sectPr>` +
    '</w:body></w:document>';
  const rels =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    ['h1:header1', 'h2:header2', 'h3:header3', 'h4:header4', 'f1:footer1'].map((x) => {
      const [id, f] = x.split(':');
      return `<Relationship Id="${id}" Type="x" Target="${f}.xml"/>`;
    }).join('') +
    '</Relationships>';
  return zipSync({
    'word/document.xml': strToU8(doc),
    'word/_rels/document.xml.rels': strToU8(rels),
    'word/header1.xml': hdr('hdr', ''),
    'word/header2.xml': hdr('hdr', '目录页眉'),
    'word/header3.xml': hdr('hdr', '正文页眉'),
    'word/header4.xml': hdr('hdr', '正文首页'),
    'word/footer1.xml': hdr('ftr', '页脚'),
  });
}

describe('各节的页眉页脚', () => {
  it('读取时各节使用自己的页眉页脚，不再提示', () => {
    const d = readDocx(threeSections());
    expect(d.lossy).toEqual([]);
    d.body.push(newParagraph('\f正文第二页'));
    expect(heads(d)).toEqual(['', '目录页眉', '正文首页', '正文页眉']);
    // 没有页脚引用的节与上一节相同，共用同一个数组；正文首页用首页页脚（各节都没有定义，为空，与 Word 相同）
    expect(feet(d)).toEqual(['页脚', '页脚', '', '页脚']);
    const secs = sectionList(d);
    expect(secs[1].hf.footer).toBe(secs[0].hf.footer);
    expect(secs.map((s) => s.hf.titlePg)).toEqual([false, false, true]);
    expect(layoutDocument(d, fake).pages[1].header[0].path).toEqual([1, 0]);
  });

  it('保存后再打开保持各节内容与“与上一节相同”的关系', () => {
    const bytes = writeDocx(readDocx(threeSections()));
    const files = unzipSync(bytes);
    // 共用的页脚只写一份；原文件里没有的分区（第一节的首页 / 偶数页页眉页脚）不写成空文件：
    // 页脚只有第一节的默认 1 个，页眉为封面 1 个 + 目录 1 个 + 正文 2 个（默认、首页）
    const parts = Object.keys(files).filter((f) => /word\/(header|footer)\d/.test(f));
    expect(parts.filter((f) => f.includes('footer'))).toHaveLength(1);
    expect(parts.filter((f) => f.includes('header'))).toHaveLength(4);
    const xml = strFromU8(files['word/document.xml']);
    const sects = xml.match(/<w:sectPr>.*?<\/w:sectPr>/g)!;
    expect(sects).toHaveLength(3);
    expect(sects[1].match(/Reference/g)).toHaveLength(1);
    const back = readDocx(bytes);
    expect(back.lossy).toEqual([]);
    expect(heads(back)).toEqual(['', '目录页眉', '正文首页']);
    const secs = sectionList(back);
    expect(secs[1].hf.footer).toBe(secs[0].hf.footer);
    expect(secs[2].hf.footer).toBe(secs[1].hf.footer);
  });

  it('插入分节符后默认与上一节相同，取消链接后可分别编辑', () => {
    const d = newDocument();
    d.body = [newParagraph('甲乙')];
    d.header[0].runs = [{ text: '页眉', props: d.defaults }];
    insertSectionBreak(d, at(0, 1), 'nextPage');
    expect(heads(d)).toEqual(['页眉', '页眉']);
    const first = [0, 0];
    expect(hfLinked(d, 'header', first)).toBeNull();
    expect(hfLinked(d, 'header', [0])).toBe(true);

    // 编辑第一节的页眉：链接的第二节一起改变
    insertRaw(d, { story: 'header', path: first, offset: 2 }, '一');
    expect(heads(d)).toEqual(['页眉一', '页眉一']);

    // 第二节取消链接后单独修改
    const pos = setHFLinked(d, 'header', [0], false);
    expect(pos.path).toEqual([0]);
    insertRaw(d, { ...pos, offset: 0 }, '二');
    expect(heads(d)).toEqual(['页眉一', '二页眉一']);
    expect(hfLinked(d, 'header', [0])).toBe(false);

    // 撤销快照保留共用关系；重新链接后恢复为上一节的内容
    const snap = structuredClone(d);
    expect(sectionList(snap)[1].hf.footer).toBe(sectionList(snap)[0].hf.footer);
    setHFLinked(d, 'header', [0], true);
    expect(heads(d)).toEqual(['页眉一', '页眉一']);
  });

  it('删除分节符时，该节消失，后一节的页眉页脚保留', () => {
    const d = newDocument();
    d.body = [newParagraph('甲乙丙')];
    insertSectionBreak(d, at(0, 1), 'nextPage');
    insertSectionBreak(d, at(1, 1), 'nextPage');
    setHFLinked(d, 'header', [1, 0], false);
    (d.body[1] as Paragraph).sectionHF!.header[0].runs = [{ text: '第二节', props: d.defaults }];
    // 第三节原来与第二节相同，随第二节一起取消与第一节的链接
    expect(heads(d)).toEqual(['', '第二节', '第二节']);
    expect(hfLinked(d, 'header', [0])).toBe(true);
    // 在第二段段首退格：第一节的分节符被删除，第二节的分节符（及页眉）移到第一段
    backspace(d, at(1));
    expect(d.body).toHaveLength(2);
    expect(getPara(d, at(0)).sectionHF?.header.map(paraText)).toEqual(['第二节']);
    expect(heads(d)).toEqual(['第二节', '第二节']);
  });

  it('首页不同按每节的首页判断，连续分节不产生新的首页', () => {
    const d = newDocument();
    d.setup.titlePg = true;
    d.header[0].runs = [{ text: '页眉', props: d.defaults }];
    d.headerFirst[0].runs = [{ text: '首', props: d.defaults }];
    d.body = [newParagraph('一\f二'), newParagraph('三\f四')];
    insertSectionBreak(d, at(1, 0), 'nextPage');
    expect(heads(d)).toEqual(['首', '页眉', '首', '页眉']);
    const c = newDocument();
    c.setup.titlePg = true;
    c.headerFirst[0].runs = [{ text: '首', props: c.defaults }];
    c.body = [newParagraph('一\f二'), newParagraph('三\f四')];
    insertSectionBreak(c, at(1, 0), 'continuous');
    expect(heads(c)).toEqual(['首', '', '']);
  });
});

describe('分节符的类型', () => {
  // OOXML：sectPr 的 w:type 说的是“这一节如何开始”，所以段后分节符的类型要看下一节的 sectPr
  const doc =
    `<w:document ${W}><w:body>` +
    '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>表头</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>条款</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>签字</w:t></w:r></w:p>' +
    '<w:sectPr><w:type w:val="oddPage"/></w:sectPr>' +
    '</w:body></w:document>';
  const types = (d: Doc) => d.body.map((b) => (b as Paragraph).props.sectionBreak);

  it('读取：类型取自下一节的节属性', () => {
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc) }));
    expect(types(d)).toEqual(['continuous', 'oddPage', null]);
    // 表头后面是连续分节符：条款接在同一页
    expect(layoutDocument(d, fake).pages[0].body.map((l) => l.path[0])).toEqual([0, 1]);
  });

  it('写出：类型写到下一节的节属性上，再读回不变', () => {
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc) }));
    const bytes = writeDocx(d);
    const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
    const sects = xml.match(/<w:sectPr>.*?<[/]w:sectPr>/g)!.map((x) => /<w:type w:val="(\w+)"/.exec(x)?.[1] ?? '');
    expect(sects).toEqual(['', 'continuous', 'oddPage']);
    expect(types(readDocx(bytes))).toEqual(['continuous', 'oddPage', null]);
  });
});

describe('各节的纸张与页边距', () => {
  /** 两节：第一节 A4 纵向，第二节 A4 横向、左边距 50 */
  function twoPages(): Uint8Array {
    const pg = (w: number, h: number, left: number) => `<w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="1440" w:bottom="1440" w:left="${left}" w:right="1800" w:header="851" w:footer="992"/>`;
    const doc =
      `<w:document ${W}><w:body>` +
      `<w:p><w:pPr><w:sectPr>${pg(11906, 16838, 1800)}</w:sectPr></w:pPr><w:r><w:t>纵向</w:t></w:r></w:p>` +
      '<w:p><w:r><w:t>横向</w:t></w:r></w:p>' +
      `<w:sectPr>${pg(16838, 11906, 1000)}</w:sectPr>` +
      '</w:body></w:document>';
    return zipSync({ 'word/document.xml': strToU8(doc) });
  }

  it('读取时各节使用自己的纸张与页边距，不再提示', () => {
    const d = readDocx(twoPages());
    expect(d.lossy).toEqual([]);
    const pages = layoutDocument(d, fake).pages;
    expect(pages.map((p) => [Math.round(p.setup.width), Math.round(p.setup.height)])).toEqual([[595, 842], [842, 595]]);
    expect(pages[0].body[0].x).toBe(90);
    expect(pages[1].body[0].x).toBe(50);
  });

  it('保存后再打开各节的页面设置不变', () => {
    const bytes = writeDocx(readDocx(twoPages()));
    const sects = strFromU8(unzipSync(bytes)['word/document.xml']).match(/<w:pgSz[^>]*>/g)!;
    expect(sects).toEqual(['<w:pgSz w:w="11906" w:h="16838"/>', '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>']);
    const back = readDocx(bytes);
    expect(back.lossy).toEqual([]);
    expect(sectionList(back).map((s) => s.setup.marginLeft)).toEqual([90, 50]);
  });

  it('页面设置可只改本节或整篇文档；连续分节换纸张时从新页开始', () => {
    const d = newDocument();
    d.body = [newParagraph('甲乙')];
    insertSectionBreak(d, at(0, 1), 'continuous');
    expect(layoutDocument(d, fake).pages).toHaveLength(1);
    const landscape = { ...pageGeom(d.setup), width: d.setup.height, height: d.setup.width };
    // 只改最后一节：第一节保持原来的纵向
    setSectionPage(d, null, landscape, false);
    const pages = layoutDocument(d, fake).pages;
    expect(pages).toHaveLength(2);
    expect(pages.map((p) => p.setup.width > p.setup.height)).toEqual([false, true]);
    // 改第一节的左边距
    setSectionPage(d, 0, { ...pageGeom(sectionList(d)[0].setup), marginLeft: 40 }, false);
    expect(layoutDocument(d, fake).pages[0].body[0].x).toBe(40);
    // 整篇文档：各节统一，不再单独保存
    setSectionPage(d, 0, landscape, true);
    expect((d.body[0] as Paragraph).sectionPage).toBeUndefined();
    expect(layoutDocument(d, fake).pages).toHaveLength(1);
  });
});

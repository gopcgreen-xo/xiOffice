// @vitest-environment jsdom
/** 一般的域（fldChar / fldSimple）：读成开始 / 结束标记 + 结果文字，保存时域代码原样或按代码生成；手写 XML，不依赖真实文档 */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { openDocx, patchDocx } from './patch';
import { pkg } from './testPackage';
import { fieldSpans } from '../core/fieldView';
import { setFieldInstr, updateFields } from '../core/fields';
import { deleteRange } from '../core/ops';
import type { Paragraph } from '../core/model';

const fld = (instr: string, res: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${res}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const docXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['word/document.xml']);
const wellFormed = (xml: string) => expect(new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagName('parsererror')).toHaveLength(0);
const text = (p: Paragraph) => p.runs.filter((r) => !r.fc).map((r) => r.text).join('');
/** document.xml 里每个段落的 fldChar 类型序列 */
const chars = (xml: string) => [...xml.matchAll(/w:fldCharType="(\w+)"/g)].map((m) => m[1]).join(',');

describe('一般的域：读写', () => {
  const body =
    `<w:p><w:r><w:t>表</w:t></w:r>${fld(' SEQ 表 \\* ARABIC ', '9')}<w:r><w:t>：报价</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>表</w:t></w:r>${fld(' SEQ 表 \\* ARABIC ', '9')}</w:p>` +
    '<w:p><w:fldSimple w:instr=" QUOTE 新文字 "><w:r><w:t>旧文字</w:t></w:r></w:fldSimple></w:p>';

  it('读成成对的标记，结果是普通文字；没改时保存 = 原文件字节', async () => {
    const bytes = pkg(body);
    const { doc, origin } = await openDocx(bytes);
    const p0 = doc.body[0] as Paragraph;
    expect(text(p0)).toBe('表9：报价');
    const [f] = fieldSpans(p0);
    expect(f.instr).toBe('SEQ 表 \\* ARABIC');
    expect(p0.runs[f.begin].fc?.raw).toContain('<w:instrText');
    expect(fieldSpans(doc.body[2] as Paragraph)[0].instr).toBe('QUOTE 新文字');
    expect(doc.lossy).not.toContain('域（保留显示文字）');
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
  });

  it('更新域：域代码部分原样写回，只换结果；再打开结果和代码都在', async () => {
    const { doc, origin } = await openDocx(pkg(body));
    expect(updateFields(doc, {}).updated).toBe(3);
    const xml = docXml(await patchDocx(doc, origin!));
    wellFormed(xml);
    expect(xml.match(/<w:instrText xml:space="preserve"> SEQ 表 \\\* ARABIC <\/w:instrText>/g)).toHaveLength(2);
    expect(chars(xml)).toBe('begin,separate,end,begin,separate,end,begin,separate,end');
    const again = await openDocx(await patchDocx(doc, origin!));
    expect(again.doc.body.map((b) => text(b as Paragraph))).toEqual(['表1：报价', '表2', '新文字']);
    expect(fieldSpans(again.doc.body[2] as Paragraph)[0].instr).toBe('QUOTE 新文字');
  });

  it('改域代码：按新代码生成，嵌套的 {…} 写成嵌套域；锁定的写 w:fldLock', async () => {
    const { doc, origin } = await openDocx(pkg(`<w:p>${fld(' QUOTE 1 ', '1')}</w:p>`));
    const p = doc.body[0] as Paragraph;
    const id = fieldSpans(p)[0].id;
    setFieldInstr(p, id, '= 1 + {QUOTE 2}');
    p.runs.find((r) => r.fc?.kind === 'begin')!.fc!.locked = true;
    expect(updateFields(doc, {}).locked).toBe(1);
    p.runs.find((r) => r.fc?.kind === 'begin')!.fc!.locked = false;
    updateFields(doc, {});
    const xml = docXml(await patchDocx(doc, origin!));
    wellFormed(xml);
    // 嵌套域带自己的结果（Word 算外层公式时用它）
    expect(chars(xml)).toBe('begin,begin,separate,end,separate,end');
    expect(xml).toContain('QUOTE 2</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">2</w:t>');
    const again = await openDocx(await patchDocx(doc, origin!));
    const f = fieldSpans(again.doc.body[0] as Paragraph)[0];
    expect(f.instr).toBe('= 1 + {QUOTE 2}');
    expect(text(again.doc.body[0] as Paragraph)).toBe('3');
  });

  it('删掉了一半的域：剩下的标记不写，结果当普通文字', async () => {
    const { doc, origin } = await openDocx(pkg(`<w:p><w:r><w:t>ab</w:t></w:r>${fld(' QUOTE 1 ', 'xy')}<w:r><w:t>cd</w:t></w:r></w:p>`));
    // 开始标记在 2，结果 3–4，结束标记在 5：删掉 4–6（结果的后一个字 y 和结束标记）
    deleteRange(doc, { story: 'body', path: [0], offset: 4 }, { story: 'body', path: [0], offset: 6 });
    const xml = docXml(await patchDocx(doc, origin!));
    wellFormed(xml);
    expect(chars(xml)).toBe('');
    expect(text((await openDocx(await patchDocx(doc, origin!))).doc.body[0] as Paragraph)).toBe('abxcd');
  });

  it('跨段的域（目录等）和结果里又有域的：照旧只保留显示文字', async () => {
    const toc =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>第一章</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>第二章</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> QUOTE a </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${fld(' QUOTE b ', 'b')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const { doc } = await openDocx(pkg(toc));
    expect(doc.body.flatMap((b) => fieldSpans(b as Paragraph))).toEqual([]);
    expect(doc.body.map((b) => text(b as Paragraph))).toEqual(['第一章', '第二章', 'b']);
    expect(doc.lossy).toContain('域（保留显示文字）');
  });
});

// 给 Microsoft Word 验收的文件（WORD_CHECK_DIR=输出目录 时才生成，再用 scripts/word-check.ps1 打开确认没有修复提示）
it.runIf(process.env.WORD_CHECK_DIR)('生成域的 Word 验收文件', async () => {
  const { mkdirSync, writeFileSync } = await import('fs');
  const { join } = await import('path');
  const out = process.env.WORD_CHECK_DIR!;
  mkdirSync(out, { recursive: true });
  const src = pkg(
    `<w:p><w:r><w:t>表</w:t></w:r>${fld(' SEQ 表 \\* ARABIC ', '9')}<w:r><w:t>：报价</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>表</w:t></w:r>${fld(' SEQ 表 \\* ARABIC ', '9')}</w:p>` +
      '<w:p><w:fldSimple w:instr=" QUOTE 新文字 "><w:r><w:t>旧文字</w:t></w:r></w:fldSimple></w:p>' +
      `<w:p><w:r><w:t>日期：</w:t></w:r>${fld(' DATE \\@ "yyyy/M/d" ', '2000/1/1')}</w:p>`,
  );
  const updated = await openDocx(src);
  updateFields(updated.doc, {});
  writeFileSync(join(out, 'fields-updated.docx'), await patchDocx(updated.doc, updated.origin!));

  const edited = await openDocx(src);
  const p = edited.doc.body[3] as Paragraph;
  setFieldInstr(p, fieldSpans(p)[0].id, 'DATE \\@ "yyyy年M月d日" \\* MERGEFORMAT');
  const q = edited.doc.body[2] as Paragraph;
  setFieldInstr(q, fieldSpans(q)[0].id, '= 1 + {QUOTE 2} \\# "0.00"');
  updateFields(edited.doc, {});
  writeFileSync(join(out, 'fields-edited.docx'), await patchDocx(edited.doc, edited.origin!));

  const half = await openDocx(src);
  deleteRange(half.doc, { story: 'body', path: [0], offset: 3 }, { story: 'body', path: [0], offset: 4 });
  writeFileSync(join(out, 'fields-half.docx'), await patchDocx(half.doc, half.origin!));
});

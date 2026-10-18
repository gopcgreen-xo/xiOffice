// @vitest-environment jsdom
/** 模板占位符与批量生成：替换（跨文本段、表格、页眉、保留格式）、数据解析、文件名、保存往返 */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { fillDoc, outputName, parseCsv, placeholders, tableFromJson, tableFromRows } from './merge';
import { applyRunProps, getPara, insertRaw, insertTable } from './ops';
import { newDocument, paraText } from './model';
import { openDocx, patchDocx } from '../ooxml/patch';
import { readDocx } from '../ooxml/docx';
import { pkg, W, R } from '../ooxml/testPackage';

const at = (offset: number, i = 0) => ({ story: 'body' as const, path: [i], offset });
const txt = (p: ReturnType<typeof getPara>) => paraText(p).replaceAll('￼', '');

describe('占位符', () => {
  it('列出文档里的占位符：正文、表格、页眉；按出现顺序，重复的计数', () => {
    const d = newDocument();
    insertRaw(d, at(0), '甲方：{{供应商}}，合同号 {{合同号}}，再次 {{ 供应商 }}');
    insertRaw(d, { story: 'header', path: [0], offset: 0 }, '编号 {{合同号}}');
    expect(placeholders(d)).toEqual([
      { key: '供应商', count: 2 },
      { key: '合同号', count: 2 },
    ]);
  });

  it('替换：值替换占位符，格式沿用占位符开头的字；占位符被拆在多个文本段里也能换', () => {
    const d = newDocument();
    insertRaw(d, at(0), '甲方：{{供应商}}（盖章）');
    // 把占位符拆成三段：{{供 / 应 / 商}}，中间一段加粗
    applyRunProps(d, at(5), at(6), { bold: true });
    applyRunProps(d, at(3), at(5), { color: 'FF0000' });
    expect(getPara(d, at(0)).runs.length).toBeGreaterThan(3);
    const r = fillDoc(d, { 供应商: '示例科技公司' });
    expect(r).toMatchObject({ missing: [], replaced: 1 });
    const p = getPara(d, at(0));
    expect(txt(p)).toBe('甲方：示例科技公司（盖章）');
    // 新文字用占位符第一个字（“{”，红色）的格式；占位符以外的文字格式不变
    const hit = p.runs.find((x) => x.text.includes('示例'))!;
    expect(hit.props.color).toBe('FF0000');
    expect(p.runs.at(-1)!.props.color).toBe('auto');
  });

  it('同一段里多个占位符、值里有换行；缺少的字段：默认原样保留并报告，blankMissing 时换成空', () => {
    const d = newDocument();
    insertRaw(d, at(0), '{{a}}-{{b}}-{{c}}');
    const r = fillDoc(d, { a: '1', b: '第一行\r\n第二行' });
    expect(r.missing).toEqual(['c']);
    expect(txt(getPara(d, at(0)))).toBe('1-第一行\n第二行-{{c}}');
    fillDoc(d, {}, { blankMissing: true });
    expect(txt(getPara(d, at(0)))).toBe('1-第一行\n第二行-');
  });

  it('值里带 {{ 不会被再次替换；表格单元格、页眉里的占位符都换', () => {
    const d = newDocument();
    insertRaw(d, at(0), '{{x}}');
    insertTable(d, { story: 'body', path: [0], offset: 0 }, 1, 2);
    insertRaw(d, { story: 'body', path: [0, 0, 1, 0], offset: 0 }, '单元格 {{x}}');
    insertRaw(d, { story: 'header', path: [0], offset: 0 }, '页眉 {{x}}');
    const r = fillDoc(d, { x: '{{y}}' });
    expect(r.replaced).toBe(3);
    expect(placeholders(d)).toEqual([]);
    expect(txt(getPara(d, { story: 'header', path: [0] } as never))).toBe('页眉 { {y}}');
  });

  it('图片等不占文字的对象在占位符之外时保持不动', () => {
    const d = newDocument();
    insertRaw(d, at(0), 'AB{{k}}CD');
    const before = getPara(d, at(0)).runs.length;
    fillDoc(d, { k: '' });
    expect(txt(getPara(d, at(0)))).toBe('ABCD');
    expect(before).toBeGreaterThan(0);
  });
});

describe('数据', () => {
  it('csv：引号、逗号、换行、双引号转义、BOM、制表符分隔', () => {
    expect(parseCsv('﻿名称,备注\r\n"甲, 乙","说""明""\n第二行"\n丙,')).toEqual([
      ['名称', '备注'],
      ['甲, 乙', '说"明"\n第二行'],
      ['丙', ''],
    ]);
    expect(parseCsv('a\tb\n1\t2')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });

  it('表格行 → 数据：第一行列名，空行跳过，重复列名加序号，数字去掉浮点尾巴', () => {
    const t = tableFromRows([['名称', '', '名称'], ['甲', 1.1 + 2.2, 'x'], [null, null, null], ['乙', 5, 'y']]);
    expect(t.columns).toEqual(['名称', '列2', '名称2']);
    expect(t.records).toEqual([
      { 名称: '甲', 列2: '3.3', 名称2: 'x' },
      { 名称: '乙', 列2: '5', 名称2: 'y' },
    ]);
  });

  it('json：对象数组；不是对象数组时报错', () => {
    expect(tableFromJson('[{"a":1,"b":"x"},{"a":2,"c":true}]')).toEqual({
      columns: ['a', 'b', 'c'],
      records: [
        { a: '1', b: 'x' },
        { a: '2', c: 'true' },
      ],
    });
    expect(() => tableFromJson('[1,2]')).toThrow();
  });

  it('输出文件名：字段、序号、非法字符、批内重名', () => {
    const taken = new Set<string>();
    expect(outputName('合同_{{供应商}}', { 供应商: '甲/乙:丙' }, 0, taken)).toBe('合同_甲_乙_丙');
    expect(outputName('合同_{{供应商}}', { 供应商: '甲/乙:丙' }, 1, taken)).toBe('合同_甲_乙_丙 (2)');
    expect(outputName('{{空}}', { 空: '' }, 4, taken)).toBe('文档5');
    expect(outputName('第{#}份', {}, 2, taken)).toBe('第3份');
  });
});

describe('保存往返（原文件就地修补）', () => {
  it('Word 把占位符拆成几个 run：填值后只有含占位符的段落变，其它段落、样式逐字节不变', async () => {
    const other = '<w:p><w:r><w:t>第二段  保持原样</w:t></w:r></w:p>';
    const tmpl = '<w:p><w:r><w:t xml:space="preserve">合同号：{</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>{编</w:t></w:r><w:r><w:t>号}}</w:t></w:r><w:r><w:t>（甲方）</w:t></w:r></w:p>';
    const hdr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr ${W} ${R}><w:p><w:r><w:t>{{供应商}} 合同</w:t></w:r></w:p></w:hdr>`;
    const bytes = pkg(`${tmpl}${other}`, {
      sect: '<w:headerReference w:type="default" r:id="rId20"/><w:pgSz w:w="11906" w:h="16838"/>',
      rels: '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      types: '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
      extra: { 'word/header1.xml': hdr },
    });
    const { doc, origin } = await openDocx(bytes);
    expect(placeholders(doc).map((p) => p.key)).toEqual(['编号', '供应商']);
    const copy = structuredClone(doc);
    const r = fillDoc(copy, { 编号: 'HT-001', 供应商: '示例科技公司' });
    expect(r.replaced).toBe(2);
    const out = await patchDocx(copy, origin!);
    const a = unzipSync(bytes);
    const b = unzipSync(out);
    const xml = strFromU8(b['word/document.xml']);
    expect(xml).toContain(other);
    expect(xml).toContain('HT-001');
    expect(xml).not.toContain('{{');
    expect(strFromU8(b['word/styles.xml'])).toBe(strFromU8(a['word/styles.xml']));
    expect(strFromU8(b['word/header1.xml']).replace(/<[^>]+>/g, '')).toContain('示例科技公司 合同');
    // 读回
    expect(txt(getPara(readDocx(out), at(0)))).toBe('合同号：HT-001（甲方）');
    // 原模板没被改动（复制的文档上填值）
    expect(placeholders(doc).length).toBe(2);
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
  });
});

describe('Word 邮件合并域（MERGEFIELD）', () => {
  const run = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
  // 复杂域：begin / instrText / separate / 结果 / end，各在独立的 w:r 里（Word 的写法）
  const complex = (instr: string, result: string) =>
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const body =
    `<w:p>${run('甲方：')}${complex(' MERGEFIELD  供应商  \\* MERGEFORMAT ', '«供应商»')}${run('（盖章）')}</w:p>` +
    `<w:p>${run('地址：')}<w:fldSimple w:instr=" MERGEFIELD &quot;送货 地址&quot; \\b &quot;【&quot; \\f &quot;】&quot; "><w:r><w:t>«送货 地址»</w:t></w:r></w:fldSimple></w:p>` +
    `<w:p>${run('第二段  保持原样')}</w:p>`;

  it('指令解析：字段名（带引号 / 不带）、前后缀开关；不是 MERGEFIELD 返回 null', async () => {
    const { parseMergeInstr } = await import('./merge');
    expect(parseMergeInstr(' MERGEFIELD  供应商  \\* MERGEFORMAT ')).toEqual({ name: '供应商' });
    expect(parseMergeInstr(' MERGEFIELD "送货 地址" \\b "【" \\f "】"')).toEqual({ name: '送货 地址', before: '【', after: '】' });
    expect(parseMergeInstr(' PAGE ')).toBeNull();
  });

  it('读：显示为 «字段名»，不算丢内容；列入占位符；打开即保存字节不变', async () => {
    const bytes = pkg(body);
    const { doc, origin } = await openDocx(bytes);
    expect(doc.lossy).toEqual([]);
    const p = getPara(doc, at(0));
    expect(p.runs.map((r) => r.opaque?.merge?.name ?? r.text)).toEqual(['甲方：', '供应商', '（盖章）']);
    expect(p.runs[1].opaque?.text).toBe('«供应商»');
    expect(placeholders(doc)).toEqual([
      { key: '供应商', count: 1 },
      { key: '送货 地址', count: 1 },
    ]);
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
  });

  it('填值：复杂域和简单域都换成值（前后缀只在有值时加），其它段落字节不变，读回是普通文字', async () => {
    const bytes = pkg(body);
    const { doc, origin } = await openDocx(bytes);
    const copy = structuredClone(doc);
    const r = fillDoc(copy, { 供应商: '示例科技公司', '送货 地址': '示例市东城区' });
    expect(r).toMatchObject({ missing: [], replaced: 2 });
    const out = await patchDocx(copy, origin!);
    const xml = strFromU8(unzipSync(out)['word/document.xml']);
    expect(xml).not.toContain('MERGEFIELD');
    expect(xml).toContain('<w:p><w:r><w:t xml:space="preserve">第二段  保持原样</w:t></w:r></w:p>');
    const back = readDocx(out);
    expect(txt(getPara(back, at(0)))).toBe('甲方：示例科技公司（盖章）');
    expect(txt(getPara(back, at(0, 1)))).toBe('地址：【示例市东城区】');
    // 值为空：前后缀也不加；字段缺失：原样保留
    const blank = structuredClone(doc);
    fillDoc(blank, { 供应商: '', '送货 地址': '' });
    expect(txt(getPara(blank, at(0, 1)))).toBe('地址：');
    const miss = structuredClone(doc);
    expect(fillDoc(miss, {}).missing).toEqual(['供应商', '送货 地址']);
    expect(getPara(miss, at(0)).runs.some((x) => x.opaque?.merge)).toBe(true);
  });

  it('新插入的邮件合并域：保存成 Word 的 MERGEFIELD，读回仍是域，也能填值', async () => {
    const { insertMergeField } = await import('./merge');
    const bytes = pkg('<w:p><w:r><w:t>合同号：</w:t></w:r></w:p>');
    const { doc, origin } = await openDocx(bytes);
    insertMergeField(doc, at(4), '合同号');
    const out = await patchDocx(doc, origin!);
    expect(strFromU8(unzipSync(out)['word/document.xml'])).toMatch(/<w:fldSimple w:instr=" MERGEFIELD 合同号 \\\* MERGEFORMAT ">/);
    const back = readDocx(out);
    expect(placeholders(back)).toEqual([{ key: '合同号', count: 1 }]);
    fillDoc(back, { 合同号: 'HT-9' });
    expect(txt(getPara(back, at(0)))).toBe('合同号：HT-9');
    expect(() => insertMergeField(doc, at(0), '')).toThrow();
  });
});

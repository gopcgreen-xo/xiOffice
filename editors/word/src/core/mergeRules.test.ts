// @vitest-environment jsdom
/** 邮件合并域的格式开关、规则域（IF / NEXT / NEXTIF / SKIPIF）、一份文档放多条数据、合并成一个文档、脚注和页眉表格里的占位符 */
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { applyCase, formatDate, formatNumber, formatValue } from './mergeFormat';
import { compare, evalRule, parseMergeInstr, parseRule } from './mergeInstr';
import { combineRecords, fillDoc, fillPartXml, fillRecords, hasNextRule, insertNextRecord, placeholders } from './merge';
import { getPara, insertRaw } from './ops';
import { newDocument, paraText, type Doc } from './model';
import { openDocx, patchDocx } from '../ooxml/patch';
import { readDocx } from '../ooxml/docx';
import { pkg, W, R } from '../ooxml/testPackage';

const at = (offset: number, i = 0) => ({ story: 'body' as const, path: [i], offset });
const txt = (d: Doc, i = 0) => paraText(getPara(d, at(0, i))).replaceAll('￼', '');
const run = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const simple = (instr: string) => `<w:fldSimple w:instr="${instr}"><w:r><w:t>«x»</w:t></w:r></w:fldSimple>`;
/** 复杂域，可以嵌套：parts 里的字符串是指令文字，数组是嵌套的域 */
const complex = (parts: (string | string[])[], result = '结果'): string => {
  const instr = (ps: (string | string[])[]): string =>
    ps
      .map((p) =>
        typeof p === 'string'
          ? `<w:r><w:instrText xml:space="preserve">${p}</w:instrText></w:r>`
          : `<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(p)}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«n»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`,
      )
      .join('');
  return `<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr(parts)}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
};

describe('格式开关：数字 \\#', () => {
  it('千分位、小数位、前后文字、负数、零', () => {
    expect(formatNumber('1234567.891', '#,##0.00')).toBe('1,234,567.89');
    expect(formatNumber('5', '0.00')).toBe('5.00');
    expect(formatNumber('5.5', '0.##')).toBe('5.5');
    expect(formatNumber('5', '0.##')).toBe('5');
    expect(formatNumber('12', '000')).toBe('012');
    expect(formatNumber('1234.5', '¥#,##0.00')).toBe('¥1,234.50');
    expect(formatNumber('88', "0'元'")).toBe('88元');
    expect(formatNumber('-3.5', '0.0')).toBe('-3.5');
    expect(formatNumber('-3.5', '0.0;(0.0)')).toBe('(3.5)');
    expect(formatNumber('0', '0.0;(0.0);零')).toBe('零');
    expect(formatNumber('1,234.5', '0')).toBe('1235');
  });
  it('不是数字：原样', () => {
    expect(formatNumber('待定', '0.00')).toBe('待定');
  });
});

describe('格式开关：日期 \\@', () => {
  it('常见写法与图片', () => {
    expect(formatDate('2026-10-04', 'yyyy年M月d日')).toBe('2026年10月4日');
    expect(formatDate('2026/3/7', 'yyyy-MM-dd')).toBe('2026-03-07');
    expect(formatDate('2026年3月7日', 'yy.MM.dd')).toBe('26.03.07');
    expect(formatDate('2026-10-04', 'dddd')).toBe('星期日');
    expect(formatDate('2026-10-04', 'MMMM')).toBe('十月');
    expect(formatDate('2026-10-04 14:05:09', 'HH:mm:ss')).toBe('14:05:09');
    expect(formatDate('2026-10-04 14:05', 'h:mm AM/PM')).toBe('2:05 PM');
    expect(formatDate('2026-10-04', "yyyy'年度'")).toBe('2026年度');
  });
  it('Excel 序列号；认不出的原样', () => {
    expect(formatDate('46299', 'yyyy-MM-dd')).toBe('2026-10-04');
    expect(formatDate('明天', 'yyyy')).toBe('明天');
    expect(formatDate('2026-02-31', 'yyyy')).toBe('2026-02-31');
  });
});

describe('格式开关：大小写 \\* 和综合', () => {
  it('Upper / Lower / FirstCap / Caps', () => {
    expect(applyCase('hello wORLD', 'Upper')).toBe('HELLO WORLD');
    expect(applyCase('Hello World', 'Lower')).toBe('hello world');
    expect(applyCase('hello wORLD', 'FirstCap')).toBe('Hello world');
    expect(applyCase('hello wORLD x-ray', 'Caps')).toBe('Hello World X-Ray');
  });
  it('指令解析带出各开关；formatValue 先数字 / 日期再大小写', () => {
    expect(parseMergeInstr(' MERGEFIELD 金额 \\# "#,##0.00" \\* MERGEFORMAT ')).toEqual({ name: '金额', numFmt: '#,##0.00' });
    expect(parseMergeInstr(' MERGEFIELD 日期 \\@ "yyyy年M月d日" ')).toEqual({ name: '日期', dateFmt: 'yyyy年M月d日' });
    expect(parseMergeInstr(' MERGEFIELD name \\* Upper ')).toEqual({ name: 'name', caseFmt: 'Upper' });
    expect(parseMergeInstr(' MERGEFIELD n \\# 0.0 \\b "¥" ')).toEqual({ name: 'n', numFmt: '0.0', before: '¥' });
    expect(formatValue('abc', { caseFmt: 'Upper' })).toBe('ABC');
    expect(formatValue('', { numFmt: '0.00' })).toBe('');
  });
  it('填值：MERGEFIELD 的格式开关生效，前缀后缀只在有值时加', async () => {
    const body = `<w:p>${run('金额：')}${simple('MERGEFIELD 金额 \\# &quot;#,##0.00&quot; \\b &quot;¥&quot;')}${run('，日期：')}${simple('MERGEFIELD 日期 \\@ &quot;yyyy年M月d日&quot;')}${run('，')}${simple('MERGEFIELD en \\* Caps')}</w:p>`;
    const { doc } = await openDocx(pkg(body));
    const d = structuredClone(doc);
    fillDoc(d, { 金额: '1234.5', 日期: '2026-10-04', en: 'hello world' });
    expect(txt(d)).toBe('金额：¥1,234.50，日期：2026年10月4日，Hello World');
    const e = structuredClone(doc);
    fillDoc(e, { 金额: '', 日期: '', en: '' });
    expect(txt(e)).toBe('金额：，日期：，');
  });
});

describe('规则域：IF', () => {
  it('解析嵌套的 MERGEFIELD；比较（数值、文字、通配符）', () => {
    const r = parseRule(' IF { MERGEFIELD 金额 } > 1000 "大额" "小额" ')!;
    expect(r.kind).toBe('if');
    expect(r.toks.map((t) => t.t)).toEqual(['field', 'op', 'word', 'str', 'str']);
    expect(evalRule(r, () => '2000')).toEqual({ kind: 'text', text: '大额' });
    expect(evalRule(r, () => '500')).toEqual({ kind: 'text', text: '小额' });
    expect(compare('10', '>', '9')).toBe(true); // 数字按数值，不是文字比较
    expect(compare('Abc', '=', 'abc')).toBe(true);
    expect(compare('示例科技公司', '=', '示例*')).toBe(true);
    expect(compare('示例科技公司', '<>', '*广州*')).toBe(true);
    expect(parseRule(' IF 缺少比较符 ')).toBeNull();
    expect(parseRule(' PAGE ')).toBeNull();
  });
  it('读：复杂域 IF 嵌套 MERGEFIELD 读成一个整体，不算丢内容；填值按条件出文字；打开即保存字节不变', async () => {
    const body = `<w:p>${run('类别：')}${complex([' IF ', [' MERGEFIELD 金额 '], ' > 1000 "大额" "小额" '], '小额')}${run('。')}</w:p><w:p>${run('保持原样')}</w:p>`;
    const bytes = pkg(body);
    const { doc, origin } = await openDocx(bytes);
    expect(doc.lossy).toEqual([]);
    const p = getPara(doc, at(0));
    expect(p.runs.map((r) => (r.opaque?.rule ? '<规则>' : r.text))).toEqual(['类别：', '<规则>', '。']);
    expect(placeholders(doc)).toEqual([{ key: '金额', count: 1 }]);
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
    const a = structuredClone(doc);
    fillDoc(a, { 金额: '5000' });
    expect(txt(a)).toBe('类别：大额。');
    const b = structuredClone(doc);
    fillDoc(b, { 金额: '20' });
    const out = await patchDocx(b, origin!);
    expect(strFromU8(unzipSync(out)['word/document.xml'])).not.toContain('instrText');
    expect(txt(readDocx(out))).toBe('类别：小额。');
    expect(strFromU8(unzipSync(out)['word/document.xml'])).toContain(run('保持原样'));
  });
  it('结果里直接放字段：IF 条件成立时显示另一个字段的值', async () => {
    const body = `<w:p>${complex([' IF { MERGEFIELD 备注 } &lt;&gt; "" ', [' MERGEFIELD 备注 '], ' "无" '])}</w:p>`;
    const { doc } = await openDocx(pkg(body));
    const a = structuredClone(doc);
    fillDoc(a, { 备注: '加急' });
    expect(txt(a)).toBe('加急');
    const b = structuredClone(doc);
    fillDoc(b, { 备注: '' });
    expect(txt(b)).toBe('无');
  });
});

describe('规则域：NEXT / NEXTIF / SKIPIF 与一份文档放多条数据', () => {
  const label = (d: Doc) => {
    // 两行：名称 + 下一条记录 + 名称
    insertRaw(d, at(0), '{{名称}}');
    insertNextRecord(d, at(6));
    insertRaw(d, at(7), '/{{名称}}');
  };
  it('NEXT：同一份文档用掉两条数据；数据不够时后面的字段为空', () => {
    const d = newDocument();
    label(d);
    expect(hasNextRule(d)).toBe(true);
    const recs = [{ 名称: '甲' }, { 名称: '乙' }, { 名称: '丙' }];
    const a = structuredClone(d);
    const r = fillRecords(a, recs, 0);
    expect(txt(a)).toBe('甲/乙');
    expect(r.consumed).toBe(2);
    const b = structuredClone(d);
    const r2 = fillRecords(b, recs, 2);
    expect(txt(b)).toBe('丙/');
    expect(r2.consumed).toBe(2);
  });
  it('单条数据：NEXT 被去掉', () => {
    const d = newDocument();
    label(d);
    fillDoc(d, { 名称: '甲' });
    expect(txt(d)).toBe('甲/甲');
  });
  it('SKIPIF：条件成立的数据不出文档；NEXTIF：条件成立才用下一条', async () => {
    const body = `<w:p>${complex([' SKIPIF { MERGEFIELD 状态 } = "作废" '], '')}${run('{{名称}}')}</w:p>`;
    const { doc } = await openDocx(pkg(body));
    expect(hasNextRule(doc)).toBe(true);
    const recs = [{ 名称: '甲', 状态: '有效' }, { 名称: '乙', 状态: '作废' }];
    const a = structuredClone(doc);
    expect(fillRecords(a, recs, 0)).toMatchObject({ skipped: false });
    expect(txt(a)).toBe('甲');
    const b = structuredClone(doc);
    expect(fillRecords(b, recs, 1)).toMatchObject({ skipped: true });
  });
});

describe('合并成一个文档', () => {
  const tmpl = () => {
    const d = newDocument();
    insertRaw(d, at(0), '合同 {{名称}}');
    return d;
  };
  it('每条数据一份，第二份起另起一页；页眉按第一条填', () => {
    const d = tmpl();
    insertRaw(d, { story: 'header', path: [0], offset: 0 }, '页眉 {{名称}}');
    const r = combineRecords(d, [{ 名称: '甲' }, { 名称: '乙' }, { 名称: '丙' }]);
    expect(r.count).toBe(3);
    expect(r.doc.body.map((b) => (b.kind === 'p' ? paraText(b) : ''))).toEqual(['合同 甲', '合同 乙', '合同 丙']);
    expect(r.doc.body.map((b) => b.kind === 'p' && b.props.pageBreakBefore)).toEqual([false, true, true]);
    expect(paraText(r.doc.header[0])).toBe('页眉 甲');
  });
  it('SKIPIF 的数据不计入；NEXT 多条合进一份；多节模板报错', () => {
    const d = newDocument();
    insertRaw(d, at(0), '{{名称}}');
    insertNextRecord(d, at(6));
    insertRaw(d, at(7), '、{{名称}}');
    const r = combineRecords(d, [{ 名称: 'a' }, { 名称: 'b' }, { 名称: 'c' }]);
    expect(r.count).toBe(2);
    expect(r.doc.body.map((b) => (b.kind === 'p' ? paraText(b).replaceAll('￼', '') : ''))).toEqual(['a、b', 'c、']);
    const two = tmpl();
    two.body.push(structuredClone(two.body[0]));
    (two.body[0] as { props: { sectionBreak: string } }).props.sectionBreak = 'nextPage';
    expect(() => combineRecords(two, [{ 名称: 'x' }])).toThrow(/一节/);
  });
  it('原文件模板：合并后的文档能按原文件修补保存，读回份数和内容对', async () => {
    const bytes = pkg(`<w:p>${run('合同 {{名称}}')}</w:p><w:p>${run('第二段 {{名称}}')}</w:p>`);
    const { doc, origin } = await openDocx(bytes);
    const r = combineRecords(doc, [{ 名称: '甲' }, { 名称: '乙' }]);
    const out = await patchDocx(r.doc, origin!);
    const back = readDocx(out);
    expect(back.body.map((b) => (b.kind === 'p' ? paraText(b) : ''))).toEqual(['合同 甲', '第二段 甲', '合同 乙', '第二段 乙']);
    expect(strFromU8(unzipSync(out)['word/document.xml'])).toMatch(/<w:pageBreakBefore\/>/);
  });
});

describe('脚注 / 尾注、页眉页脚表格里的占位符', () => {
  it('部件 XML 里的占位符：跨 w:t 也换，换行变 w:br，缺少的字段原样保留', () => {
    const xml = '<w:p><w:r><w:t>名称：{{名</w:t></w:r><w:r><w:t>称}}，{{缺}}</w:t></w:r></w:p>';
    const miss = new Set<string>();
    const out = fillPartXml(xml, (k) => (k === '名称' ? '甲\n乙' : undefined), miss);
    expect(out).toContain('名称：甲');
    expect(out).toContain('<w:br/>');
    expect(out).toContain('{{缺}}');
    expect([...miss]).toEqual(['缺']);
  });
  it('保存：脚注部件、页眉里的表格（原始 XML）都按数据替换，显示用的模型也填了', async () => {
    const fn = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:footnotes ${W} ${R}><w:footnote w:id="1"><w:p><w:r><w:t>注：{{供应商}}提供</w:t></w:r></w:p></w:footnote></w:footnotes>`;
    const hdr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr ${W} ${R}><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>{{供应商}}合同</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:hdr>`;
    const bytes = pkg(`<w:p>${run('正文')}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>`, {
      sect: '<w:headerReference w:type="default" r:id="rId20"/><w:pgSz w:w="11906" w:h="16838"/>',
      rels:
        '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      types:
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
      extra: { 'word/header1.xml': hdr, 'word/footnotes.xml': fn },
    });
    const { doc, origin } = await openDocx(bytes);
    expect(placeholders(doc)).toEqual([{ key: '供应商', count: 2 }]);
    const copy = structuredClone(doc);
    fillDoc(copy, { 供应商: '示例科技公司' });
    expect(placeholders(copy)).toEqual([]);
    const out = await patchDocx(copy, origin!);
    const files = unzipSync(out);
    expect(strFromU8(files['word/footnotes.xml'])).toContain('注：示例科技公司提供');
    expect(strFromU8(files['word/header1.xml'])).toContain('示例科技公司合同');
    expect(strFromU8(files['word/header1.xml'])).not.toContain('{{');
    // 模板本身没动
    expect(await patchDocx(doc, origin!)).toEqual(bytes);
  });
});

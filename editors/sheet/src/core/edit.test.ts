import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { listRefs, moveFormula, shiftFormula, shiftSqref } from './formula/refs';
import { dateSerial, parseInput } from './input';
import { applyEntry, blockFromText, clearRange, copyBlock, insertDelete, mergeCells, pasteBlock, setColWidth, setFreeze, setInput, setRowHeight, Tx } from './ops';
import { fillSeries } from './fill';
import { addCfRule, clearCfRules, CF_LOOKS, removeCfRule } from './cfEdit';
import { blankPackage, saveCsv, saveXlsx, markSaved } from './save';
import { deriveStyle, readStyleBook, writeStyles, formatCode } from './styles';
import { blankWorkbook, readCsv, readWorkbook, type Workbook } from './workbook';

// ---------- 最小 xlsx ----------

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function pkg(sheetXml: string[], extra: Record<string, string> = {}, wbExtra = ''): Uint8Array {
  const files = blankPackage(sheetXml.map((_, i) => `S${i + 1}`));
  sheetXml.forEach((x, i) => (files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(`<worksheet xmlns="${MAIN}" xmlns:r="${REL}">${x}</worksheet>`)));
  if (wbExtra) files['xl/workbook.xml'] = strToU8(strFromU8(files['xl/workbook.xml']).replace('</sheets>', `</sheets>${wbExtra}`));
  for (const [k, v] of Object.entries(extra)) files[k] = strToU8(v);
  return zipSync(files);
}

function open(data: Uint8Array): Workbook {
  const b = readWorkbook(data);
  b.source = data;
  return b;
}

const part = (data: Uint8Array, p: string) => strFromU8(unzipSync(data)[p]);
const edit = (b: Workbook, label: string, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, label);
  fn(tx);
  return tx.commit();
};

describe('公式引用', () => {
  it('列出引用：单格、区域、整行整列、跨表、引号表名；函数名和字符串里的不算', () => {
    const refs = listRefs(`SUM(A1:B2,$C$3)+LOG10(4)+'表 一'!D5+明细!A:A+Sheet2!$2:$3&"A1"`);
    expect(refs.map((r) => [r.sheet, r.a, r.b])).toEqual([
      [undefined, { absC: false, c: 0, absR: false, r: 0 }, { absC: false, c: 1, absR: false, r: 1 }],
      [undefined, { absC: true, c: 2, absR: true, r: 2 }, undefined],
      ['表 一', { absC: false, c: 3, absR: false, r: 4 }, undefined],
      ['明细', { absC: false, c: 0 }, { absC: false, c: 0 }],
      ['Sheet2', { absR: true, r: 1 }, { absR: true, r: 2 }],
    ]);
  });

  it('复制平移：相对引用移动，绝对不动，移出表格变 #REF!', () => {
    expect(moveFormula('A1+$B$1+B$2+$C3', 2, 1)).toBe('B3+$B$1+C$2+$C5');
    expect(moveFormula('SUM(A1:A3)', -1, 0)).toBe('SUM(#REF!)');
    expect(moveFormula('Sheet2!A1*2', 1, 1)).toBe('Sheet2!B2*2');
  });

  it('插入 / 删除行：引用跟着移动，被删掉的变 #REF!，区域缩小', () => {
    const ins = { sheet: 'S1', axis: 'row' as const, at: 2, n: 2 };
    expect(shiftFormula('A1+A3+$A$5+SUM(A2:A4)', 'S1', ins)).toBe('A1+A5+$A$7+SUM(A2:A6)');
    expect(shiftFormula('A3', 'S2', ins)).toBe('A3');
    expect(shiftFormula('S1!A3', 'S2', ins)).toBe('S1!A5');
    const del = { sheet: 'S1', axis: 'row' as const, at: 2, n: -2 };
    expect(shiftFormula('A3+A5+SUM(A1:A4)+SUM(A3:A4)', 'S1', del)).toBe('#REF!+A3+SUM(A1:A2)+SUM(#REF!)');
    expect(shiftSqref('A1:B5 D3', { axis: 'col', at: 1, n: 1 })).toBe('A1:C5 E3');
  });
});

describe('输入识别', () => {
  it('数字、千分位、百分比、货币、布尔、公式、强制文字', () => {
    expect(parseInput('12.5')).toEqual({ kind: 'value', v: 12.5 });
    expect(parseInput('1,234.50')).toEqual({ kind: 'value', v: 1234.5, fmt: '#,##0.00' });
    expect(parseInput('12%')).toEqual({ kind: 'value', v: 0.12, fmt: '0%' });
    expect(parseInput('¥8')).toEqual({ kind: 'value', v: 8, fmt: '"¥"#,##0' });
    expect(parseInput('true')).toEqual({ kind: 'value', v: true });
    expect(parseInput('=A1+1')).toEqual({ kind: 'formula', f: 'A1+1' });
    expect(parseInput("'0012")).toEqual({ kind: 'value', v: '0012' });
    expect(parseInput('')).toEqual({ kind: 'empty' });
  });

  it('编号与长数字保持文字；文本格式下都是文字；已有格式时不再套', () => {
    expect(parseInput('00123')).toEqual({ kind: 'value', v: '00123' });
    expect(parseInput('420102199001011234')).toEqual({ kind: 'value', v: '420102199001011234' });
    expect(parseInput('12', '@')).toEqual({ kind: 'value', v: '12' });
    expect(parseInput('1,000', '0.00')).toEqual({ kind: 'value', v: 1000 });
  });

  it('日期和时间', () => {
    expect(dateSerial(2026, 9, 29)).toBe(46294);
    expect(parseInput('2026-9-29')).toEqual({ kind: 'value', v: 46294, fmt: 'yyyy/m/d' });
    expect(parseInput('2026年9月29日')).toEqual({ kind: 'value', v: 46294, fmt: 'yyyy"年"m"月"d"日"' });
    expect(parseInput('9/29', 'General', false, new Date(2026, 0, 1))).toEqual({ kind: 'value', v: 46294, fmt: 'm"月"d"日"' });
    expect(parseInput('12:00')).toEqual({ kind: 'value', v: 0.5, fmt: 'h:mm' });
    expect(parseInput('2026/9/29 18:00')).toEqual({ kind: 'value', v: 46294.75, fmt: 'yyyy/m/d h:mm' });
    expect(parseInput('2026-2-30')).toEqual({ kind: 'value', v: '2026-2-30' });
  });
});

describe('样式', () => {
  it('派生格式：只追加，完全相同的复用；写回保留其它段', () => {
    const xml = `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="宋体"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>`;
    const book = readStyleBook(xml);
    const a = deriveStyle(book, 0, { numFmt: '0.00%' });
    const b = deriveStyle(book, 0, { numFmt: '0.00%' });
    const c = deriveStyle(book, a, { bold: true, fill: '#ffff00' });
    expect([a, b, c]).toEqual([1, 1, 2]);
    expect(formatCode(book, a)).toBe('0.00%');
    const custom = deriveStyle(book, 0, { numFmt: '0.000"吨"' });
    expect(formatCode(book, custom)).toBe('0.000"吨"');
    const out = writeStyles(xml, book);
    expect(out).toContain('<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000&quot;吨&quot;"/></numFmts><fonts count="2">');
    expect(out).toContain('<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>');
    expect(out).toContain('<cellXfs count="4">');
    const again = readStyleBook(out);
    expect(again.xfs.length).toBe(4);
    expect(again.fonts[1]).toBe('<font><b/><sz val="11"/><name val="宋体"/></font>');
  });
});

describe('编辑与撤销', () => {
  it('写入、清除、撤销、重做', () => {
    const b = blankWorkbook();
    const e1 = edit(b, '输入', (tx) => {
      setInput(tx, 0, 0, 0, '12%');
      setInput(tx, 0, 1, 1, '你好');
    })!;
    const s = b.sheets[0];
    expect(s.cells[0][0]).toMatchObject({ v: 0.12, t: '12%' });
    expect([s.rows, s.cols]).toEqual([2, 2]);
    const e2 = edit(b, '清除', (tx) => clearRange(tx, 0, { r0: 0, c0: 0, r1: 5, c1: 5 }))!;
    expect(b.sheets[0].cells[0][0]).toBeUndefined();
    // 清除内容保留格式
    expect(b.sheets[0].blankStyles?.[0][0]).toBe(s.cells[0]?.[0]?.s ?? 1);
    applyEntry(b, e2, -1);
    expect(b.sheets[0].cells[1][1]?.v).toBe('你好');
    applyEntry(b, e1, -1);
    expect(b.sheets[0].cells[1]?.[1]).toBeUndefined();
    expect(b.sheets[0].rows).toBe(0);
    applyEntry(b, e1, 1);
    expect(b.sheets[0].cells[0][0]?.t).toBe('12%');
    expect(edit(b, '无', () => {})).toBeNull();
  });

  it('插入 / 删除行：格子、合并、行高、其它表的公式跟着移动', () => {
    const b = blankWorkbook();
    b.sheets.push({ ...structuredClone(b.sheets[0]), name: 'Sheet2' });
    edit(b, '准备', (tx) => {
      setInput(tx, 0, 0, 0, '1');
      setInput(tx, 0, 2, 0, '3');
      setInput(tx, 0, 3, 0, '=A1+A3');
      setInput(tx, 1, 0, 0, '=Sheet1!A3*2');
      mergeCells(tx, 0, { r0: 2, c0: 1, r1: 3, c1: 2 }, true);
      setRowHeight(tx, 0, 2, 2, { height: 40 });
    });
    const e = edit(b, '插入行', (tx) => insertDelete(tx, 0, 'row', 1, 2))!;
    const s = b.sheets[0];
    expect(s.cells[4][0]?.v).toBe(3);
    expect(s.cells[5][0]?.f).toBe('A1+A5');
    expect(b.sheets[1].cells[0][0]?.f).toBe('Sheet1!A5*2');
    expect(s.merges).toEqual([{ r0: 4, c0: 1, r1: 5, c1: 2 }]);
    expect(s.rowDefs?.[4]?.ht).toBe(30);
    expect(s.shifts).toEqual([{ axis: 'row', at: 1, n: 2 }]);
    applyEntry(b, e, -1);
    expect(b.sheets[0].cells[3][0]?.f).toBe('A1+A3');
    expect(b.sheets[1].cells[0][0]?.f).toBe('Sheet1!A3*2');
    edit(b, '删除行', (tx) => insertDelete(tx, 0, 'row', 2, -1));
    expect(b.sheets[0].cells[2][0]?.f).toBe('A1+#REF!');
    expect(b.sheets[1].cells[0][0]?.f).toBe('Sheet1!#REF!*2');
  });

  it('插入列沿用左边一列的格式；列宽拆分合并', () => {
    const b = blankWorkbook();
    edit(b, '列宽', (tx) => setColWidth(tx, 0, 1, 3, { width: 140 }));
    expect(b.sheets[0].colDefs).toEqual([{ min: 1, max: 3, width: 20 }]);
    edit(b, '列宽', (tx) => setColWidth(tx, 0, 2, 2, { hidden: true }));
    expect(b.sheets[0].colDefs).toEqual([
      { min: 1, max: 1, width: 20 },
      { min: 2, max: 2, width: 20, hidden: true },
      { min: 3, max: 3, width: 20 },
    ]);
    expect(b.sheets[0].colWidths).toEqual({ 1: 140, 2: 0, 3: 140 });
    edit(b, '插入列', (tx) => insertDelete(tx, 0, 'col', 0, 1));
    expect(b.sheets[0].colDefs?.[0]).toEqual({ min: 2, max: 2, width: 20 });
  });

  it('复制粘贴：公式相对平移、格式和合并一起；外部文字按输入识别', () => {
    const b = blankWorkbook();
    edit(b, '准备', (tx) => {
      setInput(tx, 0, 0, 0, '5');
      setInput(tx, 0, 0, 1, '=A1*2');
    });
    const block = copyBlock(b, 0, { r0: 0, c0: 0, r1: 0, c1: 1 });
    edit(b, '粘贴', (tx) => pasteBlock(tx, 0, 3, 2, block));
    expect(b.sheets[0].cells[3][3]?.f).toBe('C4*2');
    const ext = blockFromText('1,000\t2026-1-2\r\n"a\tb"\t\r\n');
    expect([ext.rows, ext.cols]).toEqual([2, 2]);
    edit(b, '粘贴', (tx) => pasteBlock(tx, 0, 10, 0, ext));
    expect(b.sheets[0].cells[10][0]).toMatchObject({ v: 1000, t: '1,000' });
    expect(b.sheets[0].cells[11][0]?.v).toBe('a\tb');
    edit(b, '转置', (tx) => pasteBlock(tx, 0, 20, 0, block, 'values', true));
    expect(b.sheets[0].cells[20][0]?.v).toBe(5);
    // 只粘贴值：公式格子粘贴它的计算结果（这里还没算，是空）
    expect(b.sheets[0].cells[21]?.[0]).toBeUndefined();
  });

  it('填充柄：数字序列、单个数字复制、日期按天、带数字的文字递增、公式平移', () => {
    const b = blankWorkbook();
    edit(b, '准备', (tx) => {
      setInput(tx, 0, 0, 0, '1');
      setInput(tx, 0, 1, 0, '3');
      setInput(tx, 0, 0, 1, '7');
      setInput(tx, 0, 0, 2, '2026-1-31');
      setInput(tx, 0, 0, 3, '第09项');
      setInput(tx, 0, 0, 4, '=A1*2');
    });
    edit(b, '填充', (tx) => {
      fillSeries(tx, 0, { r0: 0, c0: 0, r1: 1, c1: 0 }, { r0: 0, c0: 0, r1: 3, c1: 0 });
      fillSeries(tx, 0, { r0: 0, c0: 1, r1: 0, c1: 4 }, { r0: 0, c0: 1, r1: 2, c1: 4 });
    });
    const col = (c: number) => [0, 1, 2, 3].map((r) => b.sheets[0].cells[r]?.[c]?.t);
    expect(col(0)).toEqual(['1', '3', '5', '7']);
    expect(col(1).slice(0, 3)).toEqual(['7', '7', '7']);
    expect(col(2).slice(0, 3)).toEqual(['2026/1/31', '2026/2/1', '2026/2/2']);
    expect(col(3).slice(0, 3)).toEqual(['第09项', '第10项', '第11项']);
    expect(b.sheets[0].cells[2][4]?.f).toBe('A3*2');
  });
});

describe('保存', () => {
  const sheet1 = `<dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="14.25"/><cols><col min="1" max="1" width="20.5" customWidth="1"/></cols><sheetData><row r="1" spans="1:2" x14ac:dyDescent="0.25"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><f>B1*2</f><v>4</v></c><c r="B2" t="s"><v>1</v></c></row></sheetData><conditionalFormatting sqref="A2:B5"><cfRule type="expression" priority="1"><formula>$A2&gt;1</formula></cfRule></conditionalFormatting><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/>`;
  const sst = `<sst xmlns="${MAIN}" count="2" uniqueCount="2"><si><t>名称</t></si><si><r><rPr><b/></rPr><t>粗</t></r><r><t>体</t></r></si></sst>`;

  function sample() {
    const data = pkg([sheet1], { 'xl/sharedStrings.xml': sst, 'xl/calcChain.xml': '<calcChain/>', 'xl/charts/chart1.xml': '<chart/>' }, '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">S1!$A$1:$B$5</definedName></definedNames>');
    // 补上 sharedStrings / calcChain 的关系与内容类型
    const files = unzipSync(data);
    const rels = strFromU8(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', `<Relationship Id="rId8" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId9" Type="${REL}/calcChain" Target="calcChain.xml"/></Relationships>`);
    files['xl/_rels/workbook.xml.rels'] = strToU8(rels);
    files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="x"/></Types>'));
    return zipSync(files);
  }

  it('读取：共享字符串序号、公式、列定义、保留内容提示', () => {
    const b = open(sample());
    const s = b.sheets[0];
    expect(s.cells[0][0]).toMatchObject({ v: '名称', ss: 0 });
    expect(s.cells[1][0]).toMatchObject({ v: 4, f: 'B1*2' });
    expect(s.colDefs).toEqual([{ min: 0, max: 0, width: 20.5 }]);
    expect(s.rowDefs?.[0]).toEqual({ extra: 'x14ac:dyDescent="0.25"' });
    expect(b.keeps).toEqual(expect.arrayContaining(['图表']));
    expect(b.keeps).not.toContain('条件格式');
  });

  it('没改的部件逐字节保留；改过的表重写，富文本共享字符串沿用序号', () => {
    const src = sample();
    const b = open(src);
    edit(b, '改', (tx) => setInput(tx, 0, 0, 1, '新文字'));
    const out = saveXlsx(b);
    const a = unzipSync(src);
    const z = unzipSync(out);
    expect(z['xl/charts/chart1.xml']).toEqual(a['xl/charts/chart1.xml']);
    expect(z['xl/calcChain.xml']).toBeUndefined();
    expect(strFromU8(z['xl/_rels/workbook.xml.rels'])).not.toContain('calcChain');
    expect(strFromU8(z['[Content_Types].xml'])).not.toContain('calcChain');
    const ws = strFromU8(z['xl/worksheets/sheet1.xml']);
    expect(ws).toContain('<c r="B1" t="s"><v>2</v></c>');
    expect(ws).toContain('<c r="B2" t="s"><v>1</v></c>');
    expect(ws).toContain('<c r="A2"><f>B1*2</f><v>4</v></c>');
    expect(ws).toContain('<row r="1" x14ac:dyDescent="0.25">');
    expect(ws).toContain('<drawing r:id="rId1"/>');
    expect(ws).toContain('<cols><col min="1" max="1" width="20.5" customWidth="1"/></cols>');
    expect(strFromU8(z['xl/sharedStrings.xml'])).toBe(`<sst uniqueCount="3" xmlns="${MAIN}"><si><t>名称</t></si><si><r><rPr><b/></rPr><t>粗</t></r><r><t>体</t></r></si><si><t>新文字</t></si></sst>`);
    expect(strFromU8(z['xl/workbook.xml'])).toContain('<calcPr fullCalcOnLoad="1" calcId="191029"/>');
    // 再读一遍内容一致
    const again = readWorkbook(out);
    expect(again.sheets[0].cells[0][1]?.v).toBe('新文字');
    expect(again.sheets[0].cells[1][1]?.v).toBe('粗体');
  });

  it('插入行后条件格式、定义名称跟着移动；冻结窗格写入 sheetView', () => {
    const b = open(sample());
    edit(b, '插入', (tx) => {
      insertDelete(tx, 0, 'row', 0, 1);
      setFreeze(tx, 0, 1, 0);
    });
    const out = saveXlsx(b);
    const ws = part(out, 'xl/worksheets/sheet1.xml');
    expect(ws).toContain('<conditionalFormatting sqref="A3:B6"><cfRule type="expression" priority="1"><formula>$A3&gt;1</formula></cfRule></conditionalFormatting>');
    expect(ws).toContain('<c r="A3"><f>B2*2</f><v>4</v></c>');
    expect(ws).toContain('<dimension ref="A2:B3"/>');
    expect(ws).toContain('<sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>');
    expect(part(out, 'xl/workbook.xml')).toContain(`<definedName name="_xlnm.Print_Area" localSheetId="0">'S1'!$A$2:$B$6</definedName>`);
    markSaved(b, out);
    expect(b.sheets[0].shifts).toEqual([]);
    // 保存后再改：不会重复平移
    edit(b, '改', (tx) => setInput(tx, 0, 5, 5, 'x'));
    expect(part(saveXlsx(b), 'xl/worksheets/sheet1.xml')).toContain('sqref="A3:B6"');
  });

  it('条件格式：新建的写入文件（带差异格式），删掉的从原 XML 去掉，保存后再改不重复；撤销恢复', () => {
    const b = open(sample());
    const e = edit(b, '条件格式', (tx) => {
      addCfRule(tx, 0, { ranges: [{ r0: 0, c0: 1, r1: 1, c1: 1 }], type: 'cellIs', operator: 'greaterThan', formulas: ['1'], look: CF_LOOKS[0] });
      addCfRule(tx, 0, { ranges: [{ r0: 0, c0: 0, r1: 1, c1: 0 }], type: 'containsText', text: '钢"板', formulas: [] });
      removeCfRule(tx, 0, b.sheets[0].cf![0]);
    })!;
    const out = saveXlsx(b);
    const ws = part(out, 'xl/worksheets/sheet1.xml');
    expect(ws).not.toContain('$A2&gt;1');
    expect(ws).toContain('<conditionalFormatting sqref="B1:B2"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>');
    expect(ws).toContain('<cfRule type="containsText" priority="3" operator="containsText" text="钢&quot;板"><formula>NOT(ISERROR(SEARCH(&quot;钢&quot;&quot;板&quot;,A1)))</formula></cfRule>');
    expect(ws.indexOf('</sheetData>')).toBeLessThan(ws.indexOf('<conditionalFormatting'));
    expect(ws.indexOf('<conditionalFormatting')).toBeLessThan(ws.indexOf('<pageMargins'));
    expect(part(out, 'xl/styles.xml')).toContain('<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>');
    // 读回来：两条规则，差异格式对得上
    const again = readWorkbook(out);
    expect(again.sheets[0].cf?.map((r) => [r.type, r.dxf])).toEqual([['cellIs', 0], ['containsText', undefined]]);
    expect(again.styleBook!.dxfs![0]).toMatchObject({ fill: '#ffc7ce', color: '#9c0006' });
    // 保存后序号重排：再删掉第一条、保存，只剩文字规则
    markSaved(b, out);
    expect(b.sheets[0].cfRemoved).toBeUndefined();
    edit(b, '清除', (tx) => void clearCfRules(tx, 0, { r0: 0, c0: 1, r1: 0, c1: 1 }));
    const ws2 = part(saveXlsx(b), 'xl/worksheets/sheet1.xml');
    expect(ws2).not.toContain('type="cellIs"');
    expect(ws2.match(/<conditionalFormatting /g)).toHaveLength(1);
    // 撤销整个第一次改动：回到文件里原来的规则
    const c = open(sample());
    const e2 = edit(c, '条件格式', (tx) => void addCfRule(tx, 0, { ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], type: 'duplicateValues', formulas: [], look: CF_LOOKS[1] }))!;
    applyEntry(c, e2, -1);
    expect(c.sheets[0].cf).toHaveLength(1);
    expect(part(saveXlsx(c), 'xl/worksheets/sheet1.xml')).toContain('$A2&gt;1');
    expect(e).toBeTruthy();
  });

  it('新建工作簿：日期格式写进 styles，没有共享字符串时新建', () => {
    const b = blankWorkbook();
    edit(b, '输入', (tx) => {
      setInput(tx, 0, 0, 0, '2026-9-29');
      setInput(tx, 0, 0, 1, '采购');
    });
    const out = saveXlsx(b);
    const z = unzipSync(out);
    expect(strFromU8(z['xl/_rels/workbook.xml.rels'])).toContain('Target="sharedStrings.xml"');
    expect(strFromU8(z['[Content_Types].xml'])).toContain('/xl/sharedStrings.xml');
    const again = readWorkbook(out);
    expect(again.sheets[0].cells[0][0]?.t).toBe('2026/9/29');
    expect(again.sheets[0].cells[0][1]?.v).toBe('采购');
  });

  it('CSV：显示文字、需要时加引号', () => {
    const b = readCsv('a,b\n"x,y",2\n');
    edit(b, '改', (tx) => setInput(tx, 0, 1, 1, '1,234'));
    expect(saveCsv(b.sheets[0])).toBe('a,b\r\n"x,y","1,234"\r\n');
  });
});

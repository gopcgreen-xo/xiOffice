import { describe, expect, it } from 'vitest';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { blankPackage } from './save';
import { overTable, specOf, tableStyleAt, type SheetTable } from './tableStyle';
import { readWorkbook } from './workbook';
import { markSaved, saveXlsx } from './save';
import { setColumnFilter } from './sortFilter';
import { Tx } from './ops';

const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const table = (over: Partial<SheetTable> = {}): SheetTable => ({ path: 'xl/tables/table1.xml', ref: { r0: 0, c0: 0, r1: 5, c1: 3 }, headerRows: 1, totalsRows: 0, style: 'TableStyleMedium2', rowStripes: true, firstCol: false, lastCol: false, ...over });

describe('表格样式', () => {
  it('Medium2：标题行强调色底白色粗体，数据行从第一行起隔行淡色底', () => {
    const t = [table()];
    expect(tableStyleAt(t, THEME, 0, 1)).toMatchObject({ bold: true, fill: '#4472C4', color: '#FFFFFF' });
    expect(tableStyleAt(t, THEME, 1, 1)?.fill).toMatch(/^#[0-9A-F]{6}$/);
    expect(tableStyleAt(t, THEME, 1, 1)?.fill).not.toBe('#4472C4');
    expect(tableStyleAt(t, THEME, 2, 1)?.fill).toBeUndefined();
    expect(tableStyleAt(t, THEME, 3, 1)?.fill).toBe(tableStyleAt(t, THEME, 1, 1)?.fill);
  });

  it('表格以外、关闭隔行、汇总行、首列加粗、不认识的样式名', () => {
    expect(tableStyleAt([table()], THEME, 6, 0)).toBeUndefined();
    expect(tableStyleAt([table()], THEME, 0, 4)).toBeUndefined();
    expect(tableStyleAt([table({ rowStripes: false })], THEME, 1, 1)?.fill).toBeUndefined();
    expect(tableStyleAt([table({ totalsRows: 1 })], THEME, 5, 1)).toMatchObject({ bold: true });
    expect(tableStyleAt([table({ firstCol: true, rowStripes: false })], THEME, 2, 0)).toMatchObject({ bold: true });
    expect(tableStyleAt([table({ style: 'MyStyle' })], THEME, 0, 0)).toBeUndefined();
    expect(tableStyleAt([table({ style: undefined })], THEME, 0, 0)).toBeUndefined();
  });

  it('样式名：Light 只加粗标题、Dark 深色底白字', () => {
    expect(specOf('TableStyleLight1')).toMatchObject({ header: 'plain', accent: null });
    expect(specOf('TableStyleLight9')).toMatchObject({ header: 'solid', accent: 0 });
    expect(tableStyleAt([table({ style: 'TableStyleLight2' })], THEME, 0, 0)).toMatchObject({ bold: true });
    expect(tableStyleAt([table({ style: 'TableStyleDark2' })], THEME, 1, 0)?.color).toBe('#FFFFFF');
    expect(specOf('TableStyleMedium28')).not.toBeNull();
  });

  it('Medium 每行下面一条强调色淡线（Medium6 = 8EA9DB），自己设的边框优先', () => {
    const t = [table({ style: 'TableStyleMedium6' })];
    // Excel 的淡线是 8EA9DB，这里的明暗换算有 1 的取整误差
    for (const r of [0, 1, 2]) {
      const b = tableStyleAt(t, THEME.map((c, i) => (i === 8 ? '4472C4' : c)), r, 1)?.border?.b;
      expect(b).toMatchObject({ w: 1, style: 'solid' });
      expect(b?.color).toMatch(/^#8[EF][AB][9A]D[BC]$/);
    }
    const own = { border: { b: { w: 2, style: 'solid' as const, color: '#000000' } } };
    expect(overTable(tableStyleAt(t, THEME, 1, 1)!, own).border?.b?.w).toBe(2);
    expect(tableStyleAt([table({ style: 'TableStyleLight2' })], THEME, 1, 1)?.border).toBeUndefined();
  });

  it('单元格自己的格式盖在表格样式上，但默认的黑色字不盖掉标题行的白字', () => {
    const head = tableStyleAt([table()], THEME, 0, 0)!;
    expect(overTable(head, { color: '#000000', bold: true, align: 'center' })).toMatchObject({ bold: true, fill: '#4472C4', color: '#FFFFFF', align: 'center' });
    expect(overTable(head, { color: '#FF0000', fill: '#00FF00' })).toMatchObject({ color: '#FF0000', fill: '#00FF00' });
  });

  it('读取 xl/tables：范围、标题行数、样式名、隔行', () => {
    const files = blankPackage(['S1']);
    files['xl/worksheets/sheet1.xml'] = strToU8(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>',
    );
    files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
    );
    files['xl/tables/table1.xml'] = strToU8(
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="表1" ref="A1:G20"><autoFilter ref="A1:G20"/><tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>',
    );
    const b = readWorkbook(zipSync(files));
    expect(b.sheets[0].tables).toEqual([{ path: 'xl/tables/table1.xml', ref: { r0: 0, c0: 0, r1: 19, c1: 6 }, headerRows: 1, totalsRows: 0, style: 'TableStyleMedium2', rowStripes: true, firstCol: false, lastCol: false, filter: true }]);
  });

  it('表格自带的筛选：像自动筛选一样使用，条件写回表格部件，工作表不写 autoFilter', () => {
    const files = blankPackage(['S1']);
    const rows = [['名称', '类别'], ['钢板', 'A'], ['螺栓', 'B'], ['钢管', 'A']]
      .map((row, r) => `<row r="${r + 1}">${row.map((v, c) => `<c r="${'AB'[c]}${r + 1}" t="inlineStr"><is><t>${v}</t></is></c>`).join('')}</row>`)
      .join('');
    files['xl/worksheets/sheet1.xml'] = strToU8(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${rows}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`,
    );
    files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
    );
    files['xl/tables/table1.xml'] = strToU8(
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="表1" ref="A1:B1048575"><autoFilter ref="A1:B1048575"/><tableColumns count="2"><tableColumn id="1" name="名称"/><tableColumn id="2" name="类别"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>',
    );
    const src = zipSync(files);
    const b = readWorkbook(src);
    b.source = src;
    // 筛选区域按数据范围截短，标记为表格自带
    expect(b.sheets[0].autoFilter).toMatchObject({ table: true, ref: { r0: 0, c0: 0, r1: 3, c1: 1 } });
    const tx = new Tx(b, '筛选');
    setColumnFilter(tx, 0, 1, { values: ['A'] });
    tx.commit();
    expect(b.sheets[0].rowDefs?.[2]?.hidden).toBe(true);
    const out = saveXlsx(b);
    const z = unzipSync(out);
    const table = new TextDecoder().decode(z['xl/tables/table1.xml']);
    expect(table).toContain('<autoFilter ref="A1:B1048575"><filterColumn colId="1"><filters><filter val="A"/></filters></filterColumn></autoFilter><tableColumns');
    const ws = new TextDecoder().decode(z['xl/worksheets/sheet1.xml']);
    expect(ws).not.toContain('<autoFilter');
    expect(ws).toContain('<tablePart r:id="rId1"/>');
    expect(new TextDecoder().decode(z['xl/workbook.xml'])).not.toContain('_FilterDatabase');
    // 读回来：条件和隐藏行都在
    const again = readWorkbook(out);
    expect(again.sheets[0].autoFilter?.cols[1]).toEqual({ values: ['A'] });
    // 清除条件后保存：表格的筛选按钮还在，条件没了
    markSaved(b, out);
    b.source = out;
    const tx2 = new Tx(b, '清除');
    setColumnFilter(tx2, 0, 1, null);
    tx2.commit();
    const table2 = new TextDecoder().decode(unzipSync(saveXlsx(b))['xl/tables/table1.xml']);
    expect(table2).toContain('<autoFilter ref="A1:B1048575"/>');
  });
});

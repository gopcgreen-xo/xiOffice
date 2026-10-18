import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { applyEntry, insertDelete, setInput, Tx } from './ops';
import { blankPackage, markSaved, saveXlsx } from './save';
import { addSheet, copySheet, deleteSheet, moveSheet, renameSheet, setSheetHidden, setTabColor, sheetNameError, uniqueSheetName } from './sheets';
import { columnExtras, columnValues, compareCells, currentRegion, guessHeader, hiddenByOthers, passes, setAutoFilter, setColumnFilter, sortRange } from './sortFilter';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const edit = (b: Workbook, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, 'x');
  fn(tx);
  return tx.commit()!;
};
const fill = (b: Workbook, si: number, rows: string[][]) => edit(b, (tx) => rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, si, r, c, v))));
const col = (b: Workbook, si: number, c: number) => b.sheets[si].cells.map((row) => row?.[c]?.v ?? '');

/** 两张表的 xlsx：S1 有图表关系、条件格式，定义名称一个全局一个属于 S2 */
function twoSheets(): Workbook {
  const files = blankPackage(['S1', 'S2']);
  files['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting><drawing r:id="rId1"/></worksheet>`);
  files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8('<Relationships/>');
  files['xl/worksheets/sheet2.xml'] = strToU8(`<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><f>S1!A1*2</f><v>2</v></c></row></sheetData></worksheet>`);
  files['xl/workbook.xml'] = strToU8(
    strFromU8(files['xl/workbook.xml']).replace('</sheets>', '</sheets><definedNames><definedName name="单价">S1!$A$1</definedName><definedName name="本表" localSheetId="1">S2!$A$1</definedName></definedNames>'),
  );
  const data = zipSync(files);
  const b = readWorkbook(data);
  b.source = data;
  return b;
}

describe('工作表管理', () => {
  it('表名校验与自动命名', () => {
    const b = blankWorkbook();
    expect(sheetNameError(b, 'a/b')).toMatch(/不能包含/);
    expect(sheetNameError(b, 'sheet1')).toBe('已有同名的工作表');
    expect(sheetNameError(b, 'sheet1', 0)).toBeNull();
    expect(sheetNameError(b, 'x'.repeat(32))).toMatch(/31/);
    expect(uniqueSheetName(b)).toBe('Sheet2');
    expect(uniqueSheetName(b, 'Sheet1')).toBe('Sheet1 (2)');
  });

  it('改名后公式和定义名称跟着改；删除后引用变 #REF!；撤销恢复', () => {
    const b = twoSheets();
    edit(b, (tx) => renameSheet(tx, 0, '明细 表'));
    expect(b.sheets[1].cells[0][0]?.f).toBe("'明细 表'!A1*2");
    expect(b.names?.[0].f).toBe("'明细 表'!$A$1");
    const e = edit(b, (tx) => deleteSheet(tx, 0));
    expect(b.sheets.map((s) => s.name)).toEqual(['S2']);
    expect(b.sheets[0].cells[0][0]?.f).toBe("'明细 表'!#REF!*2");
    expect(b.names).toEqual([{ name: '单价', f: "'明细 表'!#REF!" }, { name: '本表', f: 'S2!$A$1', local: 0 }]);
    applyEntry(b, e, -1);
    expect(b.sheets.map((s) => s.name)).toEqual(['明细 表', 'S2']);
    expect(b.names?.[1].local).toBe(1);
  });

  it('不能删除 / 隐藏最后一张可见的表', () => {
    const b = blankWorkbook();
    const tx = new Tx(b, 'x');
    expect(deleteSheet(tx, 0)).toBe(false);
    expect(setSheetHidden(tx, 0, true)).toBe(false);
  });

  it('保存：改名、移动、隐藏、复制、新增、删除都写进 workbook.xml，部件与关系一致', () => {
    const b = twoSheets();
    edit(b, (tx) => {
      renameSheet(tx, 0, '明细');
      copySheet(tx, 0); // 明细 (2)
      addSheet(tx, 3, '新表');
      moveSheet(tx, 1, 0); // 明细 (2), 明细, S2, 新表
      setSheetHidden(tx, 2, true);
      setTabColor(tx, 3, '#ff0000');
    });
    edit(b, (tx) => setInput(tx, 3, 0, 0, '新内容'));
    const out = saveXlsx(b);
    const z = unzipSync(out);
    const wb = strFromU8(z['xl/workbook.xml']);
    const sheets = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => m[0]);
    expect(sheets.map((t) => t.match(/name="([^"]*)"/)![1])).toEqual(['明细 (2)', '明细', 'S2', '新表']);
    expect(sheets[2]).toContain('state="hidden"');
    expect(wb).toContain('<definedName name="单价">明细!$A$1</definedName>');
    expect(wb).toContain('<definedName name="本表" localSheetId="2">S2!$A$1</definedName>');
    // 复制的表以原表为底：条件格式保留、图表关系去掉
    const rels = strFromU8(z['xl/_rels/workbook.xml.rels']);
    const target = (name: string) => {
      const rid = sheets.find((t) => t.includes(`name="${name}"`))!.match(/r:id="([^"]*)"/)![1];
      return `xl/${rels.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]*)"`))![1]}`;
    };
    const copy = strFromU8(z[target('明细 (2)')]);
    expect(copy).toContain('<conditionalFormatting sqref="A1">');
    expect(copy).not.toContain('<drawing');
    expect(strFromU8(z[target('新表')])).toContain('<tabColor rgb="FFFF0000"/>');
    expect(strFromU8(z['[Content_Types].xml']).match(/worksheet\+xml/g)?.length).toBe(4);
    // 再读回来
    const again = readWorkbook(out);
    expect(again.sheets.map((s) => [s.name, s.hidden ?? '', s.tabColor ?? ''])).toEqual([
      ['明细 (2)', '', ''],
      ['明细', '', ''],
      ['S2', 'hidden', ''],
      ['新表', '', '#ff0000'],
    ]);
    expect(again.sheets[3].cells[0][0]?.v).toBe('新内容');
    expect(again.sheets[2].cells[0][0]?.f).toBe('明细!A1*2');
    // 保存后再删一张：部件、关系、内容类型都去掉
    markSaved(b, out);
    edit(b, (tx) => deleteSheet(tx, 0));
    const z2 = unzipSync(saveXlsx(b));
    expect(strFromU8(z2['[Content_Types].xml']).match(/worksheet\+xml/g)?.length).toBe(3);
    expect(Object.keys(z2).filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k)).length).toBe(3);
    expect(strFromU8(z2['xl/workbook.xml'])).toContain('localSheetId="1">S2!$A$1');
  });
});

describe('排序', () => {
  it('比较：数字 < 文字 < 布尔，空白总在最后，中文按拼音', () => {
    const c = (v: string | number | boolean) => ({ v, t: String(v) });
    const list = [c('张三'), undefined, c(5), c('李四'), c(true), c(2), c('阿五')];
    expect(list.sort((a, b) => compareCells(a, b)).map((x) => x?.v ?? '')).toEqual([2, 5, '阿五', '李四', '张三', true, '']);
    expect([c(1), undefined, c(3)].sort((a, b) => compareCells(a, b, true)).map((x) => x?.v ?? '')).toEqual([3, 1, '']);
  });

  it('当前区域、标题判断、多列排序，公式随行平移', () => {
    const b = blankWorkbook();
    fill(b, 0, [
      ['名称', '类别', '数量', '小计'],
      ['螺栓', '五金', '30', '=C2*2'],
      ['钢板', '钢材', '10', '=C3*2'],
      ['垫片', '五金', '20', '=C4*2'],
    ]);
    const s = b.sheets[0];
    const g = currentRegion(s, 2, 1);
    expect(g).toEqual({ r0: 0, c0: 0, r1: 3, c1: 3 });
    expect(guessHeader(s, g)).toBe(true);
    edit(b, (tx) => sortRange(tx, 0, g, [{ c: 1, desc: false }, { c: 2, desc: true }], true));
    // 类别按拼音：钢材（gang）在五金（wu）前；同类别按数量降序
    expect(col(b, 0, 0)).toEqual(['名称', '钢板', '螺栓', '垫片']);
    expect(col(b, 0, 3).length).toBe(4);
    expect(b.sheets[0].cells.map((r) => r?.[3]?.f ?? '')).toEqual(['', 'C2*2', 'C3*2', 'C4*2']);
  });
});

describe('自动筛选', () => {
  it('勾选值、自定义条件；隐藏行；关闭后全部显示；保存往返', () => {
    const b = blankWorkbook();
    fill(b, 0, [
      ['名称', '数量'],
      ['螺栓', '30'],
      ['钢板', '10'],
      ['', '20'],
      ['螺母', '5'],
    ]);
    const ref = { r0: 0, c0: 0, r1: 4, c1: 1 };
    edit(b, (tx) => setAutoFilter(tx, 0, ref));
    expect(columnValues(b.sheets[0], ref, 0)).toEqual([
      { text: '钢板', count: 1 },
      { text: '螺母', count: 1 },
      { text: '螺栓', count: 1 },
      { text: '', count: 1 },
    ]);
    edit(b, (tx) => setColumnFilter(tx, 0, 0, { values: ['螺栓', '螺母', ''] }));
    edit(b, (tx) => setColumnFilter(tx, 0, 1, { custom: { and: true, rules: [{ op: 'greaterThanOrEqual', val: '10' }] } }));
    const hidden = () => [1, 2, 3, 4].filter((r) => b.sheets[0].rowDefs?.[r]?.hidden);
    expect(hidden()).toEqual([2, 4]);
    const again = readWorkbook(saveXlsx(b));
    expect(again.sheets[0].autoFilter).toEqual(b.sheets[0].autoFilter);
    expect([1, 2, 3, 4].filter((r) => again.sheets[0].rowDefs?.[r]?.hidden)).toEqual([2, 4]);
    expect(passes({ v: '螺栓 M8', t: '螺栓 M8' }, { custom: { and: false, rules: [{ op: 'equal', val: '*M8*' }] } })).toBe(true);
    edit(b, (tx) => setAutoFilter(tx, 0, null));
    expect(hidden()).toEqual([]);
    expect(b.sheets[0].autoFilter).toBeUndefined();
  });

  it('第二次筛选在第一次的结果里选：下拉列表只列满足其它列条件的行，本列自己的条件不算', () => {
    const b = blankWorkbook();
    fill(b, 0, [
      ['名称', '规格'],
      ['螺栓', 'M8'],
      ['螺栓', 'M10'],
      ['螺母', 'M12'],
    ]);
    const ref = { r0: 0, c0: 0, r1: 3, c1: 1 };
    edit(b, (tx) => setAutoFilter(tx, 0, ref));
    edit(b, (tx) => setColumnFilter(tx, 0, 0, { values: ['螺栓'] }));
    const skip1 = hiddenByOthers(b, 0, 1);
    expect(columnValues(b.sheets[0], ref, 1, skip1).map((x) => x.text)).toEqual(['M8', 'M10']);
    expect([...columnExtras(b, 0, ref, 1, skip1).fills.values()]).toEqual([2]);
    // 第一列自己的下拉列表仍列出全部值（可以改第一次的条件）
    expect(columnValues(b.sheets[0], ref, 0, hiddenByOthers(b, 0, 0)).map((x) => x.text)).toEqual(['螺母', '螺栓']);
    edit(b, (tx) => setColumnFilter(tx, 0, 1, { values: ['M8'] }));
    expect([1, 2, 3].filter((r) => b.sheets[0].rowDefs?.[r]?.hidden)).toEqual([2, 3]);
  });

  it('插入行时筛选区域跟着扩大', () => {
    const b = blankWorkbook();
    fill(b, 0, [['a'], ['1'], ['2']]);
    edit(b, (tx) => setAutoFilter(tx, 0, { r0: 0, c0: 0, r1: 2, c1: 0 }));
    edit(b, (tx) => insertDelete(tx, 0, 'row', 2, 1));
    expect(b.sheets[0].autoFilter?.ref).toEqual({ r0: 0, c0: 0, r1: 3, c1: 0 });
  });
});

import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { clearGroups, groupOf, groupRange, isSummary, levelAt, maxLevel, showLevel, toggleGroup, ungroupRange } from './outline';
import { insertDelete, setColWidth, Tx } from './ops';
import { blankPackage, saveXlsx } from './save';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, '测试');
  fn(tx);
  return tx.commit();
};
const hiddenRows = (b: Workbook) => Object.entries(b.sheets[0].rowDefs ?? {}).filter(([, d]) => d.hidden).map(([k]) => Number(k));

describe('行分组', () => {
  it('组合 / 取消组合：级别加减，汇总行在明细下方', () => {
    const b = blankWorkbook();
    edit(b, (tx) => groupRange(tx, 0, 'row', 1, 3));
    expect([0, 1, 2, 3, 4].map((r) => levelAt(b.sheets[0], 'row', r))).toEqual([0, 1, 1, 1, 0]);
    expect(maxLevel(b.sheets[0], 'row')).toBe(1);
    expect(isSummary(b.sheets[0], 'row', 4)).toBe(true);
    expect(isSummary(b.sheets[0], 'row', 2)).toBe(false);
    expect(groupOf(b.sheets[0], 'row', 4)).toEqual([1, 3]);
    edit(b, (tx) => groupRange(tx, 0, 'row', 2, 3));
    expect(levelAt(b.sheets[0], 'row', 2)).toBe(2);
    edit(b, (tx) => ungroupRange(tx, 0, 'row', 2, 3));
    expect(levelAt(b.sheets[0], 'row', 2)).toBe(1);
    edit(b, (tx) => ungroupRange(tx, 0, 'row', 1, 3));
    expect(b.sheets[0].rowDefs).toEqual({});
  });

  it('折叠 / 展开：汇总行记 collapsed，明细行隐藏，显示高度为 0', () => {
    const b = blankWorkbook();
    edit(b, (tx) => groupRange(tx, 0, 'row', 1, 3));
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 4));
    expect(hiddenRows(b)).toEqual([1, 2, 3]);
    expect(b.sheets[0].rowDefs?.[4]?.collapsed).toBe(true);
    expect(b.sheets[0].rowHeights[2]).toBe(0);
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 4));
    expect(hiddenRows(b)).toEqual([]);
    expect(b.sheets[0].rowDefs?.[4]).toBeUndefined();
  });

  it('嵌套：外层折叠藏起内层；内层展开状态保留；级别按钮', () => {
    const b = blankWorkbook();
    // 行 1–6 一级，其中 2–3 二级；汇总行：4（二级组之后）、7（一级组之后）
    edit(b, (tx) => groupRange(tx, 0, 'row', 1, 6));
    edit(b, (tx) => groupRange(tx, 0, 'row', 1, 2));
    expect(groupOf(b.sheets[0], 'row', 3)).toEqual([1, 2]);
    expect(groupOf(b.sheets[0], 'row', 7)).toEqual([1, 6]);
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 3));
    expect(hiddenRows(b)).toEqual([1, 2]);
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 7));
    expect(hiddenRows(b)).toEqual([1, 2, 3, 4, 5, 6]);
    // 展开外层：内层仍是折叠的
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 7));
    expect(hiddenRows(b)).toEqual([1, 2]);
    // 级别按钮：1 = 只看未分组，2 = 一级，3 = 全部
    edit(b, (tx) => showLevel(tx, 0, 'row', 1));
    expect(hiddenRows(b)).toEqual([1, 2, 3, 4, 5, 6]);
    edit(b, (tx) => showLevel(tx, 0, 'row', 2));
    expect(hiddenRows(b)).toEqual([1, 2]);
    edit(b, (tx) => showLevel(tx, 0, 'row', 3));
    expect(hiddenRows(b)).toEqual([]);
  });

  it('汇总行在上方（summaryBelow = false）：组在汇总行之后', () => {
    const b = blankWorkbook();
    b.sheets[0].outline = { summaryBelow: false, summaryRight: true };
    edit(b, (tx) => groupRange(tx, 0, 'row', 2, 4));
    expect(groupOf(b.sheets[0], 'row', 1)).toEqual([2, 4]);
    expect(isSummary(b.sheets[0], 'row', 5)).toBe(false);
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 1));
    expect(hiddenRows(b)).toEqual([2, 3, 4]);
  });

  it('取消组合 / 清除分组会把藏起来的行显示出来；在组里插入行仍在组里', () => {
    const b = blankWorkbook();
    edit(b, (tx) => groupRange(tx, 0, 'row', 1, 3));
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 4));
    edit(b, (tx) => clearGroups(tx, 0, 'row', 0, 5));
    expect(b.sheets[0].rowDefs).toEqual({});
    edit(b, (tx) => groupRange(tx, 0, 'row', 1, 3));
    edit(b, (tx) => insertDelete(tx, 0, 'row', 2, 1));
    expect([1, 2, 3, 4].map((r) => levelAt(b.sheets[0], 'row', r))).toEqual([1, 1, 1, 1]);
  });
});

describe('列分组', () => {
  it('组合、折叠、列宽修改后分组保留', () => {
    const b = blankWorkbook();
    edit(b, (tx) => groupRange(tx, 0, 'col', 1, 2));
    expect([0, 1, 2, 3].map((c) => levelAt(b.sheets[0], 'col', c))).toEqual([0, 1, 1, 0]);
    edit(b, (tx) => setColWidth(tx, 0, 1, 1, { width: 140 }));
    expect(levelAt(b.sheets[0], 'col', 1)).toBe(1);
    edit(b, (tx) => toggleGroup(tx, 0, 'col', 3));
    expect(b.sheets[0].colWidths[1]).toBe(0);
    expect(b.sheets[0].colWidths[2]).toBe(0);
    edit(b, (tx) => toggleGroup(tx, 0, 'col', 3));
    expect(b.sheets[0].colWidths[1]).toBeGreaterThan(0);
  });
});

describe('分组的读写', () => {
  it('保存：行 / 列的 outlineLevel、collapsed 与 sheetFormatPr 的最大级别；重新读出来一致', () => {
    const b = blankWorkbook();
    edit(b, (tx) => groupRange(tx, 0, 'row', 1, 3));
    edit(b, (tx) => groupRange(tx, 0, 'col', 1, 2));
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 4));
    const xml = strFromU8(unzipSync(saveXlsx(b))['xl/worksheets/sheet1.xml']);
    expect(xml).toMatch(/<sheetFormatPr[^>]*outlineLevelRow="1"/);
    expect(xml).toMatch(/<sheetFormatPr[^>]*outlineLevelCol="1"/);
    expect(xml).toContain('<row r="2" hidden="1" outlineLevel="1"');
    expect(xml).toContain('collapsed="1"');
    expect(xml).toMatch(/<col min="2" max="3"[^>]*outlineLevel="1"/);
    const back = readWorkbook(saveXlsx(b));
    expect([1, 2, 3].map((r) => back.sheets[0].rowDefs?.[r]?.level)).toEqual([1, 1, 1]);
    expect(back.sheets[0].rowDefs?.[4]?.collapsed).toBe(true);
    expect(back.sheets[0].rowDefs?.[1]?.hidden).toBe(true);
    expect(levelAt(back.sheets[0], 'col', 2)).toBe(1);
  });

  it('读 Excel 写的分组（含 outlinePr summaryBelow=0）；没改动保存不重写工作表', () => {
    const files = blankPackage(['S1']);
    files['xl/worksheets/sheet1.xml'] = strToU8(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryBelow="0"/></sheetPr><dimension ref="A1"/>' +
        '<sheetFormatPr defaultRowHeight="15" outlineLevelRow="2"/><sheetData>' +
        '<row r="1" collapsed="1"><c r="A1"><v>1</v></c></row><row r="2" outlineLevel="1" hidden="1"><c r="A2"><v>2</v></c></row>' +
        '<row r="3" outlineLevel="2" hidden="1"/></sheetData></worksheet>',
    );
    const data = zipSync(files);
    const b = readWorkbook(data);
    b.source = data;
    expect(b.sheets[0].outline).toEqual({ summaryBelow: false, summaryRight: true });
    expect(groupOf(b.sheets[0], 'row', 0)).toEqual([1, 2]);
    expect(b.sheets[0].rowDefs?.[0]?.collapsed).toBe(true);
    expect(Array.from(saveXlsx(b))).toEqual(Array.from(data));
    edit(b, (tx) => toggleGroup(tx, 0, 'row', 0));
    expect(hiddenRows(b)).toEqual([]);
    const xml = strFromU8(unzipSync(saveXlsx(b))['xl/worksheets/sheet1.xml']);
    expect(xml).toContain('summaryBelow="0"');
    expect(xml).toContain('outlineLevelRow="2"');
  });
});

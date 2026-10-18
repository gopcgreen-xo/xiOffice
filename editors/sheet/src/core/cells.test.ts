import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { shiftFormula } from './formula/refs';
import { applyEntry, cellShiftBlocked, copyBlock, insertCopied, insertCut, retargetFormula, insertDelete, mergeCells, setColWidth, setInput, setRowHeight, Tx } from './ops';
import { blankPackage, saveXlsx } from './save';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, label: string, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, label);
  fn(tx);
  return tx.commit();
};

describe('插入 / 删除单元格', () => {
  it('引用：另一方向完全在 band 里才移动', () => {
    const spec = { sheet: 'S', axis: 'row' as const, at: 1, n: 1, band: [1, 2] as [number, number] };
    expect(shiftFormula('B2+C3+A2+B2:C5+A2:C5+B:B', 'S', spec)).toBe('B3+C4+A2+B3:C6+A2:C5+B:B');
  });

  it('活动单元格下移：只移动 band 列里的格子，公式跟着；撤销复原', () => {
    const b = blankWorkbook();
    edit(b, '准备', (tx) => {
      setInput(tx, 0, 0, 0, '1'); // A1
      setInput(tx, 0, 0, 1, '2'); // B1
      setInput(tx, 0, 1, 1, '3'); // B2
      setInput(tx, 0, 0, 3, '=B1+B2+A1'); // D1
    });
    const e = edit(b, '插入单元格', (tx) => insertDelete(tx, 0, 'row', 0, 1, [1, 1]))!;
    const s = b.sheets[0];
    expect(s.cells[0][0]?.v).toBe(1); // A1 不动
    expect(s.cells[0][1]).toBeUndefined();
    expect(s.cells[1][1]?.v).toBe(2);
    expect(s.cells[2][1]?.v).toBe(3);
    expect(s.cells[0][3]?.f).toBe('B2+B3+A1');
    expect(s.shifts).toEqual([{ axis: 'row', at: 0, n: 1, band: [1, 1] }]);
    applyEntry(b, e, -1);
    expect(b.sheets[0].cells[0][1]?.v).toBe(2);
    expect(b.sheets[0].cells[0][3]?.f).toBe('B1+B2+A1');
  });

  it('删除单元格、右侧左移：被删的引用变 #REF!', () => {
    const b = blankWorkbook();
    edit(b, '准备', (tx) => {
      setInput(tx, 0, 0, 0, '1'); // A1
      setInput(tx, 0, 0, 1, '2'); // B1
      setInput(tx, 0, 0, 2, '3'); // C1
      setInput(tx, 0, 1, 0, '=A1+C1'); // A2
    });
    edit(b, '删除单元格', (tx) => insertDelete(tx, 0, 'col', 0, -1, [0, 0]));
    const s = b.sheets[0];
    expect([s.cells[0][0]?.v, s.cells[0][1]?.v, s.cells[0][2]]).toEqual([2, 3, undefined]);
    expect(s.cells[1][0]?.f).toBe('#REF!+B1');
  });

  it('合并区域横跨 band 的边时不能移动', () => {
    const b = blankWorkbook();
    edit(b, '合并', (tx) => mergeCells(tx, 0, { r0: 3, c0: 0, r1: 3, c1: 2 }, true));
    expect(cellShiftBlocked(b.sheets[0], 'row', 1, 1, [1, 1])).toMatch(/合并单元格/);
    expect(cellShiftBlocked(b.sheets[0], 'row', 1, 1, [0, 2])).toBeNull();
    expect(cellShiftBlocked(b.sheets[0], 'row', 5, 1, [1, 1])).toBeNull(); // 合并区域在插入处上方，不动
  });

  it('保存：条件格式和批注只在 band 里的跟着移动', () => {
    const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const files = blankPackage(['S1']);
    files['xl/worksheets/sheet1.xml'] = strToU8(
      `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
        `<conditionalFormatting sqref="A2:A3"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>` +
        `<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" priority="2" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting></worksheet>`,
    );
    const data = zipSync(files);
    const b = readWorkbook(data);
    b.source = data;
    edit(b, '插入单元格', (tx) => insertDelete(tx, 0, 'row', 0, 2, [0, 0]));
    const xml = strFromU8(unzipSync(saveXlsx(b))['xl/worksheets/sheet1.xml']);
    expect(xml).toContain('sqref="A4:A5"');
    expect(xml).toContain('sqref="B2:B3"');
  });
});

describe('插入复制的单元格', () => {
  const prep = () => {
    const b = blankWorkbook();
    edit(b, '准备', (tx) => {
      setInput(tx, 0, 0, 0, 'a1');
      setInput(tx, 0, 1, 0, 'a2');
      setInput(tx, 0, 2, 0, '5'); // A3
      setInput(tx, 0, 2, 1, '=A3*2'); // B3
      setInput(tx, 0, 3, 0, 'a4');
      setInput(tx, 0, 0, 2, '=A3'); // C1 引用 A3
    });
    return b;
  };

  it('整行：在目标行上方插入同样多的行再粘贴，公式按位置平移，下面的行和引用跟着下移；撤销复原', () => {
    const b = prep();
    edit(b, '行高', (tx) => setRowHeight(tx, 0, 2, 2, { height: 40 }));
    const block = copyBlock(b, 0, { r0: 2, c0: 0, r1: 2, c1: 5 }); // 复制第 3 行
    const e = edit(b, '插入复制的单元格', (tx) => insertCopied(tx, 0, 1, 0, block, 'rows'))!; // 插到第 2 行
    const s = b.sheets[0];
    expect(s.cells[1][0]?.v).toBe(5);
    expect(s.cells[1][1]?.f).toBe('A2*2');
    expect(s.cells[2][0]?.v).toBe('a2');
    expect(s.cells[3][1]?.f).toBe('A4*2');
    expect(s.cells[0][2]?.f).toBe('A4');
    expect(s.rowDefs?.[1]?.ht).toBe(30);
    applyEntry(b, e, -1);
    expect(b.sheets[0].cells[1][0]?.v).toBe('a2');
    expect(b.sheets[0].cells[0][2]?.f).toBe('A3');
    expect(b.sheets[0].rowDefs?.[1]).toBeUndefined();
  });

  it('整列：插入同样多的列再粘贴，列宽跟着来源', () => {
    const b = prep();
    edit(b, '列宽', (tx) => setColWidth(tx, 0, 0, 0, { width: 140 }));
    const block = copyBlock(b, 0, { r0: 0, c0: 0, r1: 9, c1: 0 }); // 复制 A 列
    edit(b, '插入复制的单元格', (tx) => insertCopied(tx, 0, 0, 2, block, 'cols')); // 插到 C 列
    const s = b.sheets[0];
    expect(s.cells[0][2]?.v).toBe('a1');
    expect(s.cells[0][3]?.f).toBe('A3'); // 原 C1 右移
    expect(s.colDefs?.find((d) => d.min <= 2 && d.max >= 2)?.width).toBe(20);
  });

  it('部分单元格下移：只移动块覆盖的列', () => {
    const b = prep();
    const block = copyBlock(b, 0, { r0: 0, c0: 0, r1: 0, c1: 0 }); // A1
    edit(b, '插入复制的单元格', (tx) => insertCopied(tx, 0, 1, 0, block, 'down'));
    const s = b.sheets[0];
    expect(s.cells[1][0]?.v).toBe('a1');
    expect(s.cells[2][0]?.v).toBe('a2');
    expect(s.cells[2][1]?.f).toBe('A4*2'); // B3 没动，引用的 A3 下移到 A4
  });
});

describe('插入剪切的单元格', () => {
  const prep = () => {
    const b = blankWorkbook();
    edit(b, '准备', (tx) => {
      ['r1', 'r2', 'r3', 'r4', 'r5'].forEach((v, r) => setInput(tx, 0, r, 0, v)); // A1:A5
      setInput(tx, 0, 3, 1, '=A4&"!"'); // B4 引用本行
      setInput(tx, 0, 0, 2, '=B4'); // C1 引用被移动的格子
      setInput(tx, 0, 1, 2, '=SUM(A1:A5)'); // C2 区域
    });
    return b;
  };
  const col = (b: Workbook, c = 0) => b.sheets[0].cells.map((row) => row?.[c]?.v);

  it('整行上移：第 4 行剪切插到第 2 行上方，引用跟着走；撤销复原', () => {
    const b = prep();
    edit(b, '行高', (tx) => setRowHeight(tx, 0, 3, 3, { height: 40 }));
    let out: unknown;
    const e = edit(b, '插入剪切的单元格', (tx) => (out = insertCut(tx, 0, { r0: 3, r1: 3, c0: 0, c1: 9 }, 1, 4, 'rows')))!;
    expect(out).toEqual({ r0: 1, r1: 1, c0: 0, c1: 9 });
    const s = b.sheets[0];
    expect(col(b).slice(0, 5)).toEqual(['r1', 'r4', 'r2', 'r3', 'r5']);
    expect(s.cells[1][1]?.f).toBe('A2&"!"');
    expect(s.cells[0][2]?.f).toBe('B2');
    expect(s.cells[2]?.[2]?.f).toBe('SUM(A1:A5)'); // C2 → C3，区域不变
    expect(s.rowDefs?.[1]?.ht).toBe(30);
    expect(s.rowDefs?.[3]?.ht).toBeUndefined();
    applyEntry(b, e, -1);
    expect(col(b).slice(0, 5)).toEqual(['r1', 'r2', 'r3', 'r4', 'r5']);
    expect(b.sheets[0].cells[0][2]?.f).toBe('B4');
  });

  it('整行下移：第 1 行剪切插到第 4 行上方，落在第 3 行', () => {
    const b = prep();
    let out: unknown;
    edit(b, '插入剪切的单元格', (tx) => (out = insertCut(tx, 0, { r0: 0, r1: 0, c0: 0, c1: 9 }, 3, 0, 'rows')));
    expect(out).toEqual({ r0: 2, r1: 2, c0: 0, c1: 9 });
    expect(col(b).slice(0, 5)).toEqual(['r2', 'r3', 'r1', 'r4', 'r5']);
    expect(b.sheets[0].cells[2][2]?.f).toBe('B4'); // C1 跟着到 C3，B4 所在的行没动
  });

  it('目标就是原位置或紧挨着：不动', () => {
    const b = prep();
    for (const at of [3, 4]) {
      let out: unknown;
      edit(b, '插入剪切的单元格', (tx) => (out = insertCut(tx, 0, { r0: 3, r1: 3, c0: 0, c1: 9 }, at, 0, 'rows')));
      expect(out).toEqual({ r0: 3, r1: 3, c0: 0, c1: 9 });
    }
    expect(col(b).slice(0, 5)).toEqual(['r1', 'r2', 'r3', 'r4', 'r5']);
  });

  it('整列：A 列剪切插到 D 列左侧', () => {
    const b = prep();
    edit(b, '列宽', (tx) => setColWidth(tx, 0, 0, 0, { width: 140 }));
    let out: unknown;
    edit(b, '插入剪切的单元格', (tx) => (out = insertCut(tx, 0, { r0: 0, r1: 99, c0: 0, c1: 0 }, 0, 3, 'cols')));
    expect(out).toEqual({ r0: 0, r1: 99, c0: 2, c1: 2 });
    const s = b.sheets[0];
    expect(s.cells[0][2]?.v).toBe('r1');
    expect(s.cells[3][0]?.f).toBe('C4&"!"'); // B4 → A4，指向移走的 A4 → C4
    expect(s.cells[0][1]?.f).toBe('A4'); // C1 → B1，B4 → A4
    expect(s.cells[1][1]?.f).toBe('SUM(C1:C5)');
    expect(s.colDefs?.find((d) => d.min <= 2 && d.max >= 2)?.width).toBe(20);
    expect(s.colDefs?.find((d) => d.min <= 0 && d.max >= 0)?.width).toBeUndefined();
  });

  it('部分单元格下移：同一列里移动', () => {
    const b = prep();
    let out: unknown;
    edit(b, '插入剪切的单元格', (tx) => (out = insertCut(tx, 0, { r0: 4, r1: 4, c0: 0, c1: 0 }, 0, 0, 'down')));
    expect(out).toEqual({ r0: 0, r1: 0, c0: 0, c1: 0 });
    expect(col(b).slice(0, 5)).toEqual(['r5', 'r1', 'r2', 'r3', 'r4']);
    expect(b.sheets[0].cells[3][1]?.f).toBe('A5&"!"'); // B 列不动，A4 下移到 A5
  });

  it('部分单元格下移到别的列：原位置下面的格子上移补上', () => {
    const b = prep();
    let out: unknown;
    edit(b, '插入剪切的单元格', (tx) => (out = insertCut(tx, 0, { r0: 0, r1: 0, c0: 0, c1: 0 }, 0, 4, 'down')));
    expect(out).toEqual({ r0: 0, r1: 0, c0: 4, c1: 4 });
    expect(col(b).slice(0, 4)).toEqual(['r2', 'r3', 'r4', 'r5']);
    expect(b.sheets[0].cells[0][4]?.v).toBe('r1');
  });

  it('和原位置部分重叠的列带：拒绝', () => {
    const b = prep();
    let out: unknown;
    edit(b, '插入剪切的单元格', (tx) => (out = insertCut(tx, 0, { r0: 0, r1: 0, c0: 0, c1: 1 }, 3, 1, 'down')));
    expect(typeof out).toBe('string');
  });

  it('改写引用：只改完全落在区域里的', () => {
    const g = { r0: 3, r1: 4, c0: 0, c1: 1 };
    expect(retargetFormula('A4+$B$5+A4:B5+A3:B5+Other!A4+S!A4', 'S', 'S', g, -2, 0)).toBe('A2+$B$3+A2:B3+A3:B5+Other!A4+S!A2');
  });
});

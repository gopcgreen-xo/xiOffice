import { describe, expect, it } from 'vitest';
import { removeDuplicates, splitText, splitWidth, textToColumns } from './data';
import { setInput, Tx } from './ops';
import { blankWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, 'x');
  fn(tx);
  return tx.commit();
};
const grid = (b: Workbook, r1: number, c1: number) =>
  Array.from({ length: r1 + 1 }, (_, r) => Array.from({ length: c1 + 1 }, (_, c) => b.sheets[0].cells[r]?.[c]?.v ?? null));

describe('删除重复项', () => {
  it('按选中的列比较（不分大小写），留第一次出现的行，选区外不动', () => {
    const b = blankWorkbook();
    const rows = [['名称', '数量', '外'], ['钢板', '1', 'a'], ['角钢', '2', 'b'], ['钢板', '3', 'c'], ['ABC', '4', 'd'], ['abc', '5', 'e']];
    edit(b, (tx) => rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, v))));
    let res: [number, number] | null = null;
    edit(b, (tx) => (res = removeDuplicates(tx, 0, { r0: 0, c0: 0, r1: 5, c1: 1 }, [0], true)));
    expect(res).toEqual([2, 3]);
    expect(grid(b, 5, 2)).toEqual([
      ['名称', '数量', '外'],
      ['钢板', 1, 'a'],
      ['角钢', 2, 'b'],
      ['ABC', 4, 'c'],
      [null, null, 'd'],
      [null, null, 'e'],
    ]);
  });

  it('按两列比较时两列都相同才算重复', () => {
    const b = blankWorkbook();
    const rows = [['甲', '1'], ['甲', '2'], ['甲', '1']];
    edit(b, (tx) => rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, v))));
    let res: [number, number] | null = null;
    edit(b, (tx) => (res = removeDuplicates(tx, 0, { r0: 0, c0: 0, r1: 2, c1: 1 }, [0, 1], false)));
    expect(res).toEqual([1, 2]);
  });
});

describe('分列', () => {
  it('分隔符、连续分隔符合并、引号里的不拆', () => {
    expect(splitText('a,b,,c', { delims: [','], merge: false })).toEqual(['a', 'b', '', 'c']);
    expect(splitText('a  b', { delims: [' '], merge: true })).toEqual(['a', 'b']);
    expect(splitText('"x,y",z', { delims: [','], merge: false })).toEqual(['x,y', 'z']);
  });

  it('拆到右边几列，每段按输入识别；前导 0 的编号仍是文字', () => {
    const b = blankWorkbook();
    edit(b, (tx) => {
      setInput(tx, 0, 0, 0, 'Q235/12/0012');
      setInput(tx, 0, 1, 0, '不拆');
    });
    const o = { delims: ['/'], merge: false };
    expect(splitWidth(b.sheets[0], { r0: 0, c0: 0, r1: 1, c1: 0 }, o)).toBe(3);
    edit(b, (tx) => textToColumns(tx, 0, { r0: 0, c0: 0, r1: 1, c1: 0 }, o));
    expect(grid(b, 1, 2)).toEqual([
      ['Q235', 12, '0012'],
      ['不拆', null, null],
    ]);
  });
});

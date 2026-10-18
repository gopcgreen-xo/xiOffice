import { describe, expect, it } from 'vitest';
import { copyBlock, pasteBlock, pasteOpValue, setInput, Tx } from './ops';
import { blankWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, '测试');
  fn(tx);
  return tx.commit();
};

describe('选择性粘贴', () => {
  it('运算规则：数字才运算，空目标按 0，除以 0 报错，文字不动', () => {
    expect(pasteOpValue('add', 10, 5)).toBe(15);
    expect(pasteOpValue('sub', 10, 5)).toBe(5);
    expect(pasteOpValue('mul', 10, 5)).toBe(50);
    expect(pasteOpValue('div', 10, 4)).toBe(2.5);
    expect(pasteOpValue('div', 10, 0)).toBe('#DIV/0!');
    expect(pasteOpValue('add', '', 5)).toBe(5);
    expect(pasteOpValue('add', 'abc', 5)).toBeUndefined();
    expect(pasteOpValue('add', 1, 'x')).toBeUndefined();
  });

  it('运算粘贴：只改数值，保留目标格式；公式和文字格子不变', () => {
    const b = blankWorkbook();
    edit(b, (tx) => {
      setInput(tx, 0, 0, 0, '2');
      setInput(tx, 0, 0, 1, '3');
      setInput(tx, 0, 0, 2, '4');
      setInput(tx, 0, 1, 0, '10');
      setInput(tx, 0, 1, 1, '文字');
      setInput(tx, 0, 1, 2, '=A2+1');
    });
    const block = copyBlock(b, 0, { r0: 0, c0: 0, r1: 0, c1: 2 });
    edit(b, (tx) => pasteBlock(tx, 0, 1, 0, block, 'all', false, { op: 'mul' }));
    const row = b.sheets[0].cells[1];
    expect(row[0]?.v).toBe(20);
    expect(row[0]?.t).toBe('20');
    expect(row[1]?.v).toBe('文字');
    expect(row[2]?.f).toBe('A2+1');
  });

  it('跳过空单元格：来源为空的格子不覆盖目标', () => {
    const b = blankWorkbook();
    edit(b, (tx) => {
      setInput(tx, 0, 0, 0, '1');
      setInput(tx, 0, 1, 0, '7');
      setInput(tx, 0, 1, 1, '8');
    });
    const block = copyBlock(b, 0, { r0: 0, c0: 0, r1: 0, c1: 1 });
    edit(b, (tx) => pasteBlock(tx, 0, 1, 0, block, 'all', false, { skipBlanks: true }));
    expect(b.sheets[0].cells[1][0]?.v).toBe(1);
    expect(b.sheets[0].cells[1][1]?.v).toBe(8);
    edit(b, (tx) => pasteBlock(tx, 0, 1, 0, block));
    expect(b.sheets[0].cells[1][1]).toBeUndefined();
  });

  it('只粘贴公式：带公式平移，不带格式', () => {
    const b = blankWorkbook();
    edit(b, (tx) => {
      setInput(tx, 0, 0, 0, '5');
      setInput(tx, 0, 0, 1, '=A1*2');
    });
    const block = copyBlock(b, 0, { r0: 0, c0: 0, r1: 0, c1: 1 });
    edit(b, (tx) => pasteBlock(tx, 0, 3, 0, block, 'formulas'));
    expect(b.sheets[0].cells[3][1]?.f).toBe('A4*2');
  });
});

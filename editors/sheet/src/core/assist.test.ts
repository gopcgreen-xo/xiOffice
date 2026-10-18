import { describe, expect, it } from 'vitest';
import { autoSumRange, callAt, functionNames, suggest, typingName } from './assist';
import { setInput, Tx } from './ops';
import { blankWorkbook } from './workbook';

describe('写公式的帮助', () => {
  it('正在输入的函数名：只在公式里、不在字符串里', () => {
    expect(typingName('=su', 3)).toEqual({ start: 1, prefix: 'SU' });
    expect(typingName('=A1+vl', 6)).toEqual({ start: 4, prefix: 'VL' });
    expect(typingName('="su', 4)).toBeNull();
    expect(typingName('su', 2)).toBeNull();
    expect(typingName('=A1', 3)).toEqual({ start: 1, prefix: 'A1' });
  });

  it('候选：开头相同的在前，只列已支持的函数', () => {
    expect(suggest('SUM').slice(0, 3)).toEqual(['SUM', 'SUMIF', 'SUMIFS']);
    expect(suggest('LOOKUP')).toEqual(['LOOKUP', 'HLOOKUP', 'VLOOKUP', 'XLOOKUP']);
    expect(functionNames()).toContain('VLOOKUP');
  });

  it('光标所在的函数和参数序号（嵌套、字符串里的逗号不算）', () => {
    const f = '=IF(A1>0,VLOOKUP(B1,"a,b",2),0)';
    expect(callAt(f, f.indexOf('"a') + 2)).toEqual({ name: 'VLOOKUP', arg: 1 });
    expect(callAt(f, f.indexOf(',2') + 2)).toEqual({ name: 'VLOOKUP', arg: 2 });
    expect(callAt(f, f.length - 1)).toEqual({ name: 'IF', arg: 2 });
    expect(callAt('=A1+1', 3)).toBeNull();
  });

  it('自动求和：上方连续的数字，遇到文字标题停；上方不是数字时看左边', () => {
    const b = blankWorkbook();
    const tx = new Tx(b, 'x');
    ['数量', '1', '2', '3'].forEach((v, r) => setInput(tx, 0, r, 0, v));
    ['5', '6'].forEach((v, c) => setInput(tx, 0, 5, c, v));
    tx.commit();
    const s = b.sheets[0];
    expect(autoSumRange(s, 4, 0)).toEqual({ r0: 1, c0: 0, r1: 3, c1: 0 });
    expect(autoSumRange(s, 5, 2)).toEqual({ r0: 5, c0: 0, r1: 5, c1: 1 });
    expect(autoSumRange(s, 9, 9)).toBeNull();
  });
});

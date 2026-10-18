import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import { join } from 'path';
import { evalCf } from './cf';
import { Engine } from './formula/eval';
import { cellKey } from './grid';
import { setInput, Tx } from './ops';
import { blankWorkbook, parseSqref, readWorkbook, type CfRule } from './workbook';

function book(values: (string | number)[][], rules: Omit<CfRule, 'priority' | 'formulas'>[] & Partial<CfRule>[]) {
  const b = blankWorkbook();
  const tx = new Tx(b, 'x');
  values.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, String(v))));
  tx.commit();
  b.styleBook!.dxfs = [{ fill: '#ffc7ce', color: '#9c0006' }, { bold: true }];
  b.sheets[0].cf = rules.map((x, i) => ({ priority: i + 1, formulas: [], ...x }) as CfRule);
  return b;
}

describe('条件格式', () => {
  it('单元格值大于、介于；公式规则相对引用按行变化；按优先级合并', () => {
    const b = book([[5], [15], [25]], [
      { ranges: parseSqref('A1:A3'), type: 'cellIs', operator: 'greaterThan', formulas: ['10'], dxf: 0 },
      { ranges: parseSqref('A1:A3'), type: 'expression', formulas: ['$A1>20'], dxf: 1 },
      { ranges: parseSqref('A1:A3'), type: 'cellIs', operator: 'between', formulas: ['1', '6'], dxf: 1 },
    ]);
    const m = evalCf(b, 0, new Engine(b));
    expect(m.get(cellKey(0, 0))?.style).toEqual({ bold: true });
    expect(m.get(cellKey(1, 0))?.style).toEqual({ fill: '#ffc7ce', color: '#9c0006' });
    expect(m.get(cellKey(2, 0))?.style).toEqual({ fill: '#ffc7ce', color: '#9c0006', bold: true });
  });

  it('重复值、包含文字、前 N 项、高于平均', () => {
    const b = book([['钢板'], ['螺栓'], ['钢板'], ['角钢']], [{ ranges: parseSqref('A1:A4'), type: 'duplicateValues', dxf: 0 }]);
    const m = evalCf(b, 0, null);
    expect([0, 1, 2, 3].map((r) => !!m.get(cellKey(r, 0)))).toEqual([true, false, true, false]);
    const t = book([['钢板 Q235'], ['螺栓']], [{ ranges: parseSqref('A1:A2'), type: 'containsText', text: 'q235', dxf: 1 }]);
    expect(!!evalCf(t, 0, null).get(cellKey(0, 0))).toBe(true);
    const n = book([[1], [9], [5], [7]], [
      { ranges: parseSqref('A1:A4'), type: 'top10', rank: 2, dxf: 1 },
      { ranges: parseSqref('A1:A4'), type: 'aboveAverage', above: false, dxf: 0 },
    ]);
    const nm = evalCf(n, 0, null);
    expect([0, 1, 2, 3].map((r) => JSON.stringify(nm.get(cellKey(r, 0))?.style ?? {}))).toEqual([
      JSON.stringify({ fill: '#ffc7ce', color: '#9c0006' }),
      JSON.stringify({ bold: true }),
      JSON.stringify({ fill: '#ffc7ce', color: '#9c0006' }),
      JSON.stringify({ bold: true }),
    ]);
  });

  it('色阶两色按比例混合；数据条长度按最小到最大', () => {
    const b = book([[0], [50], [100]], [
      { ranges: parseSqref('A1:A3'), type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['#ffffff', '#000000'] },
      { ranges: parseSqref('A1:A3'), type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['#638ec6'] },
    ]);
    const m = evalCf(b, 0, null);
    expect([0, 1, 2].map((r) => m.get(cellKey(r, 0))?.style?.fill)).toEqual(['#ffffff', '#808080', '#000000']);
    expect([0, 1, 2].map((r) => m.get(cellKey(r, 0))?.bar?.pct)).toEqual([0.1, 0.55, 1]);
  });

  it('读取文件里的条件格式和差异格式（回归样本）', () => {
    const data = new Uint8Array(readFileSync(join(__dirname, '../../samples/报价单样例.xlsx')));
    const b = readWorkbook(data);
    const m = evalCf(b, 0, new Engine(b));
    // D3 = 14130.65 > 10000：红字粉底；D4 = 192.1 不满足
    expect(m.get(cellKey(2, 3))?.style).toMatchObject({ fill: '#ffc7ce', color: '#c00000' });
    expect(m.get(cellKey(3, 3))).toBeUndefined();
    expect(b.sheets[0].validations?.[0]).toMatchObject({ type: 'list', formula1: '"TRUE,FALSE"', showError: true });
  });
});

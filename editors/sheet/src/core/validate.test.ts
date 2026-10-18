import { describe, expect, it } from 'vitest';
import { checkValue, listItems, validationAt } from './validate';
import { Engine } from './formula/eval';
import { setInput, Tx } from './ops';
import { blankWorkbook, parseSqref, type Validation } from './workbook';

function sample() {
  const b = blankWorkbook();
  const tx = new Tx(b, 'x');
  ['已报价', '已到货', '已付款', '已到货'].forEach((v, r) => setInput(tx, 0, r, 5, v));
  setInput(tx, 0, 0, 6, '100');
  tx.commit();
  b.names = [{ name: '状态', f: 'Sheet1!$F$1:$F$4' }];
  return b;
}
const v = (x: Partial<Validation>): Validation => ({ ranges: parseSqref('A1:A10'), type: 'list', allowBlank: true, showError: true, ...x });

describe('数据验证', () => {
  it('序列：写死的、引用区域的、名称；去重', () => {
    const b = sample();
    expect(listItems(v({ formula1: '"是,否, 待定"' }), b, 0)).toEqual(['是', '否', '待定']);
    expect(listItems(v({ formula1: '$F$1:$F$4' }), b, 0)).toEqual(['已报价', '已到货', '已付款']);
    expect(listItems(v({ formula1: '状态' }), b, 0)).toEqual(['已报价', '已到货', '已付款']);
  });

  it('检查：序列、整数介于、小数大于引用、文字长度、空值', () => {
    const b = sample();
    const e = new Engine(b);
    expect(checkValue(v({ formula1: '"是,否"' }), '是', b, 0, 0, 0, e)).toBeNull();
    expect(checkValue(v({ formula1: '"是,否"' }), '也许', b, 0, 0, 0, e)).toMatch(/下拉列表/);
    expect(checkValue(v({ type: 'whole', operator: 'between', formula1: '1', formula2: '10' }), 5, b, 0, 0, 0, e)).toBeNull();
    expect(checkValue(v({ type: 'whole', operator: 'between', formula1: '1', formula2: '10' }), 5.5, b, 0, 0, 0, e)).toBe('请输入整数');
    expect(checkValue(v({ type: 'decimal', operator: 'greaterThan', formula1: '$G$1' }), 50, b, 0, 0, 0, e)).toBe('数值应大于 100');
    expect(checkValue(v({ type: 'textLength', operator: 'lessThanOrEqual', formula1: '4' }), '采购合同书', b, 0, 0, 0, e)).toBe('文字长度应小于或等于 4');
    expect(checkValue(v({ allowBlank: false, formula1: '"是"' }), null, b, 0, 0, 0, e)).toBe('这个格子不能为空');
  });

  it('按位置找验证', () => {
    const b = sample();
    b.sheets[0].validations = [v({ ranges: parseSqref('B2:B3 D5') })];
    expect(validationAt(b.sheets[0], 2, 1)).toBeTruthy();
    expect(validationAt(b.sheets[0], 4, 3)).toBeTruthy();
    expect(validationAt(b.sheets[0], 0, 0)).toBeUndefined();
  });
});

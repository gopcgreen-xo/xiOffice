import { describe, expect, it } from 'vitest';
import { Engine } from './formula/eval';
import { toFileFormula } from './formula/parse';
import { clearRange, setInput, Tx } from './ops';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';
import { saveXlsx } from './save';
import { strFromU8, unzipSync } from 'fflate';
import { colIndex } from '@wz/editor-shared/xlsx';

function put(b: Workbook, e: Engine, cells: Record<string, string | number>) {
  const tx = new Tx(b, '输入');
  for (const [a, v] of Object.entries(cells)) {
    const m = a.match(/^([A-Z]+)(\d+)$/)!;
    setInput(tx, 0, Number(m[2]) - 1, colIndex(m[1]), String(v));
  }
  const entry = tx.commit();
  const formulas = !!entry?.cells.some((c) => c.before?.f || c.after?.f);
  e.recalc(tx.changed, { rebuild: formulas, all: tx.structural });
}
const val = (b: Workbook, a: string) => {
  const m = a.match(/^([A-Z]+)(\d+)$/)!;
  return b.sheets[0].cells[Number(m[2]) - 1]?.[colIndex(m[1])]?.v;
};
const cell = (b: Workbook, a: string) => {
  const m = a.match(/^([A-Z]+)(\d+)$/)!;
  return b.sheets[0].cells[Number(m[2]) - 1]?.[colIndex(m[1])];
};
function calc(f: string, cells: Record<string, string | number> = {}) {
  const b = blankWorkbook();
  const e = new Engine(b);
  put(b, e, cells);
  put(b, e, { Z99: `=${f}` });
  return val(b, 'Z99');
}
const close = (x: unknown, y: number, digits = 4) => expect(x as number).toBeCloseTo(y, digits);

describe('新增函数：数学 / 统计', () => {
  it('数学', () => {
    expect(calc('SQRT(16)')).toBe(4);
    expect(calc('SQRT(-1)')).toBe('#NUM!');
    expect(calc('POWER(2,10)')).toBe(1024);
    expect(calc('SIGN(-5)')).toBe(-1);
    expect(calc('LOG(8,2)')).toBe(3);
    expect(calc('LOG10(1000)')).toBe(3);
    expect(calc('TRUNC(-8.9)')).toBe(-8);
    expect(calc('TRUNC(3.14159,2)')).toBe(3.14);
    expect(calc('CEILING(2.5,1)')).toBe(3);
    expect(calc('FLOOR(2.5,2)')).toBe(2);
    expect(calc('CEILING.MATH(-2.5)')).toBe(-2);
    expect(calc('FLOOR.MATH(-2.5)')).toBe(-3);
    expect(calc('MROUND(10,3)')).toBe(9);
    expect(calc('GCD(24,36)')).toBe(12);
    expect(calc('LCM(4,6)')).toBe(12);
    expect(calc('FACT(5)')).toBe(120);
    expect(calc('SUMSQ(1,2,3)')).toBe(14);
    expect(calc('QUOTIENT(7,2)')).toBe(3);
    expect(calc('RAND()')).toBeGreaterThanOrEqual(0);
    const n = calc('RANDBETWEEN(3,5)') as number;
    expect(n >= 3 && n <= 5).toBe(true);
  });

  it('统计', () => {
    const d = { A1: 1, A2: 2, A3: 2, A4: 3, A5: 4 };
    expect(calc('MODE(A1:A5)', d)).toBe(2);
    expect(calc('PERCENTILE(A1:A5,0.5)', d)).toBe(2);
    expect(calc('QUARTILE(A1:A5,1)', d)).toBe(2);
    expect(calc('STDEV.S(A1:A5)', d)).toBeCloseTo(1.1402, 3);
    const xy = { A1: 1, A2: 2, A3: 3, B1: 2, B2: 4, B3: 6 };
    expect(calc('SLOPE(B1:B3,A1:A3)', xy)).toBe(2);
    expect(calc('INTERCEPT(B1:B3,A1:A3)', xy)).toBeCloseTo(0, 9);
    expect(calc('CORREL(A1:A3,B1:B3)', xy)).toBe(1);
    expect(calc('FORECAST(4,B1:B3,A1:A3)', xy)).toBeCloseTo(8, 9);
  });
});

describe('新增函数：文本 / 信息 / 逻辑', () => {
  it('文本', () => {
    expect(calc('REPT("ab",3)')).toBe('ababab');
    expect(calc('CHAR(65)')).toBe('A');
    expect(calc('CODE("A")')).toBe(65);
    expect(calc('PROPER("hello woRLD")')).toBe('Hello World');
    expect(calc('EXACT("a","A")')).toBe(false);
    expect(calc('REPLACE("abcdef",2,3,"X")')).toBe('aXef');
    expect(calc('TEXTBEFORE("a-b-c","-")')).toBe('a');
    expect(calc('TEXTAFTER("a-b-c","-",2)')).toBe('c');
    expect(calc('T("x")')).toBe('x');
    expect(calc('T(5)')).toBe('');
  });

  it('信息和逻辑', () => {
    expect(calc('ISEVEN(4)')).toBe(true);
    expect(calc('ISODD(4)')).toBe(false);
    expect(calc('ISLOGICAL(TRUE)')).toBe(true);
    expect(calc('ISERR(1/0)')).toBe(true);
    expect(calc('ISERR(NA())')).toBe(false);
    expect(calc('N("a")')).toBe(0);
    expect(calc('TYPE("a")')).toBe(2);
    expect(calc('XOR(TRUE,FALSE,TRUE)')).toBe(false);
    expect(calc('SWITCH(2,1,"一",2,"二","其它")')).toBe('二');
    expect(calc('SWITCH(9,1,"一",2,"二","其它")')).toBe('其它');
  });
});

describe('新增函数：查找 / 日期 / 财务', () => {
  it('查找', () => {
    const d = { A1: 10, A2: 20, A3: 30, B1: '甲', B2: '乙', B3: '丙' };
    expect(calc('LOOKUP(25,A1:A3,B1:B3)', d)).toBe('乙');
    expect(calc('XMATCH(30,A1:A3)', d)).toBe(3);
    expect(calc('XMATCH(25,A1:A3,-1)', d)).toBe(2);
    expect(calc('XMATCH(25,A1:A3,1)', d)).toBe(3);
    expect(calc('ROWS(A1:B3)')).toBe(3);
    expect(calc('COLUMNS(A1:B3)')).toBe(2);
    expect(calc('ADDRESS(2,3)')).toBe('$C$2');
    expect(calc('ADDRESS(2,3,4)')).toBe('C2');
    expect(calc('ADDRESS(2,3,1,FALSE)')).toBe('R2C3');
  });

  it('日期', () => {
    expect(calc('DAYS(DATE(2026,3,1),DATE(2026,2,1))')).toBe(28);
    expect(calc('WEEKNUM(DATE(2026,1,1))')).toBe(1);
    expect(calc('WEEKNUM(DATE(2026,1,4))')).toBe(2);
    expect(calc('ISOWEEKNUM(DATE(2026,1,1))')).toBe(1);
    expect(calc('ISOWEEKNUM(DATE(2021,1,3))')).toBe(53);
    expect(calc('HOUR(0.75)')).toBe(18);
    expect(calc('MINUTE(0.5104166667)')).toBe(15);
    expect(calc('TIME(12,0,0)')).toBe(0.5);
    expect(calc('YEARFRAC(DATE(2026,1,1),DATE(2026,7,1))')).toBe(0.5);
  });

  it('财务（对照 Excel 帮助里的例子）', () => {
    close(calc('PMT(0.08/12,10,10000)'), -1037.032089, 5);
    close(calc('FV(0.06/12,10,-200,-500,1)'), 2581.403374, 5);
    close(calc('PV(0.08/12,240,500)'), -59777.14585, 4);
    close(calc('NPER(0.12/12,-100,-1000,10000,1)'), 59.6738657, 5);
    close(calc('NPV(0.1,-10000,3000,4200,6800)'), 1188.4434, 3);
    close(calc('IRR(A1:A5)', { A1: -70000, A2: 12000, A3: 15000, A4: 18000, A5: 21000 }), -0.0212448, 5);
    close(calc('RATE(48,-200,8000)'), 0.0077, 4);
    close(calc('IPMT(0.1/12,1,36,8000)'), -66.6667, 3);
    close(calc('PPMT(0.1/12,1,36,8000)'), -191.4708, 3);
    expect(calc('SLN(30000,7500,10)')).toBe(2250);
  });
});

describe('动态数组', () => {
  const data = { A1: '苹果', A2: '香蕉', A3: '苹果', A4: '梨', B1: 3, B2: 5, B3: 2, B4: 8 };

  it('作为参数用：SUM(FILTER(...))、COUNTA(UNIQUE(...))', () => {
    expect(calc('SUM(FILTER(B1:B4,B1:B4>2))', data)).toBe(16);
    expect(calc('COUNTA(UNIQUE(A1:A4))', data)).toBe(3);
    expect(calc('SUM(SEQUENCE(4))')).toBe(10);
    expect(calc('INDEX(SORT(B1:B4,1,-1),1)', data)).toBe(8);
    expect(calc('COUNTA(TEXTSPLIT("a,b,c",","))')).toBe(3);
  });

  it('顶层结果溢出到相邻格子，写成数组公式区域', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    put(b, e, data);
    put(b, e, { D1: '=UNIQUE(A1:A4)' });
    expect([val(b, 'D1'), val(b, 'D2'), val(b, 'D3'), val(b, 'D4')]).toEqual(['苹果', '香蕉', '梨', undefined]);
    expect(cell(b, 'D1')?.fa).toBe('D1:D3');
    expect(cell(b, 'D2')?.sp).toBe(true);
    // 源数据变了：区域跟着变化
    put(b, e, { A3: '葡萄' });
    expect([val(b, 'D3'), val(b, 'D4')]).toEqual(['葡萄', '梨']);
    expect(cell(b, 'D1')?.fa).toBe('D1:D4');
    put(b, e, { A3: '苹果' });
    expect(val(b, 'D4')).toBeUndefined();
    expect(cell(b, 'D1')?.fa).toBe('D1:D3');
  });

  it('二维：FILTER 取多列、SORT 排序、SEQUENCE 铺满', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    put(b, e, data);
    put(b, e, { D1: '=SORT(A1:B4,2,-1)' });
    expect([val(b, 'D1'), val(b, 'E1'), val(b, 'D4'), val(b, 'E4')]).toEqual(['梨', 8, '苹果', 2]);
    put(b, e, { H1: '=FILTER(A1:B4,B1:B4>2)' });
    expect([val(b, 'H1'), val(b, 'I1'), val(b, 'H3'), val(b, 'I3')]).toEqual(['苹果', 3, '梨', 8]);
    put(b, e, { K1: '=SEQUENCE(2,3)' });
    expect([val(b, 'K1'), val(b, 'M1'), val(b, 'K2'), val(b, 'M2')]).toEqual([1, 3, 4, 6]);
  });

  it('引用溢出结果的公式会跟着算；挡路返回 #SPILL!，让开后恢复', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    put(b, e, data);
    put(b, e, { D1: '=UNIQUE(A1:A4)', F1: '=D3', F2: '=SUM(D1:D3)' });
    expect(val(b, 'F1')).toBe('梨');
    put(b, e, { D2: '挡路' });
    expect(val(b, 'D1')).toBe('#SPILL!');
    expect(cell(b, 'D1')?.fa).toBeUndefined();
    expect(val(b, 'D2')).toBe('挡路');
    const tx = new Tx(b, '清除');
    clearRange(tx, 0, { r0: 1, c0: 3, r1: 1, c1: 3 });
    tx.commit();
    e.recalc(tx.changed, {});
    expect(val(b, 'D1')).toBe('苹果');
    expect(val(b, 'D2')).toBe('香蕉');
    expect(val(b, 'F1')).toBe('梨');
  });

  it('删掉公式后溢出的格子一起清除', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    put(b, e, data);
    put(b, e, { D1: '=UNIQUE(A1:A4)' });
    put(b, e, { D1: '空' });
    expect(val(b, 'D2')).toBeUndefined();
    expect(val(b, 'D3')).toBeUndefined();
  });

  it('文件里的前缀：FILTER / SORT 是 _xlfn._xlws.，其它是 _xlfn.', () => {
    expect(toFileFormula('FILTER(A1:A3,B1:B3>1)')).toBe('_xlfn._xlws.FILTER(A1:A3,B1:B3>1)');
    expect(toFileFormula('SORT(A1:A3)+UNIQUE(B1:B3)')).toBe('_xlfn._xlws.SORT(A1:A3)+_xlfn.UNIQUE(B1:B3)');
    expect(toFileFormula('TEXTSPLIT(A1,",")')).toBe('_xlfn.TEXTSPLIT(A1,",")');
  });

  it('保存成数组公式（带缓存值），重新打开后仍按动态数组计算', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    put(b, e, data);
    put(b, e, { D1: '=UNIQUE(A1:A4)' });
    const out = saveXlsx(b);
    const xml = strFromU8(unzipSync(out)['xl/worksheets/sheet1.xml']);
    expect(xml).toContain('<f t="array" ref="D1:D3">_xlfn.UNIQUE(A1:A4)</f>');
    const b2 = readWorkbook(out);
    const e2 = new Engine(b2);
    e2.analyze();
    expect(cell(b2, 'D1')?.dyn).toBe(true);
    expect(cell(b2, 'D3')?.sp).toBe(true);
    put(b2, e2, { A4: '西瓜' });
    expect(val(b2, 'D3')).toBe('西瓜');
  });
});

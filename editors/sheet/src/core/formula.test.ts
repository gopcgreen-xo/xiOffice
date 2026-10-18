import { describe, expect, it } from 'vitest';
import { Engine } from './formula/eval';
import { fromFileFormula, normalizeFormula, parseFormula, toFileFormula } from './formula/parse';
import { applyEntry, insertDelete, setInput, Tx } from './ops';
import { blankWorkbook, type Workbook } from './workbook';
import { colIndex } from '@wz/editor-shared/xlsx';

/** 按 A1 地址写一批内容，然后重算（同编辑器：提交事务后 recalc） */
function put(b: Workbook, e: Engine, cells: Record<string, string | number>, si = 0) {
  const tx = new Tx(b, '输入');
  for (const [a, v] of Object.entries(cells)) {
    const m = a.match(/^([A-Z]+)(\d+)$/)!;
    setInput(tx, si, Number(m[2]) - 1, colIndex(m[1]), String(v));
  }
  const entry = tx.commit();
  const formulas = !!entry?.cells.some((c) => c.before?.f || c.after?.f);
  e.recalc(tx.changed, { rebuild: formulas, all: tx.structural });
  return entry;
}

const val = (b: Workbook, a: string, si = 0) => {
  const m = a.match(/^([A-Z]+)(\d+)$/)!;
  return b.sheets[si].cells[Number(m[2]) - 1]?.[colIndex(m[1])]?.v;
};
const txt = (b: Workbook, a: string) => {
  const m = a.match(/^([A-Z]+)(\d+)$/)!;
  return b.sheets[0].cells[Number(m[2]) - 1]?.[colIndex(m[1])]?.t;
};

function calc(f: string, cells: Record<string, string | number> = {}) {
  const b = blankWorkbook();
  const e = new Engine(b);
  put(b, e, cells);
  put(b, e, { Z99: `=${f}` });
  return val(b, 'Z99');
}

describe('公式解析', () => {
  it('运算符优先级同 Excel', () => {
    expect(calc('1+2*3')).toBe(7);
    expect(calc('-2^2')).toBe(4);
    expect(calc('2^3^2')).toBe(64);
    expect(calc('10%*5')).toBe(0.5);
    expect(calc('1+2&"x"')).toBe('3x');
    expect(calc('1+1=2')).toBe(true);
    expect(calc('0.1+0.2')).toBe(0.3);
  });

  it('语法错误抛出，_xlfn 前缀读写，用户输入规范成大写', () => {
    expect(() => parseFormula('SUM(1,')).toThrow();
    expect(fromFileFormula('_xlfn.XLOOKUP(A1,B:B,C:C)&"_xlfn.x"')).toBe('XLOOKUP(A1,B:B,C:C)&"_xlfn.x"');
    expect(toFileFormula('IFS(A1>1,1,TRUE,2)+SUM(A1)')).toBe('_xlfn.IFS(A1>1,1,TRUE,2)+SUM(A1)');
    expect(normalizeFormula('sum(a1:b2)+\'x y\'!c3')).toBe("SUM(A1:B2)+'x y'!C3");
  });
});

describe('函数', () => {
  const data = { A1: '1', A2: '2', A3: '3', A4: '苹果', B1: '10', B2: '20', B3: '30', B4: '40' };
  it('统计', () => {
    expect(calc('SUM(A1:A4)', data)).toBe(6);
    expect(calc('SUM(A1:A3,"4",TRUE)', data)).toBe(11);
    expect(calc('AVERAGE(A1:A4)', data)).toBe(2);
    expect(calc('COUNT(A1:A4)', data)).toBe(3);
    expect(calc('COUNTA(A1:A5)', data)).toBe(4);
    expect(calc('COUNTBLANK(A1:A5)', data)).toBe(1);
    expect(calc('MAX(B1:B4)', data)).toBe(40);
    expect(calc('MIN(B1:B4)', data)).toBe(10);
    expect(calc('SUMPRODUCT(A1:A3,B1:B3)', data)).toBe(140);
    expect(calc('SUMPRODUCT((A1:A3>1)*B1:B3)', data)).toBe(50);
  });

  it('条件统计', () => {
    expect(calc('COUNTIF(A1:A4,">1")', data)).toBe(2);
    expect(calc('COUNTIF(A1:A4,"苹*")', data)).toBe(1);
    expect(calc('SUMIF(A1:A3,">=2",B1:B3)', data)).toBe(50);
    expect(calc('SUMIFS(B1:B4,A1:A4,"<>苹果",B1:B4,">10")', data)).toBe(50);
    expect(calc('COUNTIFS(A1:A4,">0",B1:B4,"<35")', data)).toBe(3);
    expect(calc('AVERAGEIF(A1:A3,"<3",B1:B3)', data)).toBe(15);
  });

  it('数学', () => {
    expect(calc('ROUND(2.5,0)')).toBe(3);
    expect(calc('ROUND(-2.5,0)')).toBe(-3);
    expect(calc('ROUND(1.005,2)')).toBe(1.01);
    expect(calc('ROUND(1234,-2)')).toBe(1200);
    expect(calc('ROUNDUP(1.21,1)')).toBe(1.3);
    expect(calc('ROUNDDOWN(-1.29,1)')).toBe(-1.2);
    expect(calc('INT(-1.5)')).toBe(-2);
    expect(calc('MOD(-3,2)')).toBe(1);
    expect(calc('1/0')).toBe('#DIV/0!');
  });

  it('逻辑', () => {
    expect(calc('IF(A1>0,"正","负")', data)).toBe('正');
    expect(calc('IF(FALSE,1)')).toBe(false);
    expect(calc('IFERROR(1/0,"除零")')).toBe('除零');
    expect(calc('IFS(A3=1,"一",A3=3,"三")', data)).toBe('三');
    expect(calc('AND(A1:A3)', data)).toBe(true);
    expect(calc('OR(A1>5,B1>5)', data)).toBe(true);
    expect(calc('NOT(1)')).toBe(false);
    // IF 只算选中的分支
    expect(calc('IF(TRUE,1,1/0)')).toBe(1);
  });

  it('查找', () => {
    const t = { A1: '编号', B1: '名称', A2: 'H001', B2: '钢板', A3: 'H002', B3: '螺栓', D1: '10', D2: '20', D3: '30', E1: '甲', E2: '乙', E3: '丙' };
    expect(calc('VLOOKUP("h002",A1:B3,2,FALSE)', t)).toBe('螺栓');
    expect(calc('VLOOKUP("H00*",A2:B3,2,0)', t)).toBe('钢板');
    expect(calc('VLOOKUP("X",A1:B3,2,FALSE)', t)).toBe('#N/A');
    expect(calc('VLOOKUP(25,D1:E3,2)', t)).toBe('乙');
    expect(calc('HLOOKUP("名称",A1:B3,3,FALSE)', t)).toBe('螺栓');
    expect(calc('INDEX(A1:B3,3,2)', t)).toBe('螺栓');
    expect(calc('MATCH("H002",A1:A3,0)', t)).toBe(3);
    expect(calc('INDEX(B1:B3,MATCH("H001",A1:A3,0))', t)).toBe('钢板');
    expect(calc('XLOOKUP("H002",A2:A3,B2:B3)', t)).toBe('螺栓');
    expect(calc('XLOOKUP("H9",A2:A3,B2:B3,"无")', t)).toBe('无');
  });

  it('文本', () => {
    expect(calc('LEN("采购合同")')).toBe(4);
    expect(calc('LEFT("HT2026",2)&RIGHT("HT2026",4)&MID("ABCDE",2,3)')).toBe('HT2026BCD');
    expect(calc('TRIM("  a   b ")')).toBe('a b');
    expect(calc('CONCATENATE("a",1,TRUE)')).toBe('a1TRUE');
    expect(calc('CONCAT(A1:A3)', data)).toBe('123');
    expect(calc('SUBSTITUTE("a-b-c","-","+",2)')).toBe('a-b+c');
    expect(calc('FIND("b","abcb",3)')).toBe(4);
    expect(calc('TEXT(0.256,"0.0%")')).toBe('25.6%');
    expect(calc('TEXT(46294,"yyyy年m月d日")')).toBe('2026年9月29日');
    expect(calc('VALUE("1,234")')).toBe(1234);
    expect(calc('"12"+1')).toBe(13);
    expect(calc('"a"+1')).toBe('#VALUE!');
  });

  it('日期', () => {
    expect(calc('DATE(2026,9,29)')).toBe(46294);
    expect(calc('DATE(2026,13,1)')).toBe(46388);
    expect(calc('YEAR(46294)&"-"&MONTH(46294)&"-"&DAY(46294)')).toBe('2026-9-29');
    expect(calc('DATEDIF(DATE(2025,1,31),DATE(2026,3,1),"M")')).toBe(13);
    expect(calc('DATEDIF(DATE(2025,1,31),DATE(2026,3,1),"Y")')).toBe(1);
    expect(calc('DATEDIF(DATE(2026,1,1),DATE(2026,3,1),"D")')).toBe(59);
  });
});

describe('重算', () => {
  it('改值后依赖它的公式（含间接、跨表）重算；撤销后也重算', () => {
    const b = blankWorkbook();
    b.sheets.push({ ...structuredClone(b.sheets[0]), name: '汇总' });
    const e = new Engine(b);
    put(b, e, { A1: '2', A2: '=A1*10', A3: '=A2+1' });
    put(b, e, { A1: '=Sheet1!A3*2' }, 1);
    expect([val(b, 'A2'), val(b, 'A3'), val(b, 'A1', 1)]).toEqual([20, 21, 42]);
    const entry = put(b, e, { A1: '5' })!;
    expect([val(b, 'A3'), val(b, 'A1', 1)]).toEqual([51, 102]);
    applyEntry(b, entry, -1);
    e.recalc(entry.cells, { rebuild: true });
    expect([val(b, 'A3'), val(b, 'A1', 1)]).toEqual([21, 42]);
  });

  it('长链按顺序算，不递归；循环引用显示 #CIRC!', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    const cells: Record<string, string> = { A1: '1' };
    for (let r = 2; r <= 3000; r++) cells[`A${r}`] = `=A${r - 1}+1`;
    put(b, e, cells);
    expect(val(b, 'A3000')).toBe(3000);
    put(b, e, { B1: '=B2', B2: '=B1' });
    expect(val(b, 'B1')).toBe('#CIRC!');
    expect(e.circular).toBe(2);
  });

  it('不支持的函数保留原值并记下；插入行后全部重算', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    put(b, e, { A1: '1', A2: '2', A3: '=SUM(A1:A2)', B1: '=LET(x,1,x+1)' });
    expect(e.isUnsupported(0, 0, 1)).toBe(true);
    const tx = new Tx(b, '插入');
    insertDelete(tx, 0, 'row', 1, 1);
    tx.commit();
    e.recalc(tx.changed, { all: true });
    expect(b.sheets[0].cells[3][0]?.f).toBe('SUM(A1:A3)');
    expect(val(b, 'A4')).toBe(3);
  });

  it('TODAY / DATE 公式在常规格子里自动用日期格式', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    put(b, e, { A1: '=date(2026,9,29)' });
    expect(b.sheets[0].cells[0][0]?.f).toBe('DATE(2026,9,29)');
    expect(txt(b, 'A1')).toBe('2026/9/29');
  });

  it('定义名称', () => {
    const b = blankWorkbook();
    b.names = [{ name: '税率', f: 'Sheet1!$B$1' }];
    const e = new Engine(b);
    put(b, e, { B1: '0.13', A1: '=100*税率' });
    expect(val(b, 'A1')).toBe(13);
    put(b, e, { B1: '0.09' });
    expect(val(b, 'A1')).toBe(9);
  });
});

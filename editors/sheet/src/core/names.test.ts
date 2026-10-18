import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { deleteName, nameProblem, nameTarget, refersToProblem, renameNameInFormula, saveName } from './names';
import { applyEntry, setInput, Tx } from './ops';
import { Engine } from './formula/eval';
import { saveXlsx } from './save';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => unknown) => {
  const tx = new Tx(b, '测试');
  fn(tx);
  return tx.commit();
};

describe('名称校验', () => {
  const b = blankWorkbook();
  it('合规与不合规', () => {
    expect(nameProblem(b, '销售额', undefined)).toBeNull();
    expect(nameProblem(b, 'Tax_Rate.2026', undefined)).toBeNull();
    expect(nameProblem(b, '', undefined)).toMatch(/不能为空/);
    expect(nameProblem(b, '1abc', undefined)).toMatch(/开头/);
    expect(nameProblem(b, 'a b', undefined)).toMatch(/空格/);
    for (const bad of ['A1', 'XFD100', 'R1C1', 'R', 'c', 'r2c3']) expect(nameProblem(b, bad, undefined)).toMatch(/单元格地址/);
    expect(nameProblem(b, '_xlnm.Print_Area', undefined)).toMatch(/_xlnm/);
  });
  it('同范围内不分大小写重名，不同范围可以同名', () => {
    edit(b, (tx) => saveName(tx, { name: 'Rate', f: 'Sheet1!$A$1' }));
    expect(nameProblem(b, 'rate', undefined)).toMatch(/已经存在/);
    expect(nameProblem(b, 'rate', 0)).toBeNull();
    expect(nameProblem(b, 'rate', undefined, b.names![0])).toBeNull();
  });
  it('引用位置：语法要对', () => {
    expect(refersToProblem('Sheet1!$A$1:$B$2')).toBeNull();
    expect(refersToProblem('0.13')).toBeNull();
    expect(refersToProblem('SUM(Sheet1!A1:A3)*2')).toBeNull();
    expect(refersToProblem('')).toMatch(/不能为空/);
    expect(refersToProblem('SUM(')).toMatch(/不对/);
  });
});

describe('名称增删改', () => {
  it('新建后公式能用；改引用位置公式跟着重算；撤销还原', () => {
    const b = blankWorkbook();
    const e = new Engine(b);
    const put = (a: Record<string, string>) => {
      const tx = new Tx(b, '输入');
      for (const [k, v] of Object.entries(a)) setInput(tx, 0, Number(k.slice(1)) - 1, k.charCodeAt(0) - 65, v);
      const en = tx.commit();
      e.recalc(tx.changed, { rebuild: !!en?.cells.some((c) => c.after?.f), all: tx.structural });
    };
    put({ A1: '10', A2: '20', B1: '=Total*2' });
    const tx = new Tx(b, '定义名称');
    expect(saveName(tx, { name: 'Total', f: '=SUM(Sheet1!$A$1:$A$2)' })).toBeNull();
    const en = tx.commit()!;
    e.recalc(tx.changed, { all: tx.structural });
    expect(b.sheets[0].cells[0][1]?.v).toBe(60);
    expect(b.names).toEqual([{ name: 'Total', f: 'SUM(Sheet1!$A$1:$A$2)' }]);
    expect(b.namesDirty).toBe(true);
    applyEntry(b, en, -1);
    expect(b.names ?? []).toEqual([]);
  });

  it('改名：公式和其它名称里的引用一起改，字符串和函数名不动', () => {
    expect(renameNameInFormula('Total*2+"Total"+Totals+SUM(Total)+Sheet1!Total+Total(1)', 'total', 'Sum2')).toBe('Sum2*2+"Total"+Totals+SUM(Sum2)+Sheet1!Total+Total(1)');
    const b = blankWorkbook();
    edit(b, (tx) => saveName(tx, { name: 'Rate', f: '0.13' }));
    edit(b, (tx) => setInput(tx, 0, 0, 0, '=100*Rate'));
    edit(b, (tx) => saveName(tx, { name: 'Tax', f: 'Rate*2' }));
    const old = b.names!.find((n) => n.name === 'Rate')!;
    edit(b, (tx) => saveName(tx, { name: '税率', f: '0.13' }, old));
    expect(b.sheets[0].cells[0][0]?.f).toBe('100*税率');
    expect(b.names!.map((n) => [n.name, n.f])).toEqual([['税率', '0.13'], ['Tax', '税率*2']]);
  });

  it('删除；重名 / 不合规返回原因且不改动', () => {
    const b = blankWorkbook();
    edit(b, (tx) => saveName(tx, { name: 'A_b', f: '1' }));
    const tx = new Tx(b, 'x');
    expect(saveName(tx, { name: 'a_B', f: '2' })).toMatch(/已经存在/);
    expect(saveName(tx, { name: 'ok', f: '1+' })).toMatch(/不对/);
    expect(b.names!.length).toBe(1);
    edit(b, (t) => deleteName(t, b.names![0]));
    expect(b.names).toEqual([]);
  });

  it('名称指向区域时能跳转（含带引号的表名、工作表范围优先）', () => {
    const b = blankWorkbook();
    edit(b, (tx) => saveName(tx, { name: 'Data', f: 'Sheet1!$B$2:$C$5' }));
    edit(b, (tx) => saveName(tx, { name: 'Pi', f: '3.14' }));
    expect(nameTarget(b, 'data', 0)).toEqual({ si: 0, range: { r0: 1, c0: 1, r1: 4, c1: 2 } });
    expect(nameTarget(b, 'Pi', 0)).toBeNull();
    expect(nameTarget(b, 'nope', 0)).toBeNull();
  });
});

describe('名称的读写', () => {
  it('保存 definedNames，重新打开一致；删除后文件里没有', () => {
    const b = blankWorkbook();
    edit(b, (tx) => saveName(tx, { name: 'Rate', f: '0.13' }));
    edit(b, (tx) => saveName(tx, { name: 'Local', f: 'Sheet1!$A$1', local: 0 }));
    const out = saveXlsx(b);
    const xml = strFromU8(unzipSync(out)['xl/workbook.xml']);
    expect(xml).toContain('<definedName name="Rate">0.13</definedName>');
    expect(xml).toContain('<definedName name="Local" localSheetId="0">Sheet1!$A$1</definedName>');
    const back = readWorkbook(out);
    expect(back.names).toEqual([{ name: 'Rate', f: '0.13' }, { name: 'Local', f: 'Sheet1!$A$1', local: 0 }]);
    back.source = out;
    edit(back, (tx) => deleteName(tx, back.names![0]));
    const xml2 = strFromU8(unzipSync(saveXlsx(back))['xl/workbook.xml']);
    expect(xml2).not.toContain('name="Rate"');
    expect(xml2).toContain('name="Local"');
  });

  it('只改了名称也要保存（没有工作表被改动）', () => {
    const b = blankWorkbook();
    const out0 = saveXlsx(b);
    const b2 = readWorkbook(out0);
    b2.source = out0;
    edit(b2, (tx) => saveName(tx, { name: 'X1y', f: '1' }));
    expect(strFromU8(unzipSync(saveXlsx(b2))['xl/workbook.xml'])).toContain('name="X1y"');
  });
});

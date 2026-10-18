import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { buildPivot, createPivot, defaultConfig, fieldNames, refreshPivot, type PivotConfig } from './pivot';
import { pivotXml } from './pivotParts';
import { applyEntry, setInput, Tx } from './ops';
import { markSaved, saveXlsx } from './save';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';
import { colIndex } from '@wz/editor-shared/xlsx';

const DATA = [
  ['区域', '产品', '数量', '金额'],
  ['华东', '甲', 10, 100],
  ['华东', '乙', 5, 80],
  ['华南', '甲', 7, 70],
  ['华东', '甲', 3, 30],
  ['华南', '乙', 2, 25],
  ['', '乙', 1, 10],
];
const T = { rows: true, cols: true };
const cfg = (rows: number[], cols: number[], values: PivotConfig['values'], totals = T): PivotConfig => ({ rows, cols, values, totals });
const sum = (field: number) => ({ field, agg: 'sum' as const });

// 下面的期望值都是 Microsoft Excel 16 对同样的数据和设置（表格形式、无分类汇总）刷新出来的单元格，
// 由 editors/sheet/tools/pivot-oracle.ts 对拍确认过。
describe('透视汇总：布局和 Excel 一致', () => {
  it('一个行字段、一个值字段：分组求和，空白排最后，总计行', () => {
    const r = buildPivot(DATA, cfg([0], [], [sum(3)]));
    expect(r.cells).toEqual([['区域', '求和项:金额'], ['华东', 210], ['华南', 95], ['(空白)', 10], ['总计', 315]]);
    expect([r.headerRows, r.labelCols, r.totalRow, r.totalCol]).toEqual([1, 1, 4, -1]);
  });

  it('两个行字段：外层标签只写每组第一行，总计行', () => {
    expect(buildPivot(DATA, cfg([0, 1], [], [sum(3)])).cells).toEqual([
      ['区域', '产品', '求和项:金额'],
      ['华东', '甲', 130],
      ['', '乙', 80],
      ['华南', '甲', 70],
      ['', '乙', 25],
      ['(空白)', '乙', 10],
      ['总计', '', 315],
    ]);
  });

  it('行 + 列字段 + 一个值：两行标题，行 / 列总计', () => {
    const r = buildPivot(DATA, cfg([0], [1], [sum(3)]));
    expect(r.cells).toEqual([
      ['求和项:金额', '产品', '', ''],
      ['区域', '甲', '乙', '总计'],
      ['华东', 130, 80, 210],
      ['华南', 70, 25, 95],
      ['(空白)', '', 10, 10],
      ['总计', 200, 115, 315],
    ]);
    expect([r.headerRows, r.totalRow, r.totalCol]).toEqual([2, 5, 3]);
  });

  it('多个值字段（没有列字段）：标题上有“值”；平均值是对原始记录算的，总计不是小计的平均', () => {
    const r = buildPivot(DATA, cfg([0], [], [sum(2), { field: 3, agg: 'avg' }]));
    expect(r.cells).toEqual([
      ['', '值', ''],
      ['区域', '求和项:数量', '平均值项:金额'],
      ['华东', 18, 70],
      ['华南', 9, 47.5],
      ['(空白)', 1, 10],
      ['总计', 28, 52.5],
    ]);
  });

  it('行 + 列 + 多个值：三行标题，总计列的标题带“汇总”', () => {
    expect(buildPivot(DATA, cfg([0], [1], [sum(2), sum(3)])).cells).toEqual([
      ['', '产品', '值', '', '', '', ''],
      ['', '甲', '', '乙', '', '求和项:数量汇总', '求和项:金额汇总'],
      ['区域', '求和项:数量', '求和项:金额', '求和项:数量', '求和项:金额', '', ''],
      ['华东', 13, 130, 5, 80, 18, 210],
      ['华南', 7, 70, 2, 25, 9, 95],
      ['(空白)', '', '', 1, 10, 1, 10],
      ['总计', 20, 200, 8, 115, 28, 315],
    ]);
  });

  it('没有行 / 列字段；只有列字段（最左边多一列放值字段名）', () => {
    expect(buildPivot(DATA, cfg([], [], [sum(3)])).cells).toEqual([['求和项:金额'], [315]]);
    expect(buildPivot(DATA, cfg([], [1], [sum(3)])).cells).toEqual([
      ['', '产品', '', ''],
      ['', '甲', '乙', '总计'],
      ['求和项:金额', 200, 115, 315],
    ]);
  });

  it('各种汇总的名称同 Excel（数值计数项 / 计数项 / 最大值项 / 最小值项）', () => {
    const vals = [{ field: 3, agg: 'count' as const }, { field: 3, agg: 'counta' as const }, { field: 3, agg: 'max' as const }, { field: 3, agg: 'min' as const }];
    const r = buildPivot(DATA, cfg([0, 1], [], vals));
    expect(r.cells[1]).toEqual(['区域', '产品', '数值计数项:金额', '计数项:金额', '最大值项:金额', '最小值项:金额']);
    expect(r.cells[2]).toEqual(['华东', '甲', 2, 2, 100, 30]);
    expect(r.cells[7]).toEqual(['总计', '', 6, 6, 100, 10]);
  });

  it('关闭总计行 / 总计列；没选值字段时数第一个没被用上的字段', () => {
    expect(buildPivot(DATA, cfg([0, 1], [], [sum(3)], { rows: false, cols: true })).cells.length).toBe(6);
    const none = buildPivot(DATA, cfg([0], [1], [sum(3)], { rows: false, cols: false })).cells;
    expect(none[1]).toEqual(['区域', '甲', '乙']);
    expect(none.length).toBe(5);
    expect(buildPivot(DATA, cfg([0], [], [])).cells[0]).toEqual(['区域', '计数项:产品']);
  });

  it('数字行字段按数值排序；文字当 0 忽略；空数据只有标题', () => {
    const r = buildPivot(DATA, cfg([2], [], [sum(3)]));
    expect(r.cells.map((x) => x[0])).toEqual(['数量', 1, 2, 3, 5, 7, 10, '总计']);
    const d = [['k', 'v'], ['b', 1], [2, 'x'], ['a', 4], ['a', null]];
    expect(buildPivot(d, cfg([0], [], [{ field: 1, agg: 'counta' }], { rows: false, cols: false })).cells).toEqual([['k', '计数项:v'], [2, 1], ['a', 1], ['b', 1]]);
    expect(buildPivot([['a', 'b']], cfg([0], [], [sum(1)])).cells).toEqual([['a', '求和项:b']]);
  });
});

function sheetWith(rows: (string | number)[][]): Workbook {
  const b = blankWorkbook();
  b.sheets[0].name = '数据';
  const tx = new Tx(b, '数据');
  rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, String(v) === '' ? '' : `${typeof v === 'string' ? "'" : ''}${v}`)));
  tx.commit();
  return b;
}
const at = (b: Workbook, si: number, a: string) => {
  const m = a.match(/^([A-Z]+)(\d+)$/)!;
  return b.sheets[si].cells[Number(m[2]) - 1]?.[colIndex(m[1])]?.v;
};
const range = { r0: 0, c0: 0, r1: 6, c1: 3 };

describe('透视表写进工作表', () => {
  it('新建工作表放结果：只有有内容的格子带样式（标题加粗）；撤销整张表消失', () => {
    const b = sheetWith(DATA as (string | number)[][]);
    expect(fieldNames(b.sheets[0], range)).toEqual(['区域', '产品', '数量', '金额']);
    const tx = new Tx(b, '透视');
    const si = createPivot(tx, 0, range, cfg([0], [1], [sum(3)]));
    const e = tx.commit()!;
    expect(si).toBe(1);
    expect(b.sheets[1].name).toBe('透视表');
    expect([at(b, 1, 'A1'), at(b, 1, 'B1'), at(b, 1, 'D6')]).toEqual(['求和项:金额', '产品', 315]);
    expect(b.sheets[1].cells[0][0]?.s).toBeGreaterThan(0);
    expect(b.sheets[1].cells[2]?.[0]?.s).toBeUndefined();
    expect(b.sheets[1].blankStyles?.flat().filter((x) => x !== undefined).length ?? 0).toBe(0);
    expect(b.sheets[1].pivot?.source).toEqual({ sheet: '数据', range });
    applyEntry(b, e, -1);
    expect(b.sheets.length).toBe(1);
  });

  it('改了源数据后刷新：新分组出现、旧区域清掉；撤销回到刷新前', () => {
    const b = sheetWith(DATA as (string | number)[][]);
    const t1 = new Tx(b, '透视');
    createPivot(t1, 0, range, cfg([0], [], [sum(3)]));
    t1.commit();
    const t2 = new Tx(b, '改源数据');
    setInput(t2, 0, 6, 0, '西北');
    t2.commit();
    const t3 = new Tx(b, '刷新');
    expect(refreshPivot(t3, 1)).toBe(true);
    const e = t3.commit()!;
    expect([at(b, 1, 'A4'), at(b, 1, 'B4'), at(b, 1, 'A5'), at(b, 1, 'B5')]).toEqual(['西北', 10, '总计', 315]);
    const t4 = new Tx(b, '改');
    setInput(t4, 0, 6, 0, '华南');
    t4.commit();
    const t5 = new Tx(b, '刷新');
    refreshPivot(t5, 1);
    t5.commit();
    expect([at(b, 1, 'A3'), at(b, 1, 'B3'), at(b, 1, 'A4'), at(b, 1, 'A5')]).toEqual(['华南', 105, '总计', undefined]);
    applyEntry(b, e, -1);
    expect(at(b, 1, 'B4')).toBe(10);
  });

  it('不是透视表 / 默认设置', () => {
    const b = sheetWith(DATA as (string | number)[][]);
    expect(refreshPivot(new Tx(b, 'x'), 0)).toBe(false);
    expect(defaultConfig(b.sheets[0], range)).toEqual({ rows: [0], cols: [], values: [{ field: 3, agg: 'sum' }], totals: { rows: true, cols: true } });
  });
});

describe('Excel 真部件', () => {
  const make = (c: PivotConfig) => {
    const b = sheetWith(DATA as (string | number)[][]);
    const tx = new Tx(b, '透视');
    createPivot(tx, 0, range, c);
    tx.commit();
    return b;
  };
  const f = (bytes: Uint8Array) => unzipSync(bytes);
  const txt = (files: Record<string, Uint8Array>, p: string) => strFromU8(files[p]);

  it('保存：缓存定义 / 记录 / 透视表 + 关系 + 内容类型 + workbook 里的 pivotCaches', () => {
    const b = make(cfg([0], [1], [sum(3)]));
    const files = f(saveXlsx(b));
    expect(Object.keys(files).filter((k) => /pivot/.test(k)).sort()).toEqual([
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
      'xl/pivotCache/pivotCacheDefinition1.xml',
      'xl/pivotCache/pivotCacheRecords1.xml',
      'xl/pivotTables/_rels/pivotTable1.xml.rels',
      'xl/pivotTables/pivotTable1.xml',
    ]);
    const ct = txt(files, '[Content_Types].xml');
    expect(ct).toContain('<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>');
    expect(ct).toContain('pivotCacheDefinition+xml');
    expect(ct).toContain('pivotCacheRecords+xml');
    expect(txt(files, 'xl/pivotTables/_rels/pivotTable1.xml.rels')).toMatch(/Type="[^"]*\/pivotCacheDefinition" Target="\.\.\/pivotCache\/pivotCacheDefinition1\.xml"/);
    expect(txt(files, 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels')).toMatch(/Type="[^"]*\/pivotCacheRecords" Target="pivotCacheRecords1\.xml"/);
    expect(txt(files, 'xl/worksheets/_rels/sheet2.xml.rels')).toMatch(/Type="[^"]*\/pivotTable" Target="\.\.\/pivotTables\/pivotTable1\.xml"/);
    const wb = txt(files, 'xl/workbook.xml');
    expect(wb).toMatch(/<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/);
    expect(wb.indexOf('<pivotCaches')).toBeGreaterThan(wb.indexOf('<calcPr'));
    expect(txt(files, 'xl/_rels/workbook.xml.rels')).toMatch(/Type="[^"]*\/pivotCacheDefinition" Target="pivotCache\/pivotCacheDefinition1\.xml"/);
  });

  it('缓存：字段、共享项（空白 <m/>）、记录；打开时刷新', () => {
    const b = make(cfg([0], [], [sum(3)]));
    const files = f(saveXlsx(b));
    const def = txt(files, 'xl/pivotCache/pivotCacheDefinition1.xml');
    expect(def).toContain('refreshOnLoad="1"');
    expect(def).toContain('<worksheetSource ref="A1:D7" sheet="数据"/>');
    expect(def).toContain('<cacheField name="区域" numFmtId="0"><sharedItems containsBlank="1" count="3"><s v="华东"/><s v="华南"/><m/></sharedItems></cacheField>');
    expect(def).toContain('<cacheField name="金额" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="100"/></cacheField>');
    const rec = txt(files, 'xl/pivotCache/pivotCacheRecords1.xml');
    expect(rec).toContain('count="6"');
    expect(rec).toContain('<r><x v="0"/><x v="0"/><n v="10"/><n v="100"/></r>');
    expect(rec).toContain('<r><x v="2"/><x v="1"/><n v="1"/><n v="10"/></r>');
  });

  it('透视表定义：和 Excel 保存的结构一致（行 / 列 / 值、各级位置、r 省略、总计开关）', () => {
    const b = make(cfg([0, 1], [], [sum(3)]));
    const t = txt(f(saveXlsx(b)), 'xl/pivotTables/pivotTable1.xml');
    expect(t).toContain('<location ref="A1:C7" firstHeaderRow="1" firstDataRow="1" firstDataCol="2"/>');
    expect(t).toContain('<rowFields count="2"><field x="0"/><field x="1"/></rowFields>');
    expect(t).toContain('<rowItems count="6"><i><x/><x/></i><i r="1"><x v="1"/></i><i><x v="1"/><x/></i><i r="1"><x v="1"/></i><i><x v="2"/><x v="1"/></i><i t="grand"><x/></i></rowItems>');
    expect(t).toContain('<dataField name="求和项:金额" fld="3" baseField="0" baseItem="0"/>');
    expect(t).toContain('<colItems count="1"><i/></colItems>');
    const c = pivotXml(b.sheets[0], { source: { sheet: '数据', range }, config: cfg([0], [1], [sum(2), { field: 3, agg: 'avg' }]), extent: { r0: 0, c0: 0, r1: 6, c1: 6 } }, 'A1:G7', 'p', 3);
    expect(c.table).toContain('<colFields count="2"><field x="1"/><field x="-2"/></colFields>');
    expect(c.table).toContain('<colItems count="6"><i><x/><x/></i><i r="1" i="1"><x v="1"/></i><i><x v="1"/><x/></i><i r="1" i="1"><x v="1"/></i><i t="grand"><x/></i><i t="grand" i="1"><x/></i></colItems>');
    expect(c.table).toContain('subtotal="average"');
    expect(c.table).toContain('firstDataRow="3"');
    const off = pivotXml(b.sheets[0], { source: { sheet: '数据', range }, config: cfg([0], [1], [sum(3)], { rows: false, cols: false }), extent: { r0: 0, c0: 0, r1: 4, c1: 2 } }, 'A1:C5', 'p', 1);
    expect(off.table).toContain('rowGrandTotals="0"');
    expect(off.table).toContain('colGrandTotals="0"');
    expect(off.table).not.toContain('t="grand"');
  });

  it('保存后再改源数据并刷新、再保存：原地覆盖同一组部件，不新增缓存', () => {
    const b = make(cfg([0], [], [sum(3)]));
    const out1 = saveXlsx(b);
    markSaved(b, out1);
    expect(b.sheets[1].pivot?.parts).toMatchObject({ cacheId: 1, table: 'xl/pivotTables/pivotTable1.xml' });
    const t = new Tx(b, '改');
    setInput(t, 0, 6, 0, '西北');
    t.commit();
    const t2 = new Tx(b, '刷新');
    refreshPivot(t2, 1);
    t2.commit();
    const files = f(saveXlsx(b));
    expect(Object.keys(files).filter((k) => /pivotTable\d|pivotCacheDefinition\d/.test(k) && !/_rels/.test(k))).toEqual(['xl/pivotCache/pivotCacheDefinition1.xml', 'xl/pivotTables/pivotTable1.xml']);
    expect(txt(files, 'xl/pivotTables/pivotTable1.xml')).toContain('<rowItems count="4">');
    expect(txt(files, 'xl/workbook.xml').match(/<pivotCache /g)!.length).toBe(1);
  });

  it('重新打开后没改：字节不变（文件里的透视表部件原样保留）', () => {
    const b = make(cfg([0], [1], [sum(3)]));
    const out = saveXlsx(b);
    const b2 = readWorkbook(out);
    b2.source = out;
    expect(Array.from(saveXlsx(b2))).toEqual(Array.from(out));
    const t = new Tx(b2, '改');
    setInput(t, 1, 8, 0, 'x');
    t.commit();
    const files = f(saveXlsx(b2));
    expect(files['xl/pivotTables/pivotTable1.xml']).toEqual(f(out)['xl/pivotTables/pivotTable1.xml']);
  });

  it('不能写成真透视表的（值字段同时是行字段、两个列字段）：只留普通单元格', () => {
    const b = make(cfg([0], [1, 2], [sum(3)]));
    expect(Object.keys(f(saveXlsx(b))).filter((k) => /pivot/.test(k))).toEqual([]);
    const b2 = make(cfg([0], [], [sum(0)]));
    expect(Object.keys(f(saveXlsx(b2))).filter((k) => /pivot/.test(k))).toEqual([]);
  });
});

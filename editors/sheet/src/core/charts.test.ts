import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { addChart, chartXml, deleteDrawing, guessSource, moveDrawing, patchChartXml, seriesFromRange, updateChart } from './charts';
import { insertDelete, setInput, Tx } from './ops';
import { readDrawings } from './drawing';
import { saveXlsx } from './save';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => unknown) => {
  const tx = new Tx(b, 'x');
  const out = fn(tx);
  tx.commit();
  return out;
};
const fill = (b: Workbook, rows: (string | number)[][]) => edit(b, (tx) => rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, typeof v === 'string' ? `'${v}` : String(v)))));
const parts = (bytes: Uint8Array) => unzipSync(bytes);
const txt = (f: Record<string, Uint8Array>, p: string) => strFromU8(f[p]);
const reopen = (bytes: Uint8Array) => {
  const b = readWorkbook(bytes);
  b.source = bytes;
  return b;
};
const A = (c: number, r: number) => ({ c, cOff: 0, r, rOff: 0 });

const DATA = [['月份', '销量', '利润'], ['1月', 10, 3], ['2月', 20, 5], ['3月', 15, 4]];
const g = { r0: 0, c0: 0, r1: 3, c1: 2 };

function withChart(type: 'col' | 'line' | 'pie' = 'col'): Workbook {
  const b = blankWorkbook();
  fill(b, DATA);
  const series = seriesFromRange(b, 0, g, guessSource(b.sheets[0], g));
  edit(b, (tx) => addChart(tx, 0, { type, title: '销量', series }, A(4, 1), A(10, 17)));
  return b;
}

describe('图表数据', () => {
  it('猜标题行和分类列；按列 / 按行取系列；引用和缓存', () => {
    const b = blankWorkbook();
    fill(b, DATA);
    const src = guessSource(b.sheets[0], g);
    expect(src).toEqual({ header: true, labels: true, byRows: false });
    const s = seriesFromRange(b, 0, g, src);
    expect(s.map((x) => [x.name, x.nameRef, x.catRef, x.valRef])).toEqual([
      ['销量', 'Sheet1!$B$1', 'Sheet1!$A$2:$A$4', 'Sheet1!$B$2:$B$4'],
      ['利润', 'Sheet1!$C$1', 'Sheet1!$A$2:$A$4', 'Sheet1!$C$2:$C$4'],
    ]);
    expect(s[0].cats).toEqual(['1月', '2月', '3月']);
    expect(s[1].vals).toEqual([3, 5, 4]);
    const rows = seriesFromRange(b, 0, { r0: 0, c0: 0, r1: 2, c1: 3 }, { header: true, labels: true, byRows: true });
    expect(rows.map((x) => [x.name, x.valRef])).toEqual([['1月', 'Sheet1!$B$2:$D$2'], ['2月', 'Sheet1!$B$3:$D$3']]);
    // 单列纯数字：没有标题 / 分类
    fill(b, [[1], [2]]);
    expect(guessSource(b.sheets[0], { r0: 8, c0: 8, r1: 9, c1: 8 })).toMatchObject({ labels: false });
  });
});

describe('chart XML', () => {
  const series = [{ name: '销量', nameRef: 'Sheet1!$B$1', catRef: 'Sheet1!$A$2:$A$4', valRef: 'Sheet1!$B$2:$B$4', cats: ['1月', '2月', '3月'], vals: [10, 20, 15] }];
  it('各类型生成标准结构；标题 / 引用 / 缓存转义', () => {
    const col = chartXml({ type: 'col', title: 'A & B', series });
    expect(col).toContain('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>');
    expect(col).toContain('<a:t>A &amp; B</a:t>');
    expect(col).toContain('<c:f>Sheet1!$B$2:$B$4</c:f>');
    expect(col).toContain('<c:autoTitleDeleted val="0"/>');
    expect(chartXml({ type: 'bar', series })).toContain('<c:barDir val="bar"/>');
    expect(chartXml({ type: 'line', series })).toContain('<c:lineChart>');
    expect(chartXml({ type: 'area', series })).toContain('<c:areaChart>');
    const pie = chartXml({ type: 'pie', series: [...series, { ...series[0], name: '利润' }] });
    expect(pie).toContain('<c:pieChart>');
    expect(pie.match(/<c:ser>/g)!.length).toBe(1);
    expect(pie).not.toContain('<c:catAx>');
    expect(chartXml({ type: 'col', series })).toContain('<c:autoTitleDeleted val="1"/>');
  });

  it('在原 XML 上改标题 / 引用（保留样式）；系列个数变了返回 null', () => {
    const xml = chartXml({ type: 'col', title: '旧标题', series }).replace('<c:spPr><a:solidFill><a:srgbClr val="4472C4"/>', '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/>');
    const out = patchChartXml(xml, { type: 'col', title: '新标题', series: [{ ...series[0], valRef: 'Sheet1!$C$2:$C$4', vals: [3, 5, 4] }] })!;
    expect(out).toContain('<a:t>新标题</a:t>');
    expect(out).toContain('<c:f>Sheet1!$C$2:$C$4</c:f>');
    expect(out).toContain('srgbClr val="FF0000"');
    expect(out).not.toContain('旧标题');
    expect(patchChartXml(xml, { type: 'col', series: [] })).toBeNull();
    // 去掉标题
    expect(patchChartXml(xml, { type: 'col', series })!).not.toContain('<c:title>');
  });
});

describe('写出 xlsx', () => {
  it('新建图表：drawing + chart 部件、关系、内容类型、<drawing>；读回来一致', () => {
    const b = withChart();
    const out = saveXlsx(b);
    const f = parts(out);
    expect(txt(f, 'xl/drawings/drawing1.xml')).toContain('<xdr:from><xdr:col>4</xdr:col>');
    expect(txt(f, 'xl/drawings/drawing1.xml')).toContain('r:id="rId1"');
    expect(txt(f, 'xl/drawings/_rels/drawing1.xml.rels')).toMatch(/Type="[^"]*\/chart" Target="\.\.\/charts\/chart1\.xml"/);
    expect(txt(f, 'xl/worksheets/_rels/sheet1.xml.rels')).toMatch(/Type="[^"]*\/drawing" Target="\.\.\/drawings\/drawing1\.xml"/);
    expect(txt(f, '[Content_Types].xml')).toContain('<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
    expect(txt(f, '[Content_Types].xml')).toContain('<Override PartName="/xl/drawings/drawing1.xml"');
    const sheet = txt(f, 'xl/worksheets/sheet1.xml');
    expect(sheet).toMatch(/<drawing r:id="rId1"\/>/);
    expect(sheet.indexOf('<pageMargins')).toBeLessThan(sheet.indexOf('<drawing '));
    const d = readDrawings(f, 'xl/worksheets/sheet1.xml').drawings;
    expect(d.length).toBe(1);
    expect(d[0].chart).toMatchObject({ type: 'col', title: '销量' });
    expect(d[0].chart!.series.map((s) => s.valRef)).toEqual(['Sheet1!$B$2:$B$4', 'Sheet1!$C$2:$C$4']);
    expect([d[0].from.c, d[0].from.r, d[0].to!.c, d[0].to!.r]).toEqual([4, 1, 10, 17]);
  });

  it('没改：字节不变；保存后模型指向新文件里的部件，继续改名 / 移动', () => {
    const b = withChart();
    const out = saveXlsx(b);
    const b2 = reopen(out);
    expect(Array.from(saveXlsx(b2))).toEqual(Array.from(out));
    edit(b2, (tx) => moveDrawing(tx, 0, 0, A(1, 20), A(7, 36)));
    const f = parts(saveXlsx(b2));
    expect(txt(f, 'xl/drawings/drawing1.xml')).toContain('<xdr:from><xdr:col>1</xdr:col>');
    expect(txt(f, 'xl/charts/chart1.xml')).toBe(txt(parts(out), 'xl/charts/chart1.xml'));
  });

  it('只改位置 / 标题 / 数据：原 chart 样式保留；改类型：重新生成', () => {
    const b2 = reopen(saveXlsx(withChart()));
    const d = b2.sheets[0].drawings![0];
    edit(b2, (tx) => updateChart(tx, 0, 0, { ...d.chart!, title: '改了标题' }));
    const c1 = txt(parts(saveXlsx(b2)), 'xl/charts/chart1.xml');
    expect(c1).toContain('<a:t>改了标题</a:t>');
    expect(c1).toContain('<c:barChart>');
    edit(b2, (tx) => updateChart(tx, 0, 0, { ...b2.sheets[0].drawings![0].chart!, type: 'line' }));
    const out = saveXlsx(b2);
    expect(txt(parts(out), 'xl/charts/chart1.xml')).toContain('<c:lineChart>');
    expect(readDrawings(parts(out), 'xl/worksheets/sheet1.xml').drawings[0].chart!.type).toBe('line');
  });

  it('删除图表：anchor、chart 部件、关系、内容类型都去掉（drawing 里别的东西保留）', () => {
    const b = withChart();
    edit(b, (tx) => addChart(tx, 0, { type: 'pie', series: b.sheets[0].drawings![0].chart!.series.slice(0, 1) }, A(1, 25), A(6, 40)));
    const out = saveXlsx(b);
    expect(Object.keys(parts(out)).filter((k) => /charts\/chart\d/.test(k)).length).toBe(2);
    const b2 = reopen(out);
    edit(b2, (tx) => deleteDrawing(tx, 0, 0));
    const f = parts(saveXlsx(b2));
    expect(Object.keys(f).filter((k) => /charts\/chart\d/.test(k)).length).toBe(1);
    expect(readDrawings(f, 'xl/worksheets/sheet1.xml').drawings.map((x) => x.chart!.type)).toEqual(['pie']);
    expect(txt(f, 'xl/drawings/_rels/drawing1.xml.rels').match(/<Relationship\b/g)!.length).toBe(1);
  });

  it('图表引用随插入 / 删除行列移动；保存后写进 chart', () => {
    const b = reopen(saveXlsx(withChart()));
    edit(b, (tx) => insertDelete(tx, 0, 'row', 0, 2));
    expect(b.sheets[0].drawings![0].chart!.series[0].valRef).toBe('Sheet1!$B$4:$B$6');
    const out = saveXlsx(b);
    const d = readDrawings(parts(out), 'xl/worksheets/sheet1.xml').drawings[0];
    expect(d.chart!.series[0].valRef).toBe('Sheet1!$B$4:$B$6');
    expect(d.from.r).toBe(3);
  });
});

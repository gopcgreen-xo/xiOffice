// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { readDocx, writeDocx } from '../ooxml/docx';
import { applyCellBorders, cellRect, deleteColumn, deleteRow, distributeColumns, insertColumn, insertRow, insertRaw, insertTable, mergeTableCells, moveColumnLine, setCellFormat, splitCell, unmergeCell } from './ops';
import { newDocument, paraText, type Table } from './model';
import { layoutDocument } from './layout';

const cellPos = (r: number, c: number) => ({ story: 'body' as const, path: [0, r, c, 0], offset: 0 });

/** 新文档里插一张 rows × cols 的表，每格写上 “行列” 编号 */
function makeTable(rows: number, cols: number) {
  const doc = newDocument();
  insertTable(doc, { story: 'body', path: [0], offset: 0 }, rows, cols);
  const t = doc.body[0] as Table;
  t.rows.forEach((row, r) => row.cells.forEach((_, c) => insertRaw(doc, cellPos(r, c), `${r}${c}`)));
  return { doc, t };
}
const texts = (t: Table) => t.rows.map((row) => row.cells.map((c) => c.paragraphs.map(paraText).join('|')));

describe('表格：合并、拆分、底纹、列宽', () => {
  it('横向合并：gridSpan 相加，内容并到左边的格子', () => {
    const { doc, t } = makeTable(2, 3);
    expect(mergeTableCells(doc, cellPos(0, 0), cellPos(0, 1))).toEqual(cellPos(0, 0));
    expect(t.rows[0].cells.map((c) => c.gridSpan)).toEqual([2, 1]);
    expect(texts(t)[0]).toEqual(['00|01', '02']);
  });

  it('矩形合并：第一行 restart、下面各行 continue，取消合并后恢复格数', () => {
    const { doc, t } = makeTable(3, 3);
    mergeTableCells(doc, cellPos(0, 1), cellPos(1, 2));
    expect(t.rows[0].cells.map((c) => [c.gridSpan, c.vMerge])).toEqual([[1, 'none'], [2, 'restart']]);
    expect(t.rows[1].cells.map((c) => [c.gridSpan, c.vMerge])).toEqual([[1, 'none'], [2, 'continue']]);
    expect(texts(t)[0][1]).toBe('01|02|11|12');
    expect(cellRect(t, 1, 1)).toEqual({ r0: 0, r1: 1, g0: 1, g1: 2 });
    unmergeCell(doc, cellPos(1, 1));
    expect(t.rows.map((r) => r.cells.length)).toEqual([3, 3, 3]);
    expect(t.rows.flatMap((r) => r.cells.map((c) => c.vMerge))).not.toContain('continue');
  });

  it('矩形会切开已有的合并单元格时不合并', () => {
    const { doc, t } = makeTable(3, 3);
    mergeTableCells(doc, cellPos(0, 0), cellPos(1, 0)); // 第 1 列上两行纵向合并
    expect(mergeTableCells(doc, cellPos(1, 1), cellPos(2, 0))).toBeNull(); // 只包含了合并格的下半截
    expect(t.rows[2].cells.length).toBe(3);
  });

  it('拆分成 2 列：该网格列一分为二，其它行的格子跨两列，总宽不变', () => {
    const { doc, t } = makeTable(2, 2);
    const total = t.cols.reduce((a, b) => a + b, 0);
    splitCell(doc, cellPos(0, 0), 2);
    expect(t.cols.length).toBe(3);
    expect(t.cols.reduce((a, b) => a + b, 0)).toBeCloseTo(total);
    expect(t.rows[0].cells.map((c) => c.gridSpan)).toEqual([1, 1, 1]);
    expect(t.rows[1].cells.map((c) => c.gridSpan)).toEqual([2, 1]);
  });

  it('底纹：读写 w:shd 的 fill，排版产生底色块', () => {
    const { doc, t } = makeTable(2, 2);
    setCellFormat(doc, cellPos(0, 0), cellPos(0, 1), { shading: 'FFFF00' });
    expect(t.rows[0].cells.map((c) => c.shading)).toEqual(['FFFF00', 'FFFF00']);
    const xml = strFromU8(unzipSync(writeDocx(doc))['word/document.xml']);
    expect(xml.match(/<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"\/>/g)?.length).toBe(2);
    const back = readDocx(writeDocx(doc)).body.find((b) => b.kind === 'table') as Table;
    expect(back.rows[0].cells[0].shading).toBe('FFFF00');
    expect(back.rows[1].cells[0].shading).toBeUndefined();
    setCellFormat(doc, cellPos(0, 0), cellPos(0, 0), { shading: null, borders: false });
    expect(t.rows[0].cells[0].shading).toBeUndefined();
    expect(t.rows[0].cells[0].borders).toBe(false);
  });

  it('排版：底纹块和格子一样大', () => {
    const { doc } = makeTable(1, 2);
    setCellFormat(doc, cellPos(0, 1), cellPos(0, 1), { shading: 'D9D9D9' });
    const m = { width: () => 0.5, metrics: () => ({ ascent: 0.8, descent: 0.2 }) };
    const lay = layoutDocument(doc, m);
    expect(lay.pages[0].fills.length).toBe(1);
    expect(lay.pages[0].fills[0].color).toBe('#D9D9D9');
    expect(lay.pages[0].fills[0].h).toBeGreaterThan(0);
  });

  it('拖动列线：相邻两列此消彼长；最后一列直接变宽；平均分布', () => {
    const { t } = makeTable(1, 3);
    const [a, b] = t.cols;
    moveColumnLine(t, 0, 20);
    expect(t.cols[0]).toBeCloseTo(a + 20);
    expect(t.cols[1]).toBeCloseTo(b - 20);
    moveColumnLine(t, 0, 10000);
    expect(t.cols[1]).toBe(12);
    const last = t.cols[2];
    moveColumnLine(t, 2, -10);
    expect(t.cols[2]).toBeCloseTo(last - 10);
    distributeColumns(t);
    expect(new Set(t.cols.map((v) => v.toFixed(3))).size).toBe(1);
  });

  it('有合并单元格的表格也能插入 / 删除列', () => {
    const { doc, t } = makeTable(2, 3);
    mergeTableCells(doc, cellPos(0, 0), cellPos(0, 1)); // 第一行前两列合并
    insertColumn(doc, cellPos(1, 0)); // 在第二行第一格右侧插列：线穿过第一行的合并格
    expect(t.cols.length).toBe(4);
    expect(t.rows[0].cells.map((c) => c.gridSpan)).toEqual([3, 1]);
    expect(t.rows[1].cells.map((c) => c.gridSpan)).toEqual([1, 1, 1, 1]);
    const total = t.cols.reduce((a, b) => a + b, 0);
    deleteColumn(doc, cellPos(0, 0)); // 删除合并格占的三列
    expect(t.cols.length).toBe(1);
    expect(t.rows.map((r) => r.cells.length)).toEqual([1, 1]);
    expect(t.cols[0]).toBeCloseTo(total);
  });

  it('插入行穿过纵向合并格时接着合并；删掉合并的第一行时下一行接手', () => {
    const { doc, t } = makeTable(3, 2);
    mergeTableCells(doc, cellPos(0, 0), cellPos(2, 0));
    insertRow(doc, cellPos(0, 1), true);
    expect(t.rows.map((r) => r.cells[0].vMerge)).toEqual(['restart', 'continue', 'continue', 'continue']);
    insertRow(doc, cellPos(3, 1), true); // 合并区最下面再加一行：不合并
    expect(t.rows[4].cells[0].vMerge).toBe('none');
    deleteRow(doc, cellPos(0, 1));
    expect(t.rows.map((r) => r.cells[0].vMerge)).toEqual(['restart', 'continue', 'continue', 'none']);
    expect(paraText(t.rows[0].cells[0].paragraphs[0])).toBe('00');
  });
});

describe('百分比宽度的表格（w:tblW w:type="pct"）', () => {
  it('表宽 = 百分比 × 版心宽度，各列按网格列宽的比例缩放；读写往返', () => {
    const { doc, t } = makeTable(1, 2);
    t.cols = [100, 300];
    t.widthPct = 0.5;
    const L = layoutDocument(doc, { width: () => 1, metrics: () => ({ ascent: 0.88, descent: 0.12 }) });
    const cells = L.pages[0].cells!;
    const tw = doc.setup.width - doc.setup.marginLeft - doc.setup.marginRight;
    expect(cells[0].w).toBeCloseTo(tw * 0.5 * 0.25);
    expect(cells[1].w).toBeCloseTo(tw * 0.5 * 0.75);
    const back = readDocx(writeDocx(doc)).body[0] as Table;
    expect(back.widthPct).toBeCloseTo(0.5);
    expect(strFromU8(unzipSync(writeDocx(doc))['word/document.xml'])).toContain('<w:tblW w:w="2500" w:type="pct"/>');
  });
});

describe('表格里的行与行网格（兼容选项 w:adjustLineHeightInTable）', () => {
  it('没有该选项时单元格行高按字体算；有时对齐网格；读写往返', () => {
    const fake = { width: () => 1, metrics: () => ({ ascent: 0.88, descent: 0.12 }) };
    const { doc } = makeTable(1, 1);
    const cellLine = () => layoutDocument(doc, fake).pages[0].body.find((l) => l.path.length > 1)!.line;
    expect(cellLine().height).toBeCloseTo(10.5);
    doc.setup.tableGrid = true;
    expect(cellLine().height).toBeCloseTo(15.6);
    expect(readDocx(writeDocx(doc)).setup.tableGrid).toBe(true);
  });
});

describe('表格样式（tblStyle 的条件格式）', () => {
  it('首行底纹和加粗、镶边行、首列、内部横线按 tblLook 合成；单元格自己的底纹优先', async () => {
    const { zipSync, strToU8 } = await import('fflate');
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const bd = (c: string) => `<w:top w:val="single" w:sz="8" w:color="${c}"/><w:bottom w:val="single" w:sz="8" w:color="${c}"/>`;
    const styles =
      `<w:styles ${W}><w:style w:type="table" w:styleId="T"><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders>${bd('111111')}<w:insideH w:val="single" w:sz="4" w:color="222222"/></w:tblBorders></w:tblPr>` +
      `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="0000FF"/></w:tcPr></w:tblStylePr>` +
      `<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="DDDDDD"/></w:tcPr></w:tblStylePr>` +
      `<w:tblStylePr w:type="firstCol"><w:rPr><w:i/></w:rPr></w:tblStylePr></w:style></w:styles>`;
    const tc = (t: string, shd = '') => `<w:tc><w:tcPr>${shd}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const rows = [0, 1, 2, 3].map((r) => `<w:tr>${tc(`a${r}`)}${tc(`b${r}`, r === 3 ? '<w:shd w:val="clear" w:fill="FF0000"/>' : '')}</w:tr>`).join('');
    const doc = `<w:document ${W}><w:body><w:tbl><w:tblPr><w:tblStyle w:val="T"/><w:tblLook w:firstRow="1" w:firstColumn="1" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl><w:p/></w:body></w:document>`;
    const d = readDocx(zipSync({ 'word/document.xml': strToU8(doc), 'word/styles.xml': strToU8(styles) }));
    const t = d.body[0] as Table;
    const cell = (r: number, c: number) => t.rows[r].cells[c];
    const run = (r: number, c: number) => cell(r, c).paragraphs[0].runs[0].props;
    expect(cell(0, 1).shading).toBe('0000FF');
    expect(run(0, 1).bold).toBe(true);
    expect(run(1, 1).bold).toBe(false);
    // 镶边：首行之后第 1、3 行是 band1
    expect(cell(1, 1).shading).toBe('DDDDDD');
    expect(cell(2, 1).shading).toBeUndefined();
    expect(cell(3, 1).shading).toBe('FF0000');
    expect(run(2, 0).italic).toBe(true);
    // 外框 1 磅 111111，内部横线 0.5 磅 222222
    expect(cell(0, 0).edges?.top).toEqual({ w: 1, color: '111111' });
    expect(cell(1, 0).edges?.top).toEqual({ w: 0.5, color: '222222' });
    expect(cell(1, 0).edges?.left).toBeNull();
  });
});

describe('表格：左侧插入列、单元格垂直对齐、框线线型', () => {
  const fake = { width: () => 1, metrics: () => ({ ascent: 0.88, descent: 0.12 }) };

  it('左侧插入列：当前格最左的网格列一分为二，新格在左边，横跨这条线的合并格跨列数加 1', () => {
    const { doc, t } = makeTable(2, 2);
    mergeTableCells(doc, cellPos(1, 0), cellPos(1, 1));
    expect(insertColumn(doc, cellPos(0, 1), true)).toEqual(cellPos(0, 1));
    expect(t.cols.length).toBe(3);
    expect(texts(t)[0]).toEqual(['00', '', '01']);
    expect(t.rows[1].cells.map((c) => c.gridSpan)).toEqual([3]);
  });

  it('左侧插入第一列：每行最前面加一格', () => {
    const { doc, t } = makeTable(2, 2);
    insertColumn(doc, cellPos(1, 0), true);
    expect(texts(t)).toEqual([['', '00', '01'], ['', '10', '11']]);
  });

  it('垂直对齐：作用于选中的格子，排版时文字放到格子中间', () => {
    const { doc, t } = makeTable(1, 2);
    t.rows[0].height = 100;
    t.rows[0].heightRule = 'exact';
    setCellFormat(doc, cellPos(0, 0), cellPos(0, 0), { vAlign: 'center' });
    expect(t.rows[0].cells.map((c) => c.vAlign)).toEqual(['center', 'top']);
    const lines = layoutDocument(doc, fake).pages[0].body.filter((l) => l.path.length > 1);
    const y = (c: number) => lines.find((l) => l.path[2] === c)!.y;
    expect(y(0) - y(1)).toBeGreaterThan(30);
  });

  it('外侧框线：只改选区外圈的边，相邻格子对着的边一起改；排版按线型、颜色画', () => {
    const { doc, t } = makeTable(3, 3);
    const line = { w: 1.5, color: 'FF0000', style: 'dashed' as const };
    applyCellBorders(doc, cellPos(1, 1), cellPos(1, 1), 'outside', line);
    const mid = t.rows[1].cells[1];
    expect(mid.bdr).toEqual({ top: line, bottom: line, left: line, right: line });
    expect(t.rows[0].cells[1].bdr).toEqual({ bottom: line });
    expect(t.rows[1].cells[0].bdr).toEqual({ right: line });
    expect(t.rows[1].cells[2].bdr).toEqual({ left: line });
    // 没设的边还按表格框线画
    expect(t.rows[0].cells[1].edges?.top).toEqual({ w: 0.5, color: 'auto' });
    const segs = layoutDocument(doc, fake).pages[0].borders.filter((b) => b.style === 'dashed');
    expect(segs.length).toBeGreaterThanOrEqual(4);
    expect(segs.every((b) => b.color === '#FF0000' && b.w === 1.5)).toBe(true);
  });

  it('内部框线 / 无框线：矩形内部的边，或全部边', () => {
    const { doc, t } = makeTable(2, 2);
    applyCellBorders(doc, cellPos(0, 0), cellPos(1, 1), 'none', null);
    expect(t.rows.flatMap((r) => r.cells.map((c) => Object.values(c.edges!).every((e) => e === null)))).toEqual([true, true, true, true]);
    const line = { w: 0.5, color: 'auto' };
    applyCellBorders(doc, cellPos(0, 0), cellPos(1, 1), 'insideH', line);
    expect(t.rows[0].cells[0].edges).toEqual({ top: null, bottom: line, left: null, right: null });
    expect(t.rows[1].cells[1].edges).toEqual({ top: line, bottom: null, left: null, right: null });
  });

  it('保存：设过的边写成 tcBorders（线型、线宽 1/8 磅、颜色，无框线 nil）；读回来线型不变', () => {
    const { doc, t } = makeTable(1, 2);
    applyCellBorders(doc, cellPos(0, 0), cellPos(0, 0), 'all', { w: 1, color: '0000FF', style: 'double' });
    applyCellBorders(doc, cellPos(0, 1), cellPos(0, 1), 'right', null);
    t.rows[0].cells[1].vAlign = 'bottom';
    const xml = strFromU8(unzipSync(writeDocx(doc))['word/document.xml']);
    expect(xml).toContain('<w:tcBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="0000FF"/>');
    expect(xml).toContain('<w:right w:val="nil"/>');
    expect(xml).toContain('<w:vAlign w:val="bottom"/>');
    const back = readDocx(writeDocx(doc)).body[0] as Table;
    expect(back.rows[0].cells[0].edges?.top).toEqual({ w: 1, color: '0000FF', style: 'double' });
    expect(back.rows[0].cells[1].edges?.right).toBeNull();
    expect(back.rows[0].cells[1].vAlign).toBe('bottom');
  });

  it('读入：虚线 / 点线认作 dashed / dotted，单实线和不认识的线型不记线型', () => {
    const { doc } = makeTable(1, 1);
    applyCellBorders(doc, cellPos(0, 0), cellPos(0, 0), 'all', { w: 0.5, color: 'auto', style: 'dotted' });
    const xml = strFromU8(unzipSync(writeDocx(doc))['word/document.xml'])
      .replace('<w:top w:val="dotted"', '<w:top w:val="dashSmallGap"')
      .replace('<w:left w:val="dotted"', '<w:left w:val="wave"');
    const files = unzipSync(writeDocx(doc));
    files['word/document.xml'] = new TextEncoder().encode(xml);
    const c = (readDocx(zipSync(files)).body[0] as Table).rows[0].cells[0];
    expect(c.edges?.top?.style).toBe('dashed');
    expect(c.edges?.bottom?.style).toBe('dotted');
    expect(c.edges?.left?.style).toBeUndefined();
  });
});


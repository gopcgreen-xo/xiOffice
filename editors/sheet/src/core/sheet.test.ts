import { describe, expect, it } from 'vitest';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { writeXlsx } from '@wz/editor-shared/xlsx';
import { formatGeneral, formatNumber, isDateFormat } from './numfmt';
import { parseRef, readCsv, readWorkbook, type SheetData } from './workbook';
import { dateText, fromXls } from './xls';
import { buildAxis, buildMergeIndex, cellPadding, collect, edgeAt, expandToMerges, findAll, indexAt, moveBy, rangeName, statsOf, toHtml, toMarkdown, toTsv } from './grid';

describe('numfmt', () => {
  it('general', () => {
    expect([formatGeneral(12), formatGeneral(0.1 + 0.2), formatGeneral(1234567.891), formatGeneral(1e12)]).toEqual(['12', '0.3', '1234567.891', '1E+12']);
  });

  it('decimals, thousands, percent, scientific', () => {
    expect(formatNumber(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatNumber(-1234.5, '#,##0')).toBe('-1,235');
    expect(formatNumber(0.1234, '0.0%')).toBe('12.3%');
    expect(formatNumber(12345, '0.00E+00')).toBe('1.23E+04');
    expect(formatNumber(0.5, '0.00')).toBe('0.50');
    expect(formatNumber(3, '000')).toBe('003');
    expect(formatNumber(1234567, '#,##0,')).toBe('1,235');
  });

  it('literals, currency and sections', () => {
    expect(formatNumber(12.5, '"¥"#,##0.00')).toBe('¥12.50');
    expect(formatNumber(12.5, '[$¥-804]#,##0.00')).toBe('¥12.50');
    expect(formatNumber(8, '0"台"')).toBe('8台');
    expect(formatNumber(-5, '#,##0;(#,##0)')).toBe('(5)');
    expect(formatNumber(0, '#,##0;-#,##0;"-"')).toBe('-');
    expect(formatNumber(-5, '#,##0.00_);[Red]\\(#,##0.00\\)')).toBe('(5.00)');
    expect(formatNumber(7, '@')).toBe('7');
  });

  it('dates and times', () => {
    // 2026-09-29 = 46294
    expect(formatNumber(46294, 'yyyy/m/d')).toBe('2026/9/29');
    expect(formatNumber(46294, 'yyyy-mm-dd')).toBe('2026-09-29');
    expect(formatNumber(46294, 'yyyy"年"m"月"d"日"')).toBe('2026年9月29日');
    expect(formatNumber(46294.5, 'yyyy/m/d h:mm')).toBe('2026/9/29 12:00');
    expect(formatNumber(46294.75, 'h:mm AM/PM')).toBe('6:00 PM');
    expect(formatNumber(46294, 'aaaa')).toBe('星期二');
    expect(formatNumber(1.5, '[h]:mm:ss')).toBe('36:00:00');
    expect(formatNumber(0.5, 'mm:ss')).toBe('00:00');
    expect(isDateFormat('yyyy/m/d')).toBe(true);
    expect(isDateFormat('0.00E+00')).toBe(false);
    expect(isDateFormat('#,##0.00_);[Red]\\(#,##0.00\\)')).toBe(false);
  });
});

/** 用 writeXlsx 生成骨架，再替换样式与工作表 XML */
function book(sheetXml: string, styles?: string, workbookPr = ''): Uint8Array {
  const files = unzipSync(writeXlsx([{ name: '比价', rows: [['x']] }]));
  files['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet>${sheetXml}</worksheet>`);
  if (styles) files['xl/styles.xml'] = strToU8(styles);
  if (workbookPr) {
    const wb = new TextDecoder().decode(files['xl/workbook.xml']).replace('<sheets>', `${workbookPr}<sheets>`);
    files['xl/workbook.xml'] = strToU8(wb);
  }
  return zipSync(files);
}

const STYLES = `<styleSheet><numFmts count="1"><numFmt numFmtId="176" formatCode="#,##0.00&quot;元&quot;"/></numFmts>
<fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts>
<cellXfs count="4"><xf numFmtId="0" fontId="0"/><xf numFmtId="14" fontId="0"/><xf numFmtId="176" fontId="0"/>
<xf numFmtId="0" fontId="1"><alignment horizontal="center" wrapText="1"/></xf></cellXfs></styleSheet>`;

describe('readWorkbook', () => {
  it('formats values and reads styles', () => {
    const wb = readWorkbook(
      book(
        '<sheetData><row r="1"><c r="A1" s="3" t="inlineStr"><is><t>名称</t></is></c><c r="B1" s="1"><v>46294</v></c><c r="C1" s="2"><v>1234.5</v></c><c r="D1" t="b"><v>1</v></c></row></sheetData>',
        STYLES,
      ),
    );
    const s = wb.sheets[0];
    expect(s.name).toBe('比价');
    expect(s.cells[0].map((c) => c?.t)).toEqual(['名称', '2026/9/29', '1,234.50元', 'TRUE']);
    expect(s.cells[0][1]?.v).toBe(46294);
    expect(wb.styles[s.cells[0][0]!.s!]).toEqual({ bold: true, fontSize: 11, align: 'center', wrap: true });
    expect([s.rows, s.cols]).toEqual([1, 4]);
  });

  it('reads merges, sizes, hidden rows / cols and frozen panes', () => {
    const s = readWorkbook(
      book(
        '<sheetViews><sheetView><pane ySplit="2" xSplit="1" topLeftCell="B3" state="frozen"/></sheetView></sheetViews>' +
          '<sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="20" customWidth="1"/><col min="3" max="4" hidden="1" width="9"/></cols>' +
          '<sheetData><row r="1" ht="30" customHeight="1"><c r="A1" t="inlineStr"><is><t>标题</t></is></c></row><row r="4" hidden="1"><c r="B4"><v>1</v></c></row></sheetData>' +
          '<mergeCells count="1"><mergeCell ref="A1:E2"/></mergeCells>',
      ),
    ).sheets[0];
    expect(s.freeze).toEqual({ rows: 2, cols: 1 });
    expect(s.merges).toEqual([{ r0: 0, c0: 0, r1: 1, c1: 4 }]);
    expect(s.colWidths).toEqual({ 0: 140, 2: 0, 3: 0 });
    expect(s.rowHeights).toEqual({ 0: 40, 3: 0 });
    expect([s.rows, s.cols]).toEqual([4, 5]);
  });

  it('uses the 1904 date system when set', () => {
    const s = readWorkbook(book('<sheetData><row r="1"><c r="A1" s="1"><v>0</v></c></row></sheetData>', STYLES, '<workbookPr date1904="1"/>')).sheets[0];
    expect(s.cells[0][0]?.t).toBe('1904/1/1');
  });

  it('rejects non-zip data', () => {
    expect(() => readWorkbook(new Uint8Array([1, 2, 3]))).toThrow('xlsx');
  });
});

describe('readCsv', () => {
  it('parses quotes, delimiters and numbers', () => {
    const s = readCsv('﻿名称,单价,编号\r\n"钢丝绳, 6mm",12.5,007\r\n"说""明""\n换行",,1e3\r\n').sheets[0];
    expect(s.cells.map((r) => r.map((c) => c?.t))).toEqual([
      ['名称', '单价', '编号'],
      ['钢丝绳, 6mm', '12.5', '007'],
      ['说"明"\n换行', undefined, '1e3'],
    ]);
    expect(s.cells[1][1]?.v).toBe(12.5);
    expect(s.cells[1][2]?.v).toBe('007');
    expect([s.rows, s.cols]).toEqual([3, 3]);
  });

  it('detects tab / semicolon delimiters', () => {
    expect(readCsv('a\tb\n1\t2').sheets[0].cells[1].map((c) => c?.v)).toEqual([1, 2]);
    expect(readCsv('a;b\n1;2').sheets[0].cols).toBe(2);
  });
});

describe('fromXls', () => {
  it('places cells at the range start and formats dates', () => {
    const wb = fromXls([
      { name: '旧表', start: [1, 1], rows: [['名称', 12.5, true], [null, { d: 46294 }, { d: 46294.5 }]], merges: [[0, 0, 0, 3], [5, 5, 5, 5]] },
    ]);
    const s = wb.sheets[0];
    expect(s.cells[1].map((c) => c?.t)).toEqual([undefined, '名称', '12.5', 'TRUE']);
    expect(s.cells[2].map((c) => c?.t)).toEqual([undefined, undefined, '2026/9/29', '2026/9/29 12:00']);
    expect(s.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 3 }]);
    expect([s.rows, s.cols]).toEqual([3, 4]);
    expect(dateText(0.25)).toBe('6:00:00');
  });
});

function sheetOf(rows: (string | number | null)[][], merges: SheetData['merges'] = []): SheetData {
  return {
    name: 'S',
    cells: rows.map((r) => r.map((v) => (v === null ? undefined : { v, t: typeof v === 'number' ? v.toFixed(2) : v }))),
    rows: rows.length,
    cols: Math.max(...rows.map((r) => r.length)),
    merges,
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 72,
    defaultRowHeight: 22,
    freeze: { rows: 0, cols: 0 },
  };
}

describe('grid', () => {
  it('axis lookup', () => {
    const a = buildAxis(4, { 1: 0, 2: 50 }, 20);
    expect(Array.from(a.pos)).toEqual([0, 20, 20, 70, 90]);
    expect([indexAt(a, 0), indexAt(a, 19), indexAt(a, 20), indexAt(a, 69), indexAt(a, 999)]).toEqual([0, 0, 2, 2, 3]);
  });

  it('narrow columns shrink the cell padding so the grid line stays inside the column', () => {
    // 左右内边距 + 1px 网格线 <= 列宽
    expect([64, 13, 12, 7, 2, 1].map(cellPadding)).toEqual([6, 6, 5, 3, 0, 0]);
    for (const w of [1, 2, 7, 12, 13, 64]) expect(2 * cellPadding(w) + 1).toBeLessThanOrEqual(Math.max(1, w));
  });

  it('finds the row / column whose edge is under the pointer', () => {
    // 列 1 隐藏：pos = 0, 20, 20, 70, 90
    const a = buildAxis(4, { 1: 0, 2: 50 }, 20);
    // 第 0 列中间不在分隔线上；20 附近是第 0 列的右边线（第 1 列隐藏，跳过它）
    expect([edgeAt(a, 10), edgeAt(a, 17), edgeAt(a, 20), edgeAt(a, 23), edgeAt(a, 24)]).toEqual([-1, 0, 0, 0, -1]);
    // 70 附近是第 2 列右边线，末尾 90 是最后一列的右边线
    expect([edgeAt(a, 68), edgeAt(a, 72), edgeAt(a, 89), edgeAt(a, 92), edgeAt(a, 999)]).toEqual([2, 2, 3, 3, -1]);
    // 开头的线（第 0 列左边）不能拖
    expect([edgeAt(a, 1), edgeAt(a, -1)]).toEqual([-1, -1]);
    expect(edgeAt(buildAxis(0, {}, 20), 5)).toBe(-1);
  });

  it('parses refs and names ranges', () => {
    expect(parseRef('b3')).toEqual({ r0: 2, c0: 1, r1: 2, c1: 1 });
    expect(parseRef('$C$5:A1')).toEqual({ r0: 0, c0: 0, r1: 4, c1: 2 });
    expect(parseRef('3B')).toBeNull();
    expect(rangeName({ r0: 0, c0: 0, r1: 9, c1: 27 })).toBe('A1:AB10');
  });

  it('expands selections over merges', () => {
    const merges = [{ r0: 0, c0: 1, r1: 1, c1: 2 }, { r0: 1, c0: 3, r1: 3, c1: 3 }];
    expect(expandToMerges({ r0: 1, c0: 0, r1: 1, c1: 1 }, merges)).toEqual({ r0: 0, c0: 0, r1: 1, c1: 2 });
    expect(expandToMerges({ r0: 1, c0: 2, r1: 1, c1: 3 }, merges)).toEqual({ r0: 0, c0: 1, r1: 3, c1: 3 });
  });

  it('moves with arrows, merges, hidden rows and Ctrl-jumps', () => {
    const s = sheetOf(
      [
        ['a', 'b', 'c', null, null, 'f'],
        ['x', null, null, null, null, null],
        [null],
        ['z'],
      ],
      [{ r0: 0, c0: 1, r1: 0, c1: 2 }],
    );
    const ctx = { sheet: s, rows: 10, cols: 8, merges: buildMergeIndex(s.merges), hiddenRow: (r: number) => r === 2, hiddenCol: () => false };
    expect(moveBy(ctx, { r: 0, c: 0 }, 0, 1)).toEqual({ r: 0, c: 1 });
    expect(moveBy(ctx, { r: 0, c: 1 }, 0, 1)).toEqual({ r: 0, c: 3 });
    expect(moveBy(ctx, { r: 1, c: 0 }, 1, 0)).toEqual({ r: 3, c: 0 });
    expect(moveBy(ctx, { r: 0, c: 0 }, 0, 1, true)).toEqual({ r: 0, c: 1 });
    expect(moveBy(ctx, { r: 0, c: 2 }, 0, 1, true)).toEqual({ r: 0, c: 5 });
    expect(moveBy(ctx, { r: 0, c: 5 }, 0, 1, true)).toEqual({ r: 0, c: 7 });
    // 隐藏行不打断数据块
    expect(moveBy(ctx, { r: 0, c: 0 }, 1, 0, true)).toEqual({ r: 3, c: 0 });
    expect(moveBy(ctx, { r: 0, c: 0 }, -1, 0)).toEqual({ r: 0, c: 0 });
  });

  it('copies as TSV / HTML / Markdown, skipping hidden columns', () => {
    const s = sheetOf(
      [
        ['名称', '单价', '备注'],
        ['a\tb', 12.5, 'x|y'],
      ],
      [],
    );
    const m = collect(s, [], { r0: 0, c0: 0, r1: 1, c1: 2 }, { hiddenCol: (c) => c === 2 });
    expect(toTsv(m)).toBe('名称\t单价\r\n"a\tb"\t12.50');
    expect(toTsv(collect(s, [], { r0: 1, c0: 1, r1: 1, c1: 1 }, { raw: true }))).toBe('12.5');
    expect(toHtml(m)).toContain('text-align:right;">12.50</td>');
    expect(toMarkdown(collect(s, [], { r0: 0, c0: 0, r1: 1, c1: 2 }))).toBe('| 名称 | 单价 | 备注 |\n| --- | --- | --- |\n| a\tb | 12.50 | x\\|y |');
  });

  it('keeps merged cells as spans in HTML', () => {
    const s = sheetOf([['标题', null, null], ['a', 'b', 'c']], [{ r0: 0, c0: 0, r1: 0, c1: 2 }]);
    const m = collect(s, [], { r0: 0, c0: 1, r1: 1, c1: 2 });
    expect(m[0][0]).toMatchObject({ text: '标题', colspan: 2 });
    expect(m[0][1]).toBeNull();
    expect(toHtml(m)).toContain('colspan="2"');
  });

  it('computes stats and finds text', () => {
    const s = sheetOf([['名称', 0.1], ['钢丝绳', 0.2], ['钢板', 'x']]);
    expect(statsOf(s, { r0: 0, c0: 0, r1: 99, c1: 9 })).toEqual({ count: 6, nums: 2, sum: 0.3, avg: 0.15, min: 0.1, max: 0.2 });
    expect(statsOf(s, { r0: 0, c0: 1, r1: 2, c1: 1 }, (r) => r === 1).sum).toBe(0.1);
    expect(findAll(s, '钢')).toEqual([{ r: 1, c: 0 }, { r: 2, c: 0 }]);
  });
});

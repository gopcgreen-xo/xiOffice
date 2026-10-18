import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { cellStyleIndex, readXlsx, writeXlsx } from './xlsx';

describe('writeXlsx styles', () => {
  const book = writeXlsx([
    {
      name: '发票',
      rows: [['名称', '金额'], ['a', 1.5], ['b', 2], ['合计', 3.5]],
      widths: [10, 10, 10],
      styles: { 0: 'th', 1: 'tdBand', 2: 'td', 3: 'tf' },
      money: [1],
      center: [0],
      autoFilter: 'A1:C3',
      freezeHeader: true,
    },
  ]);
  const files = unzipSync(book);
  const sheet = strFromU8(files['xl/worksheets/sheet1.xml']);
  const styles = strFromU8(files['xl/styles.xml']);

  it('uses 0.00 (built-in 2) for money columns, combined with the row style', () => {
    expect(sheet).toContain(`<c r="B2" s="${cellStyleIndex('tdBand', true)}"><v>1.5</v></c>`);
    expect(sheet).toContain(`<c r="B3" s="${cellStyleIndex('td', true)}"><v>2</v></c>`);
    expect(sheet).toContain(`<c r="B4" s="${cellStyleIndex('tf', true)}"><v>3.5</v></c>`);
    // 表头是文字，不套数字格式，总是居中
    expect(sheet).toContain(`<c r="B1" s="${cellStyleIndex('th', false, true)}" t="inlineStr">`);
    expect(sheet).toContain(`<c r="A2" s="${cellStyleIndex('tdBand', false, true)}" t="inlineStr">`);
    expect(styles).toContain('numFmtId="2"');
  });

  it('fills the whole width of table rows (empty cells too), adds the filter and freezes the header', () => {
    expect(sheet).toContain(`<c r="C3" s="${cellStyleIndex('td', false)}"/>`);
    expect(sheet).toContain('<autoFilter ref="A1:C3"/>');
    expect(sheet).toContain('state="frozen"');
    expect(sheet.indexOf('<sheetViews>')).toBeLessThan(sheet.indexOf('<cols>'));
  });

  it('declares as many cell formats as it uses, and still reads back as plain values', () => {
    const count = Number(/<cellXfs count="(\d+)"/.exec(styles)![1]);
    expect((styles.match(/<xf [^>]*xfId="0"/g) ?? []).length).toBe(count);
    expect(cellStyleIndex('tf', true, true)).toBe(count - 1);
    expect(cellStyleIndex(undefined, false)).toBe(0);
    expect(readXlsx(book)[0].rows[2]).toEqual(['b', 2]);
  });
});

import { describe, expect, it } from 'vitest';
import { borderFor, buildFormat, fmtParts, formatLabel, NUMBER_FORMATS, stepDecimals, themePalette } from './format';
import { formatNumber } from './numfmt';
import { applyBorders, applyStyle, growRows, lostByMerge, mergeAcross, mergeCells, setInput, Tx } from './ops';
import { saveXlsx } from './save';
import { blankWorkbook, readWorkbook, styleAt, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, '格式');
  fn(tx);
  return tx.commit();
};

describe('数字格式', () => {
  it('增减小数位：常规按当前值，自定义格式每段都改，日期不变', () => {
    expect(stepDecimals('General', 1, 1.5)).toBe('0.00');
    expect(stepDecimals('General', -1, 1.25)).toBe('0.0');
    expect(stepDecimals('General', -1, 3)).toBe('0');
    expect(stepDecimals('0', 1)).toBe('0.0');
    expect(stepDecimals('0.0', -1)).toBe('0');
    expect(stepDecimals('#,##0.00', -1)).toBe('#,##0.0');
    expect(stepDecimals('0%', 1)).toBe('0.0%');
    expect(stepDecimals('"¥"#,##0.00;"¥"-#,##0.00', 1)).toBe('"¥"#,##0.000;"¥"-#,##0.000');
    expect(stepDecimals('0.00"吨"', -1)).toBe('0.0"吨"');
    expect(stepDecimals('yyyy/m/d', 1)).toBe('yyyy/m/d');
  });

  it('格式名称；会计专用的零值显示为短横', () => {
    expect([formatLabel('General'), formatLabel('0.00%'), formatLabel('yyyy-mm-dd'), formatLabel('0.0"吨"')]).toEqual(['常规', '百分比', '日期', '自定义']);
    expect(formatNumber(0, '_ "¥"* #,##0.00_ ;_ "¥"* -#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ ').trim()).toBe('¥-');
  });

  it('调色板：6 档 × 10 色', () => {
    const p = themePalette();
    expect([p.length, p[0].length, p[0][4], p[5][1]]).toEqual([6, 10, '#4472c4', '#808080']);
  });
});

describe('边框', () => {
  const g = { r0: 0, c0: 0, r1: 1, c1: 1 };
  it('外侧框线只改外圈的边；下框线只改最后一行', () => {
    expect(borderFor('outside', g, 0, 0)).toEqual({ top: { style: 'thin' }, left: { style: 'thin' } });
    expect(borderFor('outside', { r0: 0, c0: 0, r1: 2, c1: 2 }, 1, 1)).toBeNull();
    expect(borderFor('bottom', g, 0, 1)).toBeNull();
    expect(borderFor('thickBottom', g, 1, 1, '#ff0000')).toEqual({ bottom: { style: 'medium', color: '#ff0000' } });
  });

  it('设置后显示、保存、再读一致；空格子也能有边框', () => {
    const b = blankWorkbook();
    edit(b, (tx) => {
      setInput(tx, 0, 0, 0, '1');
      applyBorders(tx, 0, g, 'all');
      applyBorders(tx, 0, g, 'thickOutside');
    });
    const s = b.sheets[0];
    const at = (r: number, c: number) => b.styles[styleAt(s, r, c)!].border;
    expect(at(0, 0)).toMatchObject({ l: { w: 2, style: 'solid' }, r: { w: 1, style: 'solid' }, t: { w: 2, style: 'solid' }, b: { w: 1, style: 'solid' } });
    expect(at(1, 1)?.r).toMatchObject({ w: 2, style: 'solid' });
    const again = readWorkbook(saveXlsx(b));
    expect(again.styles[styleAt(again.sheets[0], 1, 1)!].border?.b).toMatchObject({ w: 2, style: 'solid' });
    expect(again.sheets[0].cells[0][0]?.v).toBe(1);
    // 无框线
    edit(b, (tx) => applyBorders(tx, 0, g, 'none'));
    expect(b.styles[styleAt(b.sheets[0], 0, 0)!].border).toBeUndefined();
  });
});

describe('字体、对齐、合并', () => {
  it('字体格式保存往返；字号变大自动加行高，手动行高不动', () => {
    const b = blankWorkbook();
    edit(b, (tx) => {
      setInput(tx, 0, 0, 0, '标题');
      applyStyle(tx, 0, { r0: 0, c0: 0, r1: 0, c1: 0 }, { bold: true, italic: true, underline: true, strike: true, fontName: '黑体', fontSize: 20, color: '#c00000', fill: '#ffff00', align: 'center', valign: 'top', wrap: true });
      growRows(tx, 0, 0, 0, 26);
    });
    const st = b.styles[b.sheets[0].cells[0][0]!.s!];
    expect(st).toEqual({ bold: true, italic: true, underline: true, strike: true, fontName: '黑体', fontSize: 20, color: '#c00000', fill: '#ffff00', align: 'center', valign: 'top', wrap: true });
    expect(b.sheets[0].rowDefs?.[0]).toEqual({ ht: 26 });
    const again = readWorkbook(saveXlsx(b));
    expect(again.styles[again.sheets[0].cells[0][0]!.s!]).toEqual(st);
    expect(again.sheets[0].rowDefs?.[0]?.ht).toBe(26);
    // 取消加粗、去掉底色
    edit(b, (tx) => applyStyle(tx, 0, { r0: 0, c0: 0, r1: 0, c1: 0 }, { bold: false, fill: null }));
    const st2 = b.styles[b.sheets[0].cells[0][0]!.s!];
    expect([st2.bold, st2.fill, st2.italic]).toEqual([undefined, undefined, true]);
  });

  it('合并：左上角以外有内容会提示；跨越合并按行', () => {
    const b = blankWorkbook();
    edit(b, (tx) => {
      setInput(tx, 0, 0, 0, 'a');
      setInput(tx, 0, 1, 1, 'b');
    });
    const g = { r0: 0, c0: 0, r1: 1, c1: 2 };
    expect(lostByMerge(b.sheets[0], g)).toBe(1);
    edit(b, (tx) => mergeAcross(tx, 0, g));
    expect(b.sheets[0].merges).toEqual([
      { r0: 0, c0: 0, r1: 0, c1: 2 },
      { r0: 1, c0: 0, r1: 1, c1: 2 },
    ]);
    expect(b.sheets[0].cells[1]?.[1]).toBeUndefined();
    edit(b, (tx) => mergeCells(tx, 0, g, false));
    expect(b.sheets[0].merges).toEqual([]);
  });
});

describe('设置单元格格式：数字分类', () => {
  it('格式代码 → 分类、小数位、千位分隔', () => {
    expect(fmtParts('General').cat).toBe('general');
    expect(fmtParts('@').cat).toBe('text');
    expect(fmtParts('0.00')).toEqual({ cat: 'number', decimals: 2, thousands: false });
    expect(fmtParts('#,##0')).toEqual({ cat: 'number', decimals: 0, thousands: true });
    expect(fmtParts('"¥"#,##0.00;"¥"-#,##0.00')).toEqual({ cat: 'currency', decimals: 2, thousands: true });
    expect(fmtParts(NUMBER_FORMATS.find((f) => f.label === '会计专用')!.code).cat).toBe('accounting');
    expect(fmtParts('0.0%')).toEqual({ cat: 'percent', decimals: 1, thousands: false });
    expect(fmtParts('0.00E+00').cat).toBe('scientific');
    expect(fmtParts('yyyy/m/d').cat).toBe('date');
    expect(fmtParts('h:mm').cat).toBe('time');
    expect(fmtParts('0.00"吨"').cat).toBe('custom');
  });

  it('分类 → 格式代码，和常用格式下拉里的一致', () => {
    const code = (label: string) => NUMBER_FORMATS.find((f) => f.label === label)!.code;
    expect(buildFormat({ cat: 'number', decimals: 2, thousands: false })).toBe(code('数值'));
    expect(buildFormat({ cat: 'number', decimals: 2, thousands: true })).toBe(code('千分位'));
    expect(buildFormat({ cat: 'currency', decimals: 2, thousands: true })).toBe(code('货币'));
    expect(buildFormat({ cat: 'accounting', decimals: 2, thousands: true })).toBe(code('会计专用'));
    expect(buildFormat({ cat: 'percent', decimals: 2, thousands: false })).toBe(code('百分比'));
    expect(buildFormat({ cat: 'scientific', decimals: 2, thousands: false })).toBe(code('科学记数'));
    expect(buildFormat({ cat: 'number', decimals: 0, thousands: true })).toBe('#,##0');
    expect(buildFormat({ cat: 'date', decimals: 2, thousands: false }, 'yyyy-mm-dd')).toBe('yyyy-mm-dd');
    // 往返：生成的代码再识别回同样的分类和小数位
    for (const cat of ['number', 'currency', 'accounting', 'percent', 'scientific'] as const) {
      for (const decimals of [0, 1, 3]) expect(fmtParts(buildFormat({ cat, decimals, thousands: true }))).toMatchObject({ cat, decimals });
    }
  });
});

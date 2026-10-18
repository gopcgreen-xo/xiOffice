import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { applyStyle, setInput, Tx } from './ops';
import { saveXlsx } from './save';
import { columnExtras, setAutoFilter, setColumnFilter } from './sortFilter';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, 'x');
  fn(tx);
  return tx.commit()!;
};
const hidden = (b: Workbook, n: number) => Array.from({ length: n }, (_, i) => i + 1).filter((r) => b.sheets[0].rowDefs?.[r]?.hidden);

describe('按日期分组、按颜色筛选', () => {
  it('日期列按年 / 月勾选；文字格子仍按值；保存为 dateGroupItem 并读回', () => {
    const b = blankWorkbook();
    const rows = ['日期', '2026-9-30', '2026-10-1', '2026-10-15', '2025-10-1', '待定'];
    edit(b, (tx) => rows.forEach((v, r) => setInput(tx, 0, r, 0, v)));
    const ref = { r0: 0, c0: 0, r1: 5, c1: 0 };
    edit(b, (tx) => setAutoFilter(tx, 0, ref));
    const ex = columnExtras(b, 0, ref, 0);
    expect([...ex.dates.entries()].map(([y, ms]) => [y, [...ms.entries()]])).toEqual([
      [2026, [[9, 1], [10, 2]]],
      [2025, [[10, 1]]],
    ]);
    edit(b, (tx) => setColumnFilter(tx, 0, 0, { values: ['待定'], dates: [{ y: 2026, m: 10 }] }));
    expect(hidden(b, 5)).toEqual([1, 4]);
    const xml = strFromU8(unzipSync(saveXlsx(b))['xl/worksheets/sheet1.xml']);
    expect(xml).toContain('<filters><filter val="待定"/><dateGroupItem year="2026" month="10" dateTimeGrouping="month"/></filters>');
    const again = readWorkbook(saveXlsx(b));
    expect(again.sheets[0].autoFilter?.cols[0]).toEqual({ values: ['待定'], dates: [{ y: 2026, m: 10 }] });
  });

  it('按底色：只留下黄色的行；保存时追加差异格式，colorFilter 引用它', () => {
    const b = blankWorkbook();
    edit(b, (tx) => ['名称', 'a', 'b', 'c'].forEach((v, r) => setInput(tx, 0, r, 0, v)));
    edit(b, (tx) => applyStyle(tx, 0, { r0: 1, c0: 0, r1: 1, c1: 0 }, { fill: '#ffff00' }));
    edit(b, (tx) => applyStyle(tx, 0, { r0: 3, c0: 0, r1: 3, c1: 0 }, { fill: '#ffff00' }));
    const ref = { r0: 0, c0: 0, r1: 3, c1: 0 };
    edit(b, (tx) => setAutoFilter(tx, 0, ref));
    expect(columnExtras(b, 0, ref, 0).fills.get('#ffff00')).toBe(2);
    edit(b, (tx) => setColumnFilter(tx, 0, 0, { color: { cell: true, color: '#ffff00' } }));
    expect(hidden(b, 3)).toEqual([2]);
    const z = unzipSync(saveXlsx(b));
    const xml = strFromU8(z['xl/worksheets/sheet1.xml']);
    const id = xml.match(/<colorFilter dxfId="(\d+)"\/>/)?.[1];
    expect(id).toBeDefined();
    const styles = strFromU8(z['xl/styles.xml']);
    expect(styles).toMatch(/<dxfs count="\d+">.*<bgColor rgb="FFFFFF00"\/>.*<\/dxfs>/);
    const again = readWorkbook(saveXlsx(b));
    expect(again.sheets[0].autoFilter?.cols[0]).toEqual({ color: { cell: true, color: '#ffff00', dxf: Number(id) } });
  });
});

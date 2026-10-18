import { describe, expect, it } from 'vitest';
import { editText, findAll, findNext, replaceAll, replaceCell, replaceText } from './find';
import { layoutPrint, usedRange } from './print';
import { setInput, setColWidth, setRowHeight, Tx } from './ops';
import { saveXlsx } from './save';
import { blankWorkbook, DEFAULT_MARGINS, readWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, 'x');
  fn(tx);
  return tx.commit();
};

describe('查找替换', () => {
  function sample() {
    const b = blankWorkbook();
    b.sheets.push({ ...structuredClone(b.sheets[0]), name: 'Sheet2' });
    edit(b, (tx) => {
      setInput(tx, 0, 0, 0, '钢板 Q235');
      setInput(tx, 0, 1, 0, '钢板');
      setInput(tx, 0, 1, 1, '=LEN(A1)');
      setInput(tx, 0, 2, 0, '12%');
      setInput(tx, 1, 0, 0, 'q235 钢板');
    });
    return b;
  }

  it('原文：公式带 =，百分比带 %，像数字的文字带 \'', () => {
    const b = sample();
    edit(b, (tx) => setInput(tx, 0, 3, 0, "'123"));
    const s = b.sheets[0];
    expect([editText(b, s.cells[1][1]), editText(b, s.cells[2][0]), editText(b, s.cells[3][0])]).toEqual(['=LEN(A1)', '12%', "'123"]);
  });

  it('查找：范围、大小写、完全匹配、公式 / 值', () => {
    const b = sample();
    expect(findAll(b, { query: 'q235' }).length).toBe(2);
    expect(findAll(b, { query: 'q235', matchCase: true })).toEqual([{ si: 1, r: 0, c: 0 }]);
    expect(findAll(b, { query: '钢板', whole: true, si: 0 })).toEqual([{ si: 0, r: 1, c: 0 }]);
    expect(findAll(b, { query: 'LEN' }).length).toBe(1);
    expect(findAll(b, { query: 'LEN', lookIn: 'values' }).length).toBe(0);
    const hits = findAll(b, { query: '钢板' });
    expect(findNext(hits, { si: 0, r: 1, c: 0 })).toEqual({ si: 1, r: 0, c: 0 });
    expect(findNext(hits, { si: 1, r: 0, c: 0 })).toEqual({ si: 0, r: 0, c: 0 });
    expect(findNext(hits, { si: 0, r: 0, c: 0 }, true)).toEqual({ si: 1, r: 0, c: 0 });
  });

  it('替换：不区分大小写按原位置换；替换后重新识别（数字变数字）；公式里的文字也换', () => {
    expect(replaceText('Q235 q235', { query: 'q235' }, 'Q355')).toBe('Q355 Q355');
    const b = sample();
    edit(b, (tx) => {
      expect(replaceAll(tx, { query: 'q235' }, 'Q355')).toBe(2);
      replaceCell(tx, { si: 0, r: 2, c: 0 }, { query: '12' }, '15');
      replaceCell(tx, { si: 0, r: 1, c: 1 }, { query: 'A1' }, 'A2');
    });
    expect(b.sheets[0].cells[0][0]?.v).toBe('钢板 Q355');
    expect(b.sheets[1].cells[0][0]?.v).toBe('Q355 钢板');
    expect(b.sheets[0].cells[2][0]).toMatchObject({ v: 0.15, t: '15%' });
    expect(b.sheets[0].cells[1][1]?.f).toBe('LEN(A2)');
  });
});

describe('打印分页', () => {
  function rowsOf(n: number) {
    const b = blankWorkbook();
    edit(b, (tx) => {
      for (let r = 0; r < n; r++) for (let c = 0; c < 3; c++) setInput(tx, 0, r, c, String(r));
    });
    return b;
  }

  it('A4 纵向：行高 22px（16.5 磅）一页约 44 行；先向下再向右；每页重复标题行', () => {
    const b = rowsOf(100);
    const s = b.sheets[0];
    expect(usedRange(s)).toEqual({ r0: 0, c0: 0, r1: 99, c1: 2 });
    const l = layoutPrint(s);
    expect(l.pages.length).toBe(3);
    expect(l.pages[0].rows.length).toBe(44);
    expect(l.pages[1].rows[0]).toBe(44);
    const t = layoutPrint(s, { titleRows: [0, 0] });
    expect(t.pages[1].rows[0]).toBe(0);
    expect(t.pages[1].rows[1]).toBe(44);
    expect(t.pages.flatMap((p) => p.rows).filter((r) => r !== 0).length).toBe(99);
  });

  it('宽表横向分页；调整为一页宽时缩放；隐藏的行列不打印', () => {
    const b = rowsOf(10);
    edit(b, (tx) => {
      setColWidth(tx, 0, 0, 2, { width: 400 });
      setRowHeight(tx, 0, 3, 3, { hidden: true });
    });
    const s = b.sheets[0];
    const l = layoutPrint(s);
    expect(l.pages.map((p) => p.cols)).toEqual([[0], [1], [2]]);
    expect(l.pages[0].rows).not.toContain(3);
    const fit = layoutPrint(s, { fitW: 1, fitH: 0 });
    expect(fit.pages.length).toBe(1);
    expect(fit.scale).toBeCloseTo((595.3 - 100.8) / 900, 3);
    expect(layoutPrint(s, { landscape: true }).w).toBeCloseTo(841.9);
  });

  it('页面设置保存往返', () => {
    const b = rowsOf(3);
    b.sheets[0].page = { landscape: true, paper: 8, fitW: 1, fitH: 0, gridLines: true, titleRows: [0, 0], area: { r0: 0, c0: 0, r1: 2, c1: 1 } };
    b.sheets[0].dirty = true;
    const again = readWorkbook(saveXlsx(b));
    expect(again.sheets[0].page).toEqual({ ...b.sheets[0].page, margins: DEFAULT_MARGINS });
  });
});

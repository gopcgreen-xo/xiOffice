import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { headerFooterText, layoutPrint } from './print';
import { insertDelete, setInput, Tx } from './ops';
import { blankPackage, saveXlsx } from './save';
import { blankWorkbook, readWorkbook } from './workbook';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const vars = { page: 2, pages: 5, sheet: '报价', file: '报价单.xlsx', date: new Date(2026, 9, 1, 8, 5) };

describe('页面设置：页边距、居中、页眉页脚、分页符', () => {
  it('页眉页脚代码：分节、变量、格式代码去掉、&& 是 &', () => {
    expect(headerFooterText('&L&"宋体,加粗"&12&A&C第 &P 页，共 &N 页&R&D &T', vars)).toEqual({ left: '报价', center: '第 2 页，共 5 页', right: '2026/10/1 8:05' });
    expect(headerFooterText('A&&B &F &P+1', vars)).toEqual({ left: '', center: 'A&B 报价单.xlsx 3', right: '' });
  });

  it('读写：pageMargins、printOptions 居中、oddHeader / oddFooter、手动分页符；首页页眉保留', () => {
    const files = blankPackage(['S1']);
    files['xl/worksheets/sheet1.xml'] = strToU8(
      `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
        `<printOptions horizontalCentered="1"/><pageMargins left="0.5" right="0.5" top="1" bottom="1" header="0.4" footer="0.4"/>` +
        `<headerFooter differentFirst="1"><oddFooter>&amp;C第 &amp;P 页</oddFooter><firstHeader>&amp;C首页</firstHeader></headerFooter>` +
        `<rowBreaks count="2" manualBreakCount="1"><brk id="10" max="16383" man="1"/><brk id="20" max="16383"/></rowBreaks></worksheet>`,
    );
    const data = zipSync(files);
    const b = readWorkbook(data);
    b.source = data;
    const pg = b.sheets[0].page!;
    expect(pg).toMatchObject({ centerH: true, footer: '&C第 &P 页', rowBreaks: [10], margins: { left: 0.5, top: 1, header: 0.4 } });
    // 改：页眉、垂直居中、插入两行（分页符下移）
    const tx = new Tx(b, 'x');
    tx.meta(0);
    b.sheets[0].page = { ...pg, header: '&L&A', centerV: true };
    insertDelete(tx, 0, 'row', 0, 2);
    tx.commit();
    expect(b.sheets[0].page!.rowBreaks).toEqual([12]);
    const xml = strFromU8(unzipSync(saveXlsx(b))['xl/worksheets/sheet1.xml']);
    expect(xml).toContain('<printOptions horizontalCentered="1" verticalCentered="1"/>');
    expect(xml).toContain('<pageMargins left="0.5" right="0.5" top="1" bottom="1" header="0.4" footer="0.4"/>');
    expect(xml).toContain('<headerFooter differentFirst="1"><oddHeader>&amp;L&amp;A</oddHeader><oddFooter>&amp;C第 &amp;P 页</oddFooter><firstHeader>&amp;C首页</firstHeader></headerFooter>');
    expect(xml).toContain('<rowBreaks count="1" manualBreakCount="1"><brk id="12" max="16383" man="1"/></rowBreaks>');
    expect(xml.indexOf('<pageMargins')).toBeLessThan(xml.indexOf('<headerFooter'));
  });

  it('分页：手动分页符处换页；页边距变大每页行数变少', () => {
    const b = blankWorkbook();
    const tx = new Tx(b, 'x');
    for (let r = 0; r < 30; r++) setInput(tx, 0, r, 0, String(r));
    tx.commit();
    const s = b.sheets[0];
    const plain = layoutPrint(s, {});
    expect(plain.pages.length).toBe(1);
    const brk = layoutPrint(s, { rowBreaks: [10, 25] });
    expect(brk.pages.map((p) => [p.rows[0], p.rows[p.rows.length - 1]])).toEqual([[0, 9], [10, 24], [25, 29]]);
    const wide = layoutPrint(s, { margins: { left: 0.7, right: 0.7, top: 5, bottom: 5, header: 0.3, footer: 0.3 } });
    expect(wide.pages.length).toBeGreaterThan(1);
  });
});

import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import { join } from 'path';
import { strFromU8, unzipSync } from 'fflate';
import { insertDelete, Tx } from './ops';
import { saveXlsx } from './save';
import { readWorkbook } from './workbook';

const data = () => new Uint8Array(readFileSync(join(__dirname, '../../samples/报价单样例.xlsx')));

describe('图片、图表、批注、超链接', () => {
  it('读出样本里的图表：类型、系列引用和缓存', () => {
    const b = readWorkbook(data());
    const d = b.sheets[0].drawings!;
    expect(d.length).toBe(1);
    expect(d[0]).toMatchObject({ kind: 'chart', from: { c: 6, r: 2 }, to: { c: 12, r: 16 }, name: '图表 1' });
    expect(d[0].chart).toMatchObject({ type: 'col', series: [{ catRef: '报价!$A$3:$A$4', valRef: '报价!$D$3:$D$4' }] });
  });

  it('插入行后保存：图表锚点跟着下移', () => {
    const src = data();
    const b = readWorkbook(src);
    b.source = src;
    const tx = new Tx(b, '插入');
    insertDelete(tx, 0, 'row', 1, 2);
    tx.commit();
    expect(b.sheets[0].drawings![0].from.r).toBe(4);
    const z = unzipSync(saveXlsx(b));
    const drawing = strFromU8(z['xl/drawings/drawing1.xml']);
    expect(drawing).toContain('<xdr:row>4</xdr:row>');
    expect(drawing).toContain('<xdr:row>18</xdr:row>');
  });
});

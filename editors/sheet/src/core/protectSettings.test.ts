import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { ALLOW_KEYS, allowOf, DEFAULT_ALLOW, hasPassword, isLocked, legacyHash, protectSheetAttrs, setCellsLocked, setSheetProtection, setStructureLock, structureLocked, verifyPassword } from './protect';
import { applyEntry, setInput, Tx, violatesProtection } from './ops';
import { sha384, sha512 } from './sha512';
import { saveXlsx } from './save';
import { blankWorkbook, readWorkbook, type Workbook } from './workbook';

const edit = (b: Workbook, fn: (tx: Tx) => void) => {
  const tx = new Tx(b, '测试');
  fn(tx);
  return tx.commit();
};
const sheetXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);

describe('SHA-512（同步实现）', () => {
  it('和 WebCrypto 的结果一致：各种长度（含跨块、空）', async () => {
    for (const n of [0, 1, 3, 55, 56, 111, 112, 113, 127, 128, 129, 255, 300]) {
      const data = Uint8Array.from({ length: n }, (_, i) => (i * 31 + 7) & 0xff);
      expect(Buffer.from(sha512(data)).toString('hex')).toBe(Buffer.from(await crypto.subtle.digest('SHA-512', data)).toString('hex'));
      expect(Buffer.from(sha384(data)).toString('hex')).toBe(Buffer.from(await crypto.subtle.digest('SHA-384', data)).toString('hex'));
    }
  });
});

describe('密码', () => {
  it('旧式散列：已知例子', () => {
    expect(legacyHash('')).toBe('');
    expect(legacyHash('password')).toBe('83AF');
    expect(legacyHash('test')).toBe('CBEB');
  });

  it('校验 Excel 自己写的新式（SHA-512 加盐迭代）密码', async () => {
    // 这份属性是 Microsoft Excel 16 保护工作表（密码 abc123）后写出来的
    const attrs = {
      algorithmName: 'SHA-512',
      hashValue: 'HeJWnbigrdV7X2+1poWdD2gHqafcuWWAd9XH/fkWYadhZHTTtVH+cAgK2jnG84f5dAEkYn0EgFw8PskOWqrI2w==',
      saltValue: 'TdOvoWmqTKMhKvU37kW/+w==',
      spinCount: '100000',
      sheet: '1',
    };
    expect(await verifyPassword(attrs, 'abc123')).toBe(true);
    expect(await verifyPassword(attrs, 'abc124')).toBe(false);
    expect(await verifyPassword(attrs, '')).toBe(false);
    expect(hasPassword(attrs)).toBe(true);
  }, 30000);

  it('旧式和没有密码；不认识的算法无法校验', async () => {
    expect(await verifyPassword({ sheet: '1', password: '83AF' }, 'password')).toBe(true);
    expect(await verifyPassword({ sheet: '1', password: '83AF' }, 'Password')).toBe(false);
    expect(await verifyPassword({ sheet: '1' }, '')).toBe(true);
    expect(hasPassword({ sheet: '1' })).toBe(false);
    expect(await verifyPassword({ algorithmName: 'MD5', hashValue: 'x', saltValue: 'AA==', spinCount: '1' }, 'a')).toBeNull();
  });
});

describe('保护选项', () => {
  it('默认：只允许选定单元格；属性缺省值同 Excel', () => {
    expect(DEFAULT_ALLOW.selectLockedCells).toBe(true);
    expect(DEFAULT_ALLOW.selectUnlockedCells).toBe(true);
    for (const k of ALLOW_KEYS.filter((x) => !x.startsWith('select') && x !== 'objects')) expect(DEFAULT_ALLOW[k]).toBe(false);
    expect(protectSheetAttrs(undefined, DEFAULT_ALLOW, '')).toEqual({ sheet: '1', objects: '1', scenarios: '1' });
    expect(protectSheetAttrs(undefined, DEFAULT_ALLOW, 'abc')).toMatchObject({ password: legacyHash('abc'), sheet: '1' });
  });

  it('勾选允许的操作写成 0，取消勾选选定单元格写成 1；读回来一致', () => {
    const attrs = protectSheetAttrs(undefined, { ...DEFAULT_ALLOW, formatCells: true, sort: true, selectLockedCells: false }, '');
    expect(attrs).toEqual({ sheet: '1', objects: '1', scenarios: '1', selectLockedCells: '1', formatCells: '0', sort: '0' });
    const back = allowOf(attrs);
    expect([back.formatCells, back.sort, back.selectLockedCells, back.insertRows]).toEqual([true, true, false, false]);
  });
});

describe('保护 / 取消保护', () => {
  it('保护后只能改未锁定的格子；保护本身可撤销，不受自己限制', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setCellsLocked(tx, 0, { r0: 1, c0: 1, r1: 1, c1: 1 }, false));
    expect(isLocked(b, b.sheets[0], 1, 1)).toBe(false);
    expect(isLocked(b, b.sheets[0], 0, 0)).toBe(true);
    const on = edit(b, (tx) => setSheetProtection(tx, 0, protectSheetAttrs(undefined, DEFAULT_ALLOW, 'pw')))!;
    expect(b.sheets[0].protected).toBe(true);
    expect(violatesProtection(b, on)).toBe(false);
    const tx1 = new Tx(b, '改锁定格');
    setInput(tx1, 0, 1, 1, '1');
    expect(violatesProtection(b, tx1.commit()!)).toBe(false);
    const tx2 = new Tx(b, '改锁定的格');
    setInput(tx2, 0, 0, 0, '1');
    expect(violatesProtection(b, tx2.commit()!)).toBe(true);
    const off = edit(b, (tx) => setSheetProtection(tx, 0, null))!;
    expect(violatesProtection(b, off)).toBe(false);
    expect(b.sheets[0].protected).toBeUndefined();
    applyEntry(b, off, -1);
    expect(b.sheets[0].protected).toBe(true);
    expect(b.sheets[0].protection?.attrs.password).toBe(legacyHash('pw'));
  });

  it('保存：sheetProtection 写在 sheetData 之后、重新打开一致；取消保护后元素消失；锁定状态保存在样式里', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setInput(tx, 0, 0, 0, 'x'));
    edit(b, (tx) => setCellsLocked(tx, 0, { r0: 0, c0: 1, r1: 0, c1: 1 }, false));
    edit(b, (tx) => setSheetProtection(tx, 0, protectSheetAttrs(undefined, { ...DEFAULT_ALLOW, sort: true }, 'abc')));
    const out = saveXlsx(b);
    const xml = sheetXml(out);
    expect(xml).toContain(`<sheetProtection password="${legacyHash('abc')}" sheet="1" objects="1" scenarios="1" sort="0"/>`);
    expect(xml.indexOf('</sheetData>')).toBeLessThan(xml.indexOf('<sheetProtection'));
    const styles = strFromU8(unzipSync(out)['xl/styles.xml']);
    expect(styles).toMatch(/<protection locked="0"\/>/);
    expect(styles).toMatch(/applyProtection="1"/);
    const back = readWorkbook(out);
    back.source = out;
    expect(back.sheets[0].protected).toBe(true);
    expect(back.sheets[0].protection?.attrs).toEqual({ password: legacyHash('abc'), sheet: '1', objects: '1', scenarios: '1', sort: '0' });
    edit(back, (tx) => setSheetProtection(tx, 0, null));
    expect(sheetXml(saveXlsx(back))).not.toContain('sheetProtection');
  });

  it('没改保护的工作表保存不动它的保护属性（Excel 写的新式密码原样保留）', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setSheetProtection(tx, 0, { algorithmName: 'SHA-512', hashValue: 'AAAA', saltValue: 'BBBB', spinCount: '100000', sheet: '1', objects: '1', scenarios: '1' }));
    const out = saveXlsx(b);
    const back = readWorkbook(out);
    back.source = out;
    edit(back, (tx) => setInput(tx, 0, 0, 0, 'edit'));
    expect(sheetXml(saveXlsx(back))).toContain('<sheetProtection algorithmName="SHA-512" hashValue="AAAA" saltValue="BBBB" spinCount="100000" sheet="1" objects="1" scenarios="1"/>');
  });
});

describe('工作簿结构保护', () => {
  it('锁定 / 解锁；写进 workbook.xml 的位置在 bookViews 之前', () => {
    const b = blankWorkbook();
    expect(structureLocked(b)).toBe(false);
    const e = edit(b, (tx) => setStructureLock(tx, { password: 'pw' }))!;
    expect(structureLocked(b)).toBe(true);
    const xml = strFromU8(unzipSync(saveXlsx(b))['xl/workbook.xml']);
    expect(xml).toContain(`<workbookProtection workbookPassword="${legacyHash('pw')}" lockStructure="1"/>`);
    expect(xml.indexOf('<workbookProtection')).toBeLessThan(xml.indexOf('<bookViews'));
    applyEntry(b, e, -1);
    expect(structureLocked(b)).toBe(false);
    const back = readWorkbook(saveXlsx(b));
    expect(structureLocked(back)).toBe(false);
  });

  it('读回来；解除后文件里没有；只改了这一项也会保存', () => {
    const b = blankWorkbook();
    edit(b, (tx) => setStructureLock(tx, { password: '' }));
    const out = saveXlsx(b);
    const back = readWorkbook(out);
    back.source = out;
    expect(back.wbProtection?.attrs).toEqual({ lockStructure: '1' });
    edit(back, (tx) => setStructureLock(tx, false));
    expect(strFromU8(unzipSync(saveXlsx(back))['xl/workbook.xml'])).not.toContain('workbookProtection');
  });
});

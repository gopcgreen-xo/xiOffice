/**
 * 插入 / 删除行后保存：工作表关联的二进制部件（打印机设置 printerSettings*.bin）必须逐字节保留。
 * 以前 shiftParts 把所有关联部件都当 UTF-8 文字解码再写回，非 UTF-8 字节变成替换字符，Excel 打不开文件。
 */
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { insertDelete, Tx } from './ops';
import { saveXlsx } from './save';
import { readWorkbook } from './workbook';

const xml = (s: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`);
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function book(bin: Uint8Array): Uint8Array {
  return zipSync({
    '[Content_Types].xml': xml(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    ),
    '_rels/.rels': xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    'xl/workbook.xml': xml(`<workbook ${NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    'xl/worksheets/sheet1.xml': xml(
      `<worksheet ${NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><f>SUM(A1:A1)</f><v>1</v></c></row></sheetData><pageSetup paperSize="9" r:id="rId1"/></worksheet>`,
    ),
    'xl/worksheets/_rels/sheet1.xml.rels': xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/printerSettings" Target="../printerSettings/printerSettings1.bin"/></Relationships>`),
    'xl/printerSettings/printerSettings1.bin': bin,
  });
}

describe('二进制关联部件', () => {
  it('历史 UTF-8 损坏的打印设置只断开引用，数据、打印纸张和其它部件保持不变', () => {
    const binary = new Uint8Array(224);
    const header = new DataView(binary.buffer);
    header.setUint16(64, 0x0401, true);
    header.setUint16(68, 220, true);
    header.setUint16(70, 4, true);
    binary[80] = 0x9a;
    const damaged = strToU8(strFromU8(binary));
    const source = book(damaged);
    const b = readWorkbook(source);
    b.source = source;
    const before = unzipSync(source), after = unzipSync(saveXlsx(b));
    const path = 'xl/worksheets/sheet1.xml';
    expect(strFromU8(after[path])).toBe(strFromU8(before[path]).replace('paperSize="9" r:id="rId1"', 'paperSize="9"'));
    for (const name of Object.keys(before).filter(name => name !== path)) expect(after[name]).toEqual(before[name]);
    const healthy = readWorkbook(book(binary));
    healthy.source = book(binary);
    expect(unzipSync(saveXlsx(healthy))['xl/worksheets/sheet1.xml']).toEqual(unzipSync(healthy.source)['xl/worksheets/sheet1.xml']);
  });
  it('插入行后保存，printerSettings.bin 逐字节不变', () => {
    // 0xDC、0x9A 这类字节单独出现时不是合法 UTF-8
    const bin = new Uint8Array([0x00, 0x01, 0xdc, 0x00, 0x58, 0x9a, 0x0b, 0xff, 0xfe, 0x80]);
    const b = readWorkbook(book(bin));
    b.source = book(bin);
    const tx = new Tx(b, '插入行');
    insertDelete(tx, 0, 'row', 1, 3);
    tx.commit();
    const out = unzipSync(saveXlsx(b));
    expect(out['xl/printerSettings/printerSettings1.bin']).toEqual(bin);
  });
});

import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { insertDelete, setInput, Tx, violatesProtection } from './ops';
import { blankPackage } from './save';
import { readWorkbook } from './workbook';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

describe('工作表保护', () => {
  it('受保护的表：未锁定（protection locked="0"）的格子能改，锁定的、结构改动都不行', () => {
    const files = blankPackage(['S1']);
    files['xl/styles.xml'] = strToU8(
      `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0"/></xf></cellXfs></styleSheet>`,
    );
    files['xl/worksheets/sheet1.xml'] = strToU8(
      `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"/></row></sheetData><sheetProtection sheet="1" objects="1" scenarios="1"/></worksheet>`,
    );
    const b = readWorkbook(zipSync(files));
    expect(b.sheets[0].protected).toBe(true);
    const edit = (fn: (tx: Tx) => void) => {
      const tx = new Tx(b, 'x');
      fn(tx);
      return violatesProtection(b, tx.commit()!);
    };
    expect(edit((tx) => setInput(tx, 0, 0, 1, '5'))).toBe(false); // B1 未锁定
    expect(edit((tx) => setInput(tx, 0, 0, 0, '5'))).toBe(true); // A1 锁定
    expect(edit((tx) => setInput(tx, 0, 3, 3, '5'))).toBe(true); // 空格子缺省锁定
    expect(edit((tx) => insertDelete(tx, 0, 'row', 0, 1))).toBe(true);
  });
});

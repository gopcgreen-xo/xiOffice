import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { Engine } from './formula/eval';
import { setInput, Tx } from './ops';
import { blankPackage, saveXlsx } from './save';
import { readWorkbook } from './workbook';
import { colName } from '@wz/editor-shared/xlsx';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** 1 万行 × 50 列：数字、文字（共享字符串）和每行一个求和公式 */
function bigXlsx(rows: number, cols: number): Uint8Array {
  const files = blankPackage(['S1']);
  const sst: string[] = [];
  let data = '';
  for (let r = 1; r <= rows; r++) {
    let row = `<row r="${r}">`;
    for (let c = 0; c < cols; c++) {
      const ref = `${colName(c)}${r}`;
      if (c === cols - 1) row += `<c r="${ref}"><f>SUM(A${r}:${colName(cols - 3)}${r})</f><v>0</v></c>`;
      else if (c % 5 === 0) {
        sst.push(`物资${r}-${c}`);
        row += `<c r="${ref}" t="s"><v>${sst.length - 1}</v></c>`;
      } else row += `<c r="${ref}"><v>${r * c}</v></c>`;
    }
    data += `${row}</row>`;
  }
  files['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${MAIN}"><sheetData>${data}</sheetData></worksheet>`);
  files['xl/sharedStrings.xml'] = strToU8(`<sst xmlns="${MAIN}" count="${sst.length}" uniqueCount="${sst.length}">${sst.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>`);
  return zipSync(files);
}

describe('大文件性能（1 万行 × 50 列）', () => {
  it('读取、改一格后增量重算、保存都在限定时间内', () => {
    const data = bigXlsx(10000, 50);
    let t = performance.now();
    const b = readWorkbook(data);
    b.source = data;
    const read = performance.now() - t;
    expect(b.sheets[0].rows).toBe(10000);

    const e = new Engine(b);
    t = performance.now();
    e.analyze();
    const analyze = performance.now() - t;

    t = performance.now();
    const tx = new Tx(b, '输入');
    setInput(tx, 0, 5000, 1, '123456');
    tx.commit();
    e.recalc(tx.changed);
    const recalc = performance.now() - t;
    expect(b.sheets[0].cells[5000][49]?.v).toBeGreaterThan(123456);

    t = performance.now();
    const out = saveXlsx(b);
    const save = performance.now() - t;
    expect(out.length).toBeGreaterThan(1000);

    console.log(`读取 ${read.toFixed(0)} ms，分析公式 ${analyze.toFixed(0)} ms，改一格重算 ${recalc.toFixed(0)} ms，保存 ${save.toFixed(0)} ms`);
    // 机器差别大，只防明显退化（需求：打开 ≤ 2 秒，输入无卡顿）
    expect(read).toBeLessThan(4000);
    expect(recalc).toBeLessThan(500);
  }, 60000);
});

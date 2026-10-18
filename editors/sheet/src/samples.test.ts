// @vitest-environment jsdom
/**
 * 回归样本的往返测试：本编辑器 samples/ 下的每个文件打开 → 修改 → 保存 → 再打开，内容不能丢、不能变。
 * 真实文件（脱敏后）放进这个目录就会自动参加测试；在 Office 里打开的人工核对步骤见 editors/SAMPLES.md。
 */
import { readdirSync, readFileSync, statSync } from 'fs';
import { join } from 'path';
import { expect, describe, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { readWorkbook } from './core/workbook';
import { saveXlsx, markSaved } from './core/save';
import { setInput, Tx } from './core/ops';

const DIR = join(__dirname, '../samples');

function files(dir: string): string[] {
  let out: string[] = [];
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) out = out.concat(files(p));
    else if (/\.(xlsx|xlsm)$/i.test(name) && !name.startsWith('~$')) out.push(p);
  }
  return out;
}

const xlsx = files(DIR);

describe.runIf(xlsx.length)('Excel 样本往返', () => {
  it.each(xlsx)('%s', (file) => {
    const data = new Uint8Array(readFileSync(file));
    const book = readWorkbook(data);
    book.source = data;
    const values = (b: typeof book) => b.sheets.map((s) => s.cells.map((row) => (row ?? []).map((c) => (c ? `${c.f ?? ''}|${c.v}` : ''))));
    const before = values(book);
    // 在第一张可见表的已用区域外写一个值，保存后再读：只多这一格
    const si = book.sheets.findIndex((s) => !s.hidden);
    const r = book.sheets[si].rows + 1;
    const tx = new Tx(book, '样本');
    setInput(tx, si, r, 0, '回归测试');
    tx.commit();
    const out = saveXlsx(book);
    markSaved(book, out);
    const again = readWorkbook(out);
    expect(again.sheets[si].cells[r]?.[0]?.v).toBe('回归测试');
    const after = values(again);
    after[si][r] = before[si][r] ?? [];
    expect(after.map((s) => s.filter((row) => row.some(Boolean)))).toEqual(before.map((s) => s.filter((row) => row.some(Boolean))));
    // 没改的部件逐字节保留（工作表、共享字符串、样式、工作簿、关系、内容类型、calcChain 之外的）
    const a = unzipSync(data);
    const z = unzipSync(out);
    for (const [name, bytes] of Object.entries(a)) {
      if (/^xl\/(worksheets\/sheet|sharedStrings|styles|workbook\.xml|_rels\/workbook|calcChain)|\[Content_Types\]/.test(name)) continue;
      expect(strFromU8(z[name] ?? new Uint8Array()), name).toBe(strFromU8(bytes));
    }
  });
});

it('样本目录存在', () => {
  expect(statSync(DIR).isDirectory()).toBe(true);
});

/**
 * 生成给 Excel 验收的 xlsx（`scripts/excel-check.ps1` 打开它们）：
 *   npx vite-node editors/sheet/tools/make-check-files.ts <输出目录>
 * 每个文件覆盖一类功能，随功能增加往 CASES 里加。
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { Engine } from '../src/core/formula/eval';
import { saveXlsx } from '../src/core/save';
import { setInput, Tx } from '../src/core/ops';
import { blankWorkbook, readWorkbook, type Workbook } from '../src/core/workbook';
import { saveName } from '../src/core/names';
import { setNote } from '../src/core/notes';
import { addChart, deleteDrawing, guessSource, moveDrawing, seriesFromRange, updateChart } from '../src/core/charts';
import { DEFAULT_ALLOW, protectSheetAttrs, setCellsLocked, setSheetProtection, setStructureLock } from '../src/core/protect';

const out = process.argv[2];
if (!out) throw new Error('用法：make-check-files.ts <输出目录>');
mkdirSync(out, { recursive: true });

const edit = (b: Workbook, fn: (tx: Tx) => unknown) => {
  const tx = new Tx(b, 'x');
  fn(tx);
  tx.commit();
};
const fill = (b: Workbook, rows: (string | number)[][]) =>
  edit(b, (tx) => rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, typeof v === 'string' ? `'${v}` : String(v)))));

const CASES: Record<string, () => Workbook> = {
  '保护和名称': () => {
    const b = blankWorkbook();
    fill(b, [['项目', '金额'], ['甲', 10], ['乙', 20]]);
    edit(b, (tx) => setInput(tx, 0, 3, 1, '=SUM(Amounts)*Rate'));
    edit(b, (tx) => saveName(tx, { name: 'Amounts', f: 'Sheet1!$B$2:$B$3' }));
    edit(b, (tx) => saveName(tx, { name: 'Rate', f: '1.5' }));
    edit(b, (tx) => saveName(tx, { name: '本表名称', f: 'Sheet1!$A$1', local: 0 }));
    edit(b, (tx) => setCellsLocked(tx, 0, { r0: 1, c0: 1, r1: 2, c1: 1 }, false));
    edit(b, (tx) => setSheetProtection(tx, 0, protectSheetAttrs(undefined, { ...DEFAULT_ALLOW, sort: true }, 'abc123')));
    edit(b, (tx) => setStructureLock(tx, { password: 'book1' }));
    new Engine(b).analyze();
    return b;
  },
  '批注': () => {
    const b = blankWorkbook();
    fill(b, [['项目', '金额'], ['甲', 10], ['乙', 20]]);
    edit(b, (tx) => setNote(tx, 0, 0, 0, '这里是标题 & <说明>', '小王'));
    edit(b, (tx) => setNote(tx, 0, 2, 1, '第二条批注\n两行', '小李'));
    return b;
  },
  '图表': () => {
    const b = blankWorkbook();
    fill(b, [['月份', '销量', '利润'], ['1月', 10, 3], ['2月', 20, 5], ['3月', 15, 4], ['4月', 25, 9]]);
    const g = { r0: 0, c0: 0, r1: 4, c1: 2 };
    const series = seriesFromRange(b, 0, g, guessSource(b.sheets[0], g));
    const types = ['col', 'bar', 'line', 'area', 'pie'] as const;
    types.forEach((type, i) => edit(b, (tx) => addChart(tx, 0, { type, title: `${type} 图`, series }, { c: 4 + (i % 2) * 7, cOff: 0, r: 1 + Math.floor(i / 2) * 17, rOff: 0 }, { c: 10 + (i % 2) * 7, cOff: 0, r: 16 + Math.floor(i / 2) * 17, rOff: 0 })));
    return b;
  },
  // 打开上面生成的“图表”再编辑一遍：改标题、换类型、移动、删一个、再加一个
  '图表编辑': () => {
    const bytes = saveXlsx(CASES['图表']());
    const b = readWorkbook(bytes);
    b.source = bytes;
    const d = b.sheets[0].drawings!;
    edit(b, (tx) => updateChart(tx, 0, 0, { ...d[0].chart!, title: '改过的标题', series: d[0].chart!.series.slice(0, 1) }));
    edit(b, (tx) => updateChart(tx, 0, 1, { ...b.sheets[0].drawings![1].chart!, type: 'line' }));
    edit(b, (tx) => moveDrawing(tx, 0, 2, { c: 1, cOff: 0, r: 60, rOff: 0 }, { c: 7, cOff: 0, r: 75, rOff: 0 }));
    edit(b, (tx) => deleteDrawing(tx, 0, 3));
    const series = b.sheets[0].drawings![0].chart!.series;
    edit(b, (tx) => addChart(tx, 0, { type: 'col', title: '新加的', series }, { c: 14, cOff: 0, r: 1, rOff: 0 }, { c: 20, cOff: 0, r: 16, rOff: 0 }));
    return b;
  },
};

for (const [name, make] of Object.entries(CASES)) {
  const b = make();
  writeFileSync(join(out, `${name}.xlsx`), saveXlsx(b));
  console.log('生成', name);
}

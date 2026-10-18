/**
 * 透视表对拍：生成多种设置的 xlsx（含 Excel 真透视表部件），让 Microsoft Excel 打开并刷新，再把 Excel 里每个透视表的
 * 实际单元格和我们排出来的布局逐格比较。
 *   1. npx vite-node editors/sheet/tools/pivot-oracle.ts gen <目录>      生成文件 + expected.json
 *   2. powershell -File scripts/excel-pivot-grid.ps1 <目录>              Excel 打开并导出 <名>.grid.json
 *   3. npx vite-node editors/sheet/tools/pivot-oracle.ts compare <目录>  比较
 */
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { createPivot, buildPivot, readSource, type PivotConfig } from '../src/core/pivot';
import { saveXlsx } from '../src/core/save';
import { setInput, Tx } from '../src/core/ops';
import { blankWorkbook } from '../src/core/workbook';

const [mode, dir] = process.argv.slice(2);
if (!mode || !dir) throw new Error('用法：pivot-oracle.ts gen|compare <目录>');

const DATA: (string | number | null)[][] = [
  ['区域', '产品', '数量', '金额'],
  ['华东', '甲', 10, 100],
  ['华东', '乙', 5, 80],
  ['华南', '甲', 7, 70],
  ['华东', '甲', 3, 30],
  ['华南', '乙', 2, 25],
  [null, '乙', 1, 10],
];
const T = { rows: true, cols: true };
const cfg = (rows: number[], cols: number[], values: PivotConfig['values'], totals = T): PivotConfig => ({ rows, cols, values, totals });
const sum = (field: number) => ({ field, agg: 'sum' as const });
const CASES: Record<string, PivotConfig> = {
  c01_一行一值: cfg([0], [], [sum(3)]),
  c02_两行一值: cfg([0, 1], [], [sum(3)]),
  c03_一行一列一值: cfg([0], [1], [sum(3)]),
  c04_一行两值: cfg([0], [], [sum(2), { field: 3, agg: 'avg' }]),
  c05_一行一列两值: cfg([0], [1], [sum(2), sum(3)]),
  c06_无行无列: cfg([], [], [sum(3)]),
  c07_各种汇总: cfg([0, 1], [], [{ field: 3, agg: 'count' }, { field: 3, agg: 'counta' }, { field: 3, agg: 'max' }, { field: 3, agg: 'min' }]),
  c08_两行一列: cfg([0, 2], [1], [sum(3)]),
  c09_关总计: cfg([0], [1], [sum(3)], { rows: false, cols: false }),
  c10_只有列: cfg([], [1], [sum(3)]),
  c11_只有列两值: cfg([], [1], [sum(2), sum(3)]),
  c12_无值字段: cfg([0], [], []),
  c13_数字行字段: cfg([2], [], [sum(3)]),
  c14_关总计行: cfg([0, 1], [], [sum(3)], { rows: false, cols: true }),
  c15_关总计列: cfg([0], [1], [sum(2), sum(3)], { rows: true, cols: false }),
};

function make(name: string, c: PivotConfig) {
  const b = blankWorkbook();
  b.sheets[0].name = '数据';
  const tx = new Tx(b, 'd');
  DATA.forEach((row, r) => row.forEach((v, col) => v !== null && setInput(tx, 0, r, col, typeof v === 'string' ? `'${v}` : String(v))));
  tx.commit();
  const t2 = new Tx(b, 'p');
  createPivot(t2, 0, { r0: 0, c0: 0, r1: 6, c1: 3 }, c);
  t2.commit();
  writeFileSync(join(dir, `${name}.xlsx`), saveXlsx(b));
  return buildPivot(readSource(b.sheets[0], { r0: 0, c0: 0, r1: 6, c1: 3 }), c).cells;
}

if (mode === 'gen') {
  mkdirSync(dir, { recursive: true });
  const expected: Record<string, unknown> = {};
  for (const [name, c] of Object.entries(CASES)) expected[name] = make(name, c);
  writeFileSync(join(dir, 'expected.json'), JSON.stringify(expected));
  console.log('生成', Object.keys(CASES).length, '个');
} else {
  const expected = JSON.parse(readFileSync(join(dir, 'expected.json'), 'utf8')) as Record<string, (string | number)[][]>;
  let bad = 0;
  for (const [name, grid] of Object.entries(expected)) {
    const p = join(dir, `${name}.grid.json`);
    if (!existsSync(p)) {
      console.log(`✗ ${name}：没有 Excel 的结果（打开失败？）`);
      bad++;
      continue;
    }
    const raw = JSON.parse(readFileSync(p, 'utf8').replace(/^\uFEFF/, '')) as { ref: string; grid: string[][] } | null;
    if (!raw) {
      console.log(`✗ ${name}：Excel 里没有透视表（部件被丢弃？）`);
      bad++;
      continue;
    }
    const ours = grid.map((r) => r.map((v) => String(v)));
    const theirs = raw.grid;
    const diffs: string[] = [];
    const h = Math.max(ours.length, theirs.length);
    for (let r = 0; r < h; r++) {
      const w = Math.max(ours[r]?.length ?? 0, theirs[r]?.length ?? 0);
      for (let c = 0; c < w; c++) {
        const a = ours[r]?.[c] ?? '';
        const b = theirs[r]?.[c] ?? '';
        // 数字按数值比（Excel 显示 47.5、我们是 47.5；浮点末位忽略）
        const same = a === b || (a !== '' && b !== '' && !Number.isNaN(Number(a)) && !Number.isNaN(Number(b)) && Math.abs(Number(a) - Number(b)) < 1e-9);
        if (!same) diffs.push(`[${r + 1},${c + 1}] 我们=“${a}” Excel=“${b}”`);
      }
    }
    if (diffs.length) {
      bad++;
      console.log(`✗ ${name}（Excel 区域 ${raw.ref}，我们 ${ours.length}×${ours[0]?.length}）\n   ${diffs.slice(0, 8).join('\n   ')}`);
    } else console.log(`✓ ${name}  ${raw.ref}`);
  }
  console.log(bad ? `不一致 ${bad} 个` : '全部一致');
  process.exitCode = bad ? 1 : 0;
}

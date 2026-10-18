/** 精确修复旧版 UTF-8 损坏的打印机设置引用，其它 ZIP 部件逐字节保留。
 * npx vite-node editors/sheet/tools/repair-printer-settings.ts <输入.xlsx> <输出.xlsx>
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { unzipSync, zipSync } from 'fflate';
import { detachDamagedPrinterSettings } from '../src/core/printerSettings';

const [input, output] = process.argv.slice(2);
if (!input || !output) throw new Error('需要输入和输出路径；原文件不覆盖');
const files = unzipSync(new Uint8Array(readFileSync(input)));
const count = detachDamagedPrinterSettings(files);
if (!count) throw new Error('没有识别到可修复的历史打印机数据损坏，未写出文件');
writeFileSync(output, zipSync(files, { level: 6 }), { flag: 'wx' });
console.log(`已断开 ${count} 个损坏的打印机设置引用；其它部件保持不变。`);

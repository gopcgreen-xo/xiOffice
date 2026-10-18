/** 把一个 xlsx 经本编辑器打开、改一格、保存：用来验证文件里的图表 / 透视表 / 批注等原有内容保存后完好（再用 excel-check.ps1 看）。
 *  npx vite-node editors/sheet/tools/roundtrip.ts <输入.xlsx> <输出.xlsx> */
import { readFileSync, writeFileSync } from 'node:fs';
import { unzipSync } from 'fflate';
import { saveXlsx } from '../src/core/save';
import { setInput, Tx } from '../src/core/ops';
import { readWorkbook } from '../src/core/workbook';

const [input, output] = process.argv.slice(2);
if (!input || !output) throw new Error('用法：roundtrip.ts <输入.xlsx> <输出.xlsx>');
const bytes = new Uint8Array(readFileSync(input));
const b = readWorkbook(bytes);
b.source = bytes;
const same = Buffer.compare(Buffer.from(saveXlsx(b)), Buffer.from(bytes)) === 0;
console.log('没改动保存 = 原文件字节：', same);
{
  const a = unzipSync(bytes);
  const c = unzipSync(saveXlsx(b));
  const diff = [...new Set([...Object.keys(a), ...Object.keys(c)])].filter((k) => !a[k] || !c[k] || Buffer.compare(Buffer.from(a[k]), Buffer.from(c[k])) !== 0);
  console.log('没改动保存：内容不同的部件 =', diff.length ? diff.join(', ') : '（无，只是 zip 压缩方式不同）');
}
const tx = new Tx(b, 'x');
setInput(tx, 0, 0, 5, '往返测试');
tx.commit();
writeFileSync(output, saveXlsx(b));
console.log('已保存', output);

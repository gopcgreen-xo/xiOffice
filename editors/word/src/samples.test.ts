// @vitest-environment jsdom
/**
 * 回归样本的往返测试：本编辑器 samples/ 下的每个文件打开 → 修改 → 保存 → 再打开，内容不能丢、不能变。
 * 真实文件（脱敏后）放进这个目录就会自动参加测试；在 Office 里打开的人工核对步骤见 editors/SAMPLES.md。
 */
import { readdirSync, readFileSync, statSync } from 'fs';
import { join } from 'path';
import { expect, describe, it } from 'vitest';
import { readDocx, writeDocx } from './ooxml/docx';
import { insertRaw, paragraphPaths, getPara } from './core/ops';
import { paraText } from './core/model';

const DIR = join(__dirname, '../samples');

function files(dir: string): string[] {
  let out: string[] = [];
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) out = out.concat(files(p));
    else if (/\.(docx)$/i.test(name) && !name.startsWith('~$')) out.push(p);
  }
  return out;
}

const docx = files(DIR);

describe.runIf(docx.length)('Word 样本往返', () => {
  it.each(docx)('%s', (file) => {
    const doc = readDocx(new Uint8Array(readFileSync(file)));
    const texts = (d: typeof doc) => paragraphPaths(d, 'body').map((path) => paraText(getPara(d, { story: 'body', path })));
    const before = texts(doc);
    // 第一段开头插入一个字，保存后再读：只多这一个字，其它段落不变
    insertRaw(doc, { story: 'body', path: paragraphPaths(doc, 'body')[0], offset: 0 }, '★');
    const again = readDocx(writeDocx(doc));
    const after = texts(again);
    expect(after[0]).toBe(`★${before[0]}`);
    expect(after.slice(1)).toEqual(before.slice(1));
  });
});

it('样本目录存在', () => {
  expect(statSync(DIR).isDirectory()).toBe(true);
});

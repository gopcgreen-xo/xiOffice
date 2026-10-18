// @vitest-environment jsdom
/**
 * 生成给 Microsoft Word 人工 / 脚本验收的文件（WORD_CHECK_DIR=输出目录 时才运行）：
 * 每个样本做几种编辑后用字节级保留保存，再由 scripts/word-check.ps1 用 Word 打开，确认没有修复提示、内容在。
 */
import { mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'fs';
import { basename, join } from 'path';
import { it } from 'vitest';
import { openDocx, patchDocx } from './patch';
import { addComment, applyNumbering, insertImage, insertRaw, paragraphPaths } from '../core/ops';
import { registerImage } from '../core/images';
import type { Paragraph } from '../core/model';

const OUT = process.env.WORD_CHECK_DIR;
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 15, 4, 0, 9, 251, 3, 253, 227, 85, 242, 156, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);

it.runIf(OUT)('生成 Word 验收文件', async () => {
  mkdirSync(OUT!, { recursive: true });
  const dir = join(__dirname, '../../samples');
  const files = (d: string): string[] => readdirSync(d).flatMap((n) => (statSync(join(d, n)).isDirectory() ? files(join(d, n)) : /\.docx$/i.test(n) ? [join(d, n)] : []));
  const extra = (process.env.WORD_CHECK_EXTRA ?? '').split(';').filter(Boolean);
  for (const file of [...files(dir), ...extra]) {
    const name = basename(file, '.docx');
    const bytes = new Uint8Array(readFileSync(file));
    for (const kind of ['text', 'format', 'objects']) {
      const { doc, origin } = await openDocx(bytes);
      if (!origin) throw new Error(`${name}: docx-engine 读不了`);
      const paths = paragraphPaths(doc, 'body');
      const p0 = paths[0];
      if (kind === 'text') insertRaw(doc, { story: 'body', path: p0, offset: 0 }, '【已修改】');
      if (kind === 'format') {
        for (const b of doc.body) {
          if (b.kind !== 'p') continue;
          b.props.align = 'center';
          for (const r of b.runs) r.props = { ...r.props, bold: !r.props.bold };
          break;
        }
      }
      if (kind === 'objects') {
        const last = paths[paths.length - 1];
        insertImage(doc, { story: 'body', path: p0, offset: 0 }, { id: registerImage(PNG, 'image/png'), w: 30, h: 30 });
        applyNumbering(doc, { story: 'body', path: last, offset: 0 }, { story: 'body', path: last, offset: 0 }, 'decimal');
        const target = doc.body.find((b): b is Paragraph => b.kind === 'p' && b.runs.some((r) => r.text.trim()));
        if (target) {
          const at = doc.body.indexOf(target);
          addComment(doc, { story: 'body', path: [at], offset: 0 }, { story: 'body', path: [at], offset: 1 }, '验收批注', 'Claude');
        }
      }
      writeFileSync(join(OUT!, `${name}-${kind}.docx`), await patchDocx(doc, origin));
    }
  }
}, 120000);

// @vitest-environment jsdom
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { openDocx, patchDocx } from './patch';
import { applyAiText, clearCellContents, deleteRange } from '../core/ops';
import { paraText, type Doc, type Paragraph } from '../core/model';

const bytes = (name: string) => new Uint8Array(readFileSync(resolve('editors/word/samples/editing-baseline', `${name}.docx`)));
const at = (path: number[], offset = 0) => ({ story: 'body' as const, path, offset });
// Word adds/moves _GoBack bookmarks when saving. They are zero-width in Word,
// but represented by atomic runs in our editing coordinates.
const visibleText = (p: Paragraph) => p.runs.filter((r) => r.opaque?.kind !== 'bookmark').map((r) => r.text).join('');
const content = (d: Doc) => d.body.map((b) => b.kind === 'p' ? visibleText(b) : b.rows.map((r) => r.cells.map((c) => c.paragraphs.map(visibleText))));

describe('Microsoft Word generated editing fixtures', () => {
  for (const name of ['body-table-body', 'body-table-body-type', 'whole-cells', 'whole-table']) {
    it(`${name}: edit and DOCX roundtrip match Word document structure and text`, async () => {
      const { doc, origin } = await openDocx(bytes(`${name}-before`));
      const expected = (await openDocx(bytes(`${name}-after`))).doc;
      const tableIndex = doc.body.findIndex((b) => b.kind === 'table');
      const endIndex = doc.body.findIndex((b) => b.kind === 'p' && paraText(b) === 'after');
      expect(tableIndex).toBeGreaterThan(0);
      expect(endIndex).toBeGreaterThan(tableIndex);
      const start = at([0], paraText(doc.body[0] as Paragraph).indexOf('before') + 2);
      if (name === 'body-table-body') deleteRange(doc, start, at([endIndex], 2));
      else if (name === 'body-table-body-type') applyAiText(doc, start, at([endIndex], 2), 'X', 'replace');
      else clearCellContents(doc, at([tableIndex, 0, 0, 0]), at([tableIndex, name === 'whole-table' ? 1 : 0, 1, 0]));
      expect(content(doc)).toEqual(content(expected));
      const saved = await patchDocx(doc, origin!);
      if (process.env.WORD_EDITING_OUT) {
        mkdirSync(process.env.WORD_EDITING_OUT, { recursive: true });
        writeFileSync(resolve(process.env.WORD_EDITING_OUT, `${name}.docx`), saved);
      }
      const reopened = (await openDocx(saved)).doc;
      expect(content(reopened)).toEqual(content(expected));
      const actualParas = reopened.body.filter((b): b is Paragraph => b.kind === 'p');
      const expectedParas = expected.body.filter((b): b is Paragraph => b.kind === 'p');
      expect(actualParas.map((p) => p.props.line)).toEqual(expectedParas.map((p) => p.props.line));
      expect(reopened.setup).toEqual(expected.setup);
    });
  }
});

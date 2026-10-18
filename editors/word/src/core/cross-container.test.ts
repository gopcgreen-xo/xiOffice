import { describe, expect, it } from 'vitest';
import { newDocument, newParagraph, newTable, paraText } from './model';
import { applyAiText, clearCellContents, deleteRange, History, tracking } from './ops';
import type { Pos } from './layout';

const at = (path: number[], offset = 0): Pos => ({ story: 'body', path, offset });
const fixture = () => {
  const doc = newDocument();
  doc.body = [newParagraph('before'), newTable(2, 2, 240, doc.defaults), newParagraph('after')];
  return doc;
};

describe('Word 16.0 cross-container editing baseline', () => {
  it('Delete spanning a complete table preserves the two boundary paragraphs', () => {
    const doc = fixture();
    const first = doc.body[0];
    const last = doc.body[2];
    expect(deleteRange(doc, at([0], 2), at([2], 2))).toEqual(at([0], 2));
    expect(doc.body).toEqual([first, last]);
    expect(doc.body.map((p) => paraText(p as ReturnType<typeof newParagraph>))).toEqual(['be', 'ter']);
  });

  it('typing over the same range joins the remaining text, including reverse selection', () => {
    const doc = fixture();
    expect(applyAiText(doc, at([2], 2), at([0], 2), 'X', 'replace')).toEqual(at([0], 3));
    expect(doc.body).toHaveLength(1);
    expect(paraText(doc.body[0] as ReturnType<typeof newParagraph>)).toBe('beXter');
  });

  it('partial table boundaries are rejected without mutation, as in Word Range.Delete', () => {
    for (const [a, b] of [
      [at([0], 2), at([1, 0, 0, 0], 0)],
      [at([1, 0, 0, 0], 0), at([2], 2)],
    ]) {
      const doc = fixture();
      const before = structuredClone(doc);
      expect(deleteRange(doc, a, b)).toBeNull();
      expect(doc).toEqual(before);
    }
  });

  it('different stories cannot delete matching paragraph paths, with or without tracking', () => {
    for (const enabled of [false, true]) {
      const doc = fixture();
      const before = structuredClone(doc);
      tracking.on = enabled;
      try {
        expect(deleteRange(doc, at([0], 0), { story: 'header', path: [0], offset: 2 })).toBeNull();
        expect(doc).toEqual(before);
      } finally { tracking.on = false; }
    }
  });

  it('undo restores the complete table and untouched formatting', () => {
    const doc = fixture();
    const before = structuredClone(doc);
    const history = new History<Pos>();
    history.record(doc, at([0], 2), 'delete');
    deleteRange(doc, at([0], 2), at([2], 2));
    expect(history.undo(doc, at([0], 2))?.doc).toEqual(before);
  });

  it('Delete on a cell rectangle clears contents but preserves widths, shading and other cells', () => {
    const doc = fixture();
    const table = doc.body[1];
    if (table.kind !== 'table') throw new Error('fixture table missing');
    table.rows.forEach((row, r) => row.cells.forEach((c, ci) => {
      c.paragraphs = [newParagraph(`${r}:${ci}`), newParagraph('second paragraph')];
      c.shading = 'FF0000';
    }));
    const before = structuredClone(table);
    expect(clearCellContents(doc, at([1, 0, 1, 0]), at([1, 0, 0, 0]))).toEqual(at([1, 0, 0, 0]));
    expect(table.cols).toEqual(before.cols);
    expect(table.rows[1]).toEqual(before.rows[1]);
    for (const c of table.rows[0].cells) {
      expect(c.shading).toBe('FF0000');
      expect(c.paragraphs).toHaveLength(1);
      expect(paraText(c.paragraphs[0])).toBe('');
    }
  });

  it('a selection on a vertical merge continuation clears its owner without splitting the merge', () => {
    const doc = fixture();
    const table = doc.body[1];
    if (table.kind !== 'table') throw new Error('fixture table missing');
    table.rows[0].cells[0].vMerge = 'restart';
    table.rows[0].cells[0].paragraphs = [newParagraph('merged')];
    table.rows[1].cells[0].vMerge = 'continue';
    table.rows[0].cells[1].paragraphs = [newParagraph('keep')];
    clearCellContents(doc, at([1, 1, 0, 0]), at([1, 1, 0, 0]));
    expect(paraText(table.rows[0].cells[0].paragraphs[0])).toBe('');
    expect(table.rows[0].cells[0].vMerge).toBe('restart');
    expect(table.rows[1].cells[0].vMerge).toBe('continue');
    expect(paraText(table.rows[0].cells[1].paragraphs[0])).toBe('keep');
  });

  it('tracking records cell text deletions without removing table structure', () => {
    const doc = fixture();
    const table = doc.body[1];
    if (table.kind !== 'table') throw new Error('fixture table missing');
    table.rows[0].cells[0].paragraphs = [newParagraph('tracked')];
    tracking.on = true;
    try {
      clearCellContents(doc, at([1, 0, 0, 0]), at([1, 0, 1, 0]));
      const p = table.rows[0].cells[0].paragraphs[0];
      expect(paraText(p)).toBe('tracked');
      expect(p.runs[0].props.rev?.kind).toBe('del');
      expect(table.rows).toHaveLength(2);
      expect(table.rows[0].cells).toHaveLength(2);
    } finally { tracking.on = false; }
  });
});

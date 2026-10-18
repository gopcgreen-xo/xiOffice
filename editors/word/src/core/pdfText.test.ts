import { describe, expect, it } from 'vitest';
import type { Glyph, PageLayout, PlacedLine } from './layout';
import { FIELD_CHAR } from './model';
import { pageTextLines } from './pdfText';

const g = (ch: string, x: number, w = 10, size = 10.5, extra: Partial<Glyph> = {}): Glyph =>
  ({ ch, x, w, props: { size } as Glyph['props'], family: 'SimSun', cjk: true, ...extra }) as Glyph;

const placed = (glyphs: Glyph[], x = 90, y = 72, baseline = 12): PlacedLine =>
  ({ story: 'body', path: [0], x, y, line: { start: 0, end: glyphs.length, xs: [], glyphs, right: 0, height: 16, baseline, brk: 'end' } }) as unknown as PlacedLine;

const page = (body: PlacedLine[], extra: Partial<PageLayout> = {}): PageLayout =>
  ({ index: 0, section: 0, setup: { width: 595, height: 842 }, body, header: [], footer: [], borders: [], fills: [], floats: [], ...extra }) as unknown as PageLayout;

describe('pageTextLines', () => {
  it('puts each line at its baseline in PDF coordinates (origin bottom-left)', () => {
    const [l] = pageTextLines(page([placed([g('合', 0), g('同', 10)])]));
    expect(l).toEqual({ text: '合同', origin: [90, 842 - 84], size: 10.5, width: 20 });
  });

  it('splits a line at pictures, tabs and leaders, and skips field placeholders', () => {
    const lines = pageTextLines(
      page([placed([g('第', 0), g('一', 10), g('\t', 20, 30), g('…', 50, 100, 10.5, { leader: '…' }), g('5', 150, 6), g(FIELD_CHAR, 156, 20, 10.5, { image: { id: 'i', w: 20, h: 20 } as Glyph['image'] }), g('页', 176)])]),
    );
    expect(lines.map((l) => [l.text, l.origin[0], l.width])).toEqual([
      ['第一', 90, 20],
      ['5', 240, 6],
      ['页', 266, 10],
    ]);
  });

  it('includes header, footer and footnotes, ignores blank lines and uses the largest size in a segment', () => {
    const lines = pageTextLines(
      page([placed([g(' ', 0), g(' ', 10)])], {
        header: [placed([g('页', 0, 9, 9)], 90, 30)],
        footer: [placed([g('1', 0, 5, 9), g('A', 5, 8, 14)], 280, 800)],
        notes: [placed([g('注', 0)], 90, 760)],
      }),
    );
    expect(lines.map((l) => [l.text, l.size])).toEqual([
      ['页', 9],
      ['1A', 14],
      ['注', 10.5],
    ]);
  });
});

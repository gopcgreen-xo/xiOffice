import { describe, expect, it } from 'vitest';
import { layoutDocument, type Measurer } from './layout';
import { newDocument, newParagraph } from './model';
import { drawPage } from './render';
import { isCJK } from './text';

const fake: Measurer = { width: (ch) => (isCJK(ch) ? 1 : 0.5), metrics: () => ({ ascent: 0.88, descent: 0.12 }) };

/** 记录 fillRect 调用顺序（连同当时的 fillStyle）的假画布 */
function recorder() {
  const calls: { op: string; style: string }[] = [];
  const state: Record<string, unknown> = { fillStyle: '' };
  const ctx = new Proxy(state, {
    get(t, k: string) {
      if (k in t) return t[k];
      if (k === 'measureText') return () => ({ width: 5 });
      if (k === 'fillRect') return () => calls.push({ op: 'fillRect', style: String(t.fillStyle) });
      if (k === 'fillText') return () => calls.push({ op: 'fillText', style: String(t.fillStyle) });
      return () => undefined;
    },
    set(t, k: string, v) {
      t[k] = v;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, calls };
}

describe('选区底色画在单元格底纹之上', () => {
  it('有底纹（白色）的单元格里，选中的蓝色底画在白色底纹之后、文字之前，不会被盖住', () => {
    const d = newDocument();
    d.body = [newParagraph('签字')];
    const pg = layoutDocument(d, fake).pages[0];
    pg.fills = [{ x: 50, y: 50, w: 200, h: 40, color: '#FFFFFF' }];
    const { ctx, calls } = recorder();
    drawPage(ctx, pg, 1, { story: 'body', selection: [{ x: 60, y: 60, w: 30, h: 12 }], showMarks: false });
    const white = calls.findIndex((c) => c.op === 'fillRect' && c.style === '#FFFFFF' && true && calls.indexOf(c) > 0);
    const blue = calls.findIndex((c) => c.op === 'fillRect' && c.style === 'rgba(64, 128, 255, 0.28)');
    const text = calls.findIndex((c) => c.op === 'fillText');
    expect(white).toBeGreaterThanOrEqual(0);
    expect(blue).toBeGreaterThan(white);
    expect(text).toBeGreaterThan(blue);
  });
});

import { describe, expect, it } from 'vitest';
import { bodyLineAt, layoutDocument, type Measurer } from './layout';
import { newDocument, newParagraph } from './model';
import { isCJK } from './text';

const fake: Measurer = { width: (ch) => (isCJK(ch) ? 1 : 0.5), metrics: () => ({ ascent: 0.88, descent: 0.12 }) };

describe('bodyLineAt：点在正文行上（含落在页边距里的正文行）', () => {
  it('点在某一行文字范围内返回这一行；离开文字范围返回 null', () => {
    const d = newDocument();
    d.body = [newParagraph('合同编号')];
    const pg = layoutDocument(d, fake).pages[0];
    const pl = pg.body[0];
    const cx = pl.x + pl.line.xs[0] + 5;
    const cy = pl.y + pl.line.height / 2;
    expect(bodyLineAt(pg, cx, cy)).toBe(pl);
    expect(bodyLineAt(pg, cx, pl.y + pl.line.height + 50)).toBeNull();
    expect(bodyLineAt(pg, pl.x + pl.line.right + 200, cy)).toBeNull();
  });

  it('正文行被排到页面下边距里（浮动表格超出版心）时，也能在那里点中它', () => {
    const d = newDocument();
    d.body = [newParagraph('签字')];
    const pg = layoutDocument(d, fake).pages[0];
    // 人为把这一行挪到下边距里（浮动表格就是这样落在下边距里的）
    const pl = pg.body[0];
    pl.y = pg.setup.height - pg.setup.marginBottom + 10;
    const hit = bodyLineAt(pg, pl.x + pl.line.xs[0] + 3, pl.y + 2);
    expect(hit).toBe(pl);
    expect(pl.y).toBeGreaterThan(pg.setup.height - pg.setup.marginBottom);
  });
});

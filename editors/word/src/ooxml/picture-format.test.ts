// @vitest-environment jsdom
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { openDocx, patchDocx } from './patch';
import { readDocx, writeDocx } from './docx';
import { applyPictureFormat, pictureBounds, pictureFormat } from '../core/picture';
import { getPara, paragraphPaths } from '../core/ops';
import { registerImage } from '../core/images';
import { FIELD_CHAR, newDocument, newParagraph, type Doc, type Run } from '../core/model';
import type { Pos } from '../core/layout';

const fixture = (name: string) => new Uint8Array(readFileSync(resolve('editors/word/samples/picture-baseline', `${name}.docx`)));
function find(doc: Doc): { run: Run; at: Pos } {
  for (const path of paragraphPaths(doc, 'body')) {
    let offset = 0;
    for (const run of getPara(doc, { story: 'body', path }).runs) {
      if (run.image || run.box?.image) return { run, at: { story: 'body', path, offset } };
      offset += run.text.length;
    }
  }
  throw new Error('picture not found');
}
const xml = (b: Uint8Array) => strFromU8(unzipSync(b)['word/document.xml']);
const exportResult = (name: string, bytes: Uint8Array) => {
  if (!process.env.WORD_PICTURE_OUT) return;
  mkdirSync(process.env.WORD_PICTURE_OUT, { recursive: true });
  writeFileSync(resolve(process.env.WORD_PICTURE_OUT, `${name}.docx`), bytes);
};

describe('picture editing with Microsoft Word baselines', () => {
  for (const name of ['inline', 'crop', 'rotate90', 'rotate30', 'flip', 'square', 'top-bottom']) {
    it(`${name}: reads actual Word transforms and keeps an untouched file byte-identical`, async () => {
      const bytes = fixture(name);
      const { doc, origin } = await openDocx(bytes);
      const f = pictureFormat(find(doc).run)!;
      expect(f.image.w).toBeGreaterThan(0);
      if (name === 'crop') {
        expect(f.image.crop?.l).toBeCloseTo(0.13331, 5);
        expect(f.image.crop?.t).toBeCloseTo(0.06666, 5);
        expect(f.image.w).toBeCloseTo(104, 2);
      }
      if (name === 'rotate90') expect(f.image.rot).toBe(90);
      if (name === 'rotate30') expect(f.image.rot).toBe(30);
      if (name === 'flip') expect(f.image.flipH).toBe(true);
      if (name === 'square') expect(f.wrap).toBe('square');
      if (name === 'top-bottom') expect(f.wrap).toBe('topAndBottom');
      expect(await patchDocx(doc, origin!)).toEqual(bytes);
    });
  }

  for (const name of ['crop', 'rotate90', 'flip', 'square']) {
    it(`${name}: resizing/moving preserves transforms and original drawing metadata`, async () => {
      const bytes = fixture(name);
      const { doc, origin } = await openDocx(bytes);
      const { run, at } = find(doc);
      const f = pictureFormat(run)!;
      f.image.w *= 1.5; f.image.h *= 1.5;
      f.posH = { rel: 'page', offset: 100 };
      applyPictureFormat(doc, at, f);
      const saved = await patchDocx(doc, origin!);
      const back = pictureFormat(find((await openDocx(saved)).doc).run)!;
      expect(back.image.w).toBeCloseTo(f.image.w, 3);
      expect(back.image.h).toBeCloseTo(f.image.h, 3);
      expect(back.image.crop).toEqual(f.image.crop);
      expect(back.image.rot ?? 0).toBe(f.image.rot ?? 0);
      expect(back.image.flipH ?? false).toBe(f.image.flipH ?? false);
      expect(xml(saved).match(/<wp:docPr\b[^>]*\/>/)?.[0]).toBe(xml(bytes).match(/<wp:docPr\b[^>]*\/>/)?.[0]);
      if (f.wrap !== 'inline') expect(back.posH).toEqual(f.posH);
      exportResult(`resize-${name}`, saved);
    });
  }

  it('new pictures roundtrip rotation, flip and crop through full DOCX serialization', () => {
    const doc = newDocument();
    const id = registerImage(new Uint8Array(readFileSync('editors/word/samples/picture-baseline/quadrants.png')), 'image/png');
    const para = newParagraph();
    para.runs = [{ text: FIELD_CHAR, props: doc.defaults, image: { id, w: 120, h: 60, rot: 90, flipV: true, crop: { l: 0.1, t: 0, r: 0.2, b: 0 } } }];
    doc.body = [para];
    const saved = writeDocx(doc);
    const f = pictureFormat(find(readDocx(saved)).run)!;
    expect(f.image.rot).toBe(90);
    expect(f.image.flipV).toBe(true);
    expect(f.image.crop).toEqual({ l: 0.1, t: 0, r: 0.2, b: 0 });
    expect(pictureBounds(f.image).w).toBeCloseTo(60);
    expect(pictureBounds(f.image).h).toBeCloseTo(120);
    exportResult('new-rotated-inline', saved);
  });

  it('switches inline pictures to floating and back without losing source transform', async () => {
    const { doc, origin } = await openDocx(fixture('crop'));
    const { run, at } = find(doc);
    const f = pictureFormat(run)!;
    f.wrap = 'square';
    f.posH = { rel: 'page', offset: 72 };
    f.image.rot = 30;
    applyPictureFormat(doc, at, f);
    const floating = await patchDocx(doc, origin!);
    exportResult('converted-floating', floating);
    const next = await openDocx(floating);
    const hit = find(next.doc);
    const back = pictureFormat(hit.run)!;
    expect(back.wrap).toBe('square');
    expect(back.image.crop).toEqual(f.image.crop);
    expect(back.image.rot).toBe(30);
    back.wrap = 'inline';
    applyPictureFormat(next.doc, hit.at, back);
    const inline = await patchDocx(next.doc, next.origin!);
    for (const [name, data] of Object.entries(unzipSync(inline))) {
      if (!name.endsWith('.rels')) continue;
      const ids = [...strFromU8(data).matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]);
      expect(new Set(ids).size, `${name}: no duplicate relationship ids after repeated saves`).toBe(ids.length);
    }
    expect(pictureFormat(find((await openDocx(inline)).doc).run)!.wrap).toBe('inline');
    exportResult('converted-inline', inline);
  });

  it('replaces image bytes while retaining crop, dimensions and original metadata', async () => {
    const { doc, origin } = await openDocx(fixture('crop'));
    const { run, at } = find(doc);
    const f = pictureFormat(run)!;
    const png = new Uint8Array(readFileSync('editors/word/samples/picture-baseline/quadrants.png'));
    // Valid PNG trailing data makes a distinct media id without changing its decoded pixels.
    const replacement = new Uint8Array(png.length + 1); replacement.set(png);
    const oldId = f.image.id;
    f.image.id = registerImage(replacement, 'image/png');
    expect(f.image.id).not.toBe(oldId);
    applyPictureFormat(doc, at, f);
    const saved = await patchDocx(doc, origin!);
    const back = pictureFormat(find((await openDocx(saved)).doc).run)!;
    expect(back.image.id).toBe(f.image.id);
    expect(back.image.crop).toEqual(f.image.crop);
    expect(back.image.w).toBe(f.image.w);
    exportResult('replaced', saved);
  });
});

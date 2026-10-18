import { type Doc, type Run, type RunImage, type TextBox } from './model';
import { getPara } from './ops';
import type { Pos } from './layout';

export type PictureWrap = 'inline' | 'square' | 'topAndBottom' | 'front' | 'behind' | 'tight' | 'through';
export interface PictureFormat {
  image: RunImage;
  wrap: PictureWrap;
  posH: TextBox['posH'];
  posV: TextBox['posV'];
  dist: NonNullable<TextBox['dist']>;
}

export const pictureKey = (image: RunImage) => JSON.stringify({ ...image, raw: undefined, rawId: undefined });

export function pictureBounds(image: Pick<RunImage, 'w' | 'h' | 'rot'>) {
  const angle = (image.rot ?? 0) * Math.PI / 180;
  const cos = Math.abs(Math.cos(angle)), sin = Math.abs(Math.sin(angle));
  return { w: image.w * cos + image.h * sin, h: image.w * sin + image.h * cos };
}

export function pictureLocalPoint(image: Pick<RunImage, 'w' | 'h' | 'rot'>, dx: number, dy: number) {
  const angle = -(image.rot ?? 0) * Math.PI / 180;
  const x = dx - image.w / 2, y = dy - image.h / 2;
  return { x: x * Math.cos(angle) - y * Math.sin(angle) + image.w / 2, y: x * Math.sin(angle) + y * Math.cos(angle) + image.h / 2 };
}

export function pictureAt(doc: Doc, at: Pos): Run | null {
  let offset = 0;
  for (const run of getPara(doc, at).runs) {
    if (offset === at.offset && (run.image || run.box?.image)) return run;
    offset += run.text.length;
  }
  return null;
}

export function pictureFormat(run: Run): PictureFormat | null {
  const image = run.image ?? run.box?.image;
  if (!image) return null;
  const box = run.box;
  return {
    image: { ...image, crop: image.crop && { ...image.crop } },
    wrap: !box ? 'inline' : box.wrap === 'none' ? box.behind ? 'behind' : 'front' : box.wrap,
    posH: { ...(box?.posH ?? { rel: 'column', offset: 0 }) },
    posV: { ...(box?.posV ?? { rel: 'paragraph', offset: 0 }) },
    dist: { ...(box?.dist ?? { l: 9, r: 9, t: 0, b: 0 }) },
  };
}

export function pictureFormatError(value: PictureFormat): string | null {
  const i = value.image;
  if (![i.w, i.h].every((v) => Number.isFinite(v) && v > 0 && v <= 1584)) return '图片宽度和高度必须大于 0 且不超过 55.88 厘米';
  if (!Number.isFinite(i.rot ?? 0)) return '请输入有效旋转角度';
  const c = i.crop;
  if (c && (!Object.values(c).every((n) => Number.isFinite(n)) || c.l + c.r >= 1 || c.t + c.b >= 1)) return '裁剪后必须保留可见的图片区域';
  if (![value.posH.offset, value.posV.offset, ...Object.values(value.dist)].every(Number.isFinite)) return '请输入有效的位置和文字距离';
  if (Object.values(value.dist).some((n) => n < 0)) return '与文字的距离不能为负数';
  return null;
}

/** Update only the selected picture. Unedited raw effects and alt text stay available to the OOXML patcher. */
export function applyPictureFormat(doc: Doc, at: Pos, value: PictureFormat): Run | null {
  const error = pictureFormatError(value);
  if (error) throw new Error(error);
  const run = pictureAt(doc, at);
  if (!run) return null;
  const image = { ...value.image, rot: ((value.image.rot ?? 0) % 360 + 360) % 360 };
  if (value.wrap === 'inline') {
    // A container conversion must rebuild its drawing, not reuse the old anchor.
    if (run.box) { delete image.raw; delete image.rawId; }
    run.image = image;
    delete run.box;
  } else {
    const old = run.box;
    if (!old) { delete image.raw; delete image.rawId; }
    run.box = {
      ...(old ?? { fill: null, line: null, inset: { l: 0, t: 0, r: 0, b: 0 }, vAlign: 'top', paragraphs: [] }),
      image, w: image.w, h: image.h,
      wrap: value.wrap === 'front' || value.wrap === 'behind' ? 'none' : value.wrap,
      behind: value.wrap === 'behind',
      posH: { ...value.posH }, posV: { ...value.posV }, dist: { ...value.dist },
    };
    delete run.image;
  }
  return run;
}

/** Draw within the unrotated picture viewport, cropping without discarding source pixels. */
export function drawPicture(ctx: CanvasRenderingContext2D, source: CanvasImageSource, image: RunImage, x: number, y: number) {
  const c = image.crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const sw = 1 - c.l - c.r, sh = 1 - c.t - c.b;
  if (!(sw > 0 && sh > 0)) return;
  ctx.save();
  ctx.translate(x + image.w / 2, y + image.h / 2);
  ctx.rotate((image.rot ?? 0) * Math.PI / 180);
  ctx.scale(image.flipH ? -1 : 1, image.flipV ? -1 : 1);
  ctx.translate(-image.w / 2, -image.h / 2);
  ctx.beginPath();
  ctx.rect(0, 0, image.w, image.h);
  ctx.clip();
  ctx.drawImage(source, -c.l * image.w / sw, -c.t * image.h / sh, image.w / sw, image.h / sh);
  ctx.restore();
}

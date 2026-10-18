/**
 * 把文档模型画成页面图片（和编辑器里看到的一致）：给“套打”预览模板、给“转换为 PDF”出页面用。
 * 字体度量（行高和 Word 一致要靠它）只读一次；图片先解码好再排版绘制。
 */
import type { Doc } from '../core/model';
import { imageIds } from '../core/model';
import { loadImages } from '../core/images';
import { registerFontTables } from '../core/fontfile';
import { registerOfficeFallback } from '../core/officeMetrics';
import { isMobile } from '@wz/editor-shared';
import { layoutDocument, type DocLayout } from '../core/layout';
import { CanvasMeasurer, renderPagesToImages } from '../core/render';

let measurer: CanvasMeasurer | null = null;
let fontsLoaded: Promise<unknown> | null = null;

/** 系统字体的度量表（宿主的 fontFaces），所有调用共用一次 */
export function prepareFonts(fontFaces?: () => Promise<Record<string, string>[]>): Promise<unknown> {
  if (fontFaces)
    fontsLoaded ??= fontFaces()
      .then(registerFontTables)
      .catch(() => 0)
      // 手机版：系统字体之后补上常用 Office 字体的度量（docs/ANDROID.md §9）
      .then(() => isMobile() && registerOfficeFallback());
  return fontsLoaded ?? Promise.resolve();
}

/** 排版（图片已解码） */
export async function layoutForRender(doc: Doc, fontFaces?: () => Promise<Record<string, string>[]>): Promise<DocLayout> {
  await prepareFonts(fontFaces);
  await loadImages(imageIds(doc));
  return layoutDocument(doc, (measurer ??= new CanvasMeasurer()));
}

export interface PageImage {
  /** PNG data URL */
  url: string;
  /** 页面大小（磅） */
  w: number;
  h: number;
}

/** 每页一张图片；pxPerPt = 每磅多少像素（预览 1.5 左右就够清楚） */
export async function renderDocPages(doc: Doc, opts: { fontFaces?: () => Promise<Record<string, string>[]>; pxPerPt?: number } = {}): Promise<PageImage[]> {
  const layout = await layoutForRender(doc, opts.fontFaces);
  const urls = renderPagesToImages(layout, opts.pxPerPt ?? 1.5);
  return layout.pages.map((pg, i) => ({ url: urls[i]!, w: pg.setup.width, h: pg.setup.height }));
}

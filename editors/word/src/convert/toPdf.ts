/**
 * Word 文档（.docx / .doc）→ PDF：用编辑器的排版和渲染（和编辑器里看到的一致）把每页画成图片，
 * 再按排版里每行文字的位置写一层不可见文字（`addTextLayer` 由调用方注入：PDF 编辑器的 PDFium 引擎），
 * 这样 PDF 里的文字可以选中、复制、查找。没有注入或写入失败时，结果是只有图片的 PDF。
 */
import { pagesToPdf } from '@wz/editor-shared';
import { openDocx } from '../ooxml/patch';
import { isDoc, readDoc } from '../doc97/read';
import { renderPagesToImages } from '../core/render';
import { layoutForRender } from './preview';
import { pageTextLines, type PdfTextLine } from '../core/pdfText';

export interface TextLayerPage {
  pageIndex: number;
  rotate: number;
  lines: PdfTextLine[];
}

export interface WordPdfDeps {
  /** 系统字体的度量表（宿主的 fontFaces），行高和 Word 一致要靠它；只读一次 */
  fontFaces?: () => Promise<Record<string, string>[]>;
  addTextLayer?: (pdf: Uint8Array, pages: TextLayerPage[]) => Promise<{ bytes: Uint8Array | null; problems: string[] }>;
}

export interface WordPdfResult {
  bytes: Uint8Array;
  pages: number;
  /** 写了文字层的页数（0 = 只有图片） */
  textPages: number;
  problems: string[];
  /** 文档里编辑器还不支持、PDF 里没有的内容 */
  lossy: string[];
}

export async function wordToPdf(bytes: Uint8Array, deps: WordPdfDeps = {}): Promise<WordPdfResult> {
  const doc = isDoc(bytes) ? readDoc(bytes) : (await openDocx(bytes)).doc;
  const layout = await layoutForRender(doc, deps.fontFaces);
  const urls = renderPagesToImages(layout, 2);
  const raster = await pagesToPdf(layout.pages.map((pg, i) => ({ url: urls[i]!, w: pg.setup.width, h: pg.setup.height })));
  const base = { pages: layout.pages.length, lossy: doc.lossy };
  const text = layout.pages.map((pg, i) => ({ pageIndex: i, rotate: 0, lines: pageTextLines(pg) })).filter((p) => p.lines.length);
  if (!deps.addTextLayer || !text.length) return { ...base, bytes: raster, textPages: 0, problems: text.length ? [] : ['文档里没有文字'] };
  try {
    const r = await deps.addTextLayer(raster, text);
    if (r.bytes) return { ...base, bytes: r.bytes, textPages: text.length, problems: r.problems };
    return { ...base, bytes: raster, textPages: 0, problems: r.problems };
  } catch (e) {
    return { ...base, bytes: raster, textPages: 0, problems: [e instanceof Error ? e.message : String(e)] };
  }
}

import { PDFDocument } from 'pdf-lib';

/** 每页大小可以不同（Word 各节的纸张）：页面图片 + 页面大小（磅） → PDF 字节 */
export async function pagesToPdf(pages: { url: string; w: number; h: number }[]): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  for (const pg of pages) {
    const png = await pdf.embedPng(pg.url);
    pdf.addPage([pg.w, pg.h]).drawImage(png, { x: 0, y: 0, width: pg.w, height: pg.h });
  }
  return pdf.save();
}

/** 页面图片（PNG data URL）→ PDF 字节，每页按纸张大小 */
export async function imagesToPdf(urls: string[], w: number, h: number): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  for (const url of urls) {
    const png = await pdf.embedPng(url);
    const p = pdf.addPage([w, h]);
    p.drawImage(png, { x: 0, y: 0, width: w, height: h });
  }
  return pdf.save();
}

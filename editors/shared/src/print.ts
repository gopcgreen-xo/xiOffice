/**
 * 打印一组页面图片（data URL）：放进隐藏 iframe 后调用其 print()，
 * 这样只打印内容，不会把整个工作台界面打印出来。
 * pageSize 为 CSS @page size，如 '210mm 297mm'；省略时由打印机决定。
 */
export async function printImages(urls: string[], pageSize = 'auto'): Promise<void> {
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  Object.assign(iframe.style, { position: 'fixed', right: '0', bottom: '0', width: '0', height: '0', border: '0' });
  iframe.srcdoc = `<!doctype html><html><head><style>
@page { size: ${pageSize}; margin: 0; }
html, body { margin: 0; padding: 0; }
img { display: block; width: 100%; height: auto; break-after: page; page-break-after: always; }
img:last-child { break-after: auto; page-break-after: auto; }
</style></head><body>${urls.map((u) => `<img src="${u}">`).join('')}</body></html>`;
  document.body.appendChild(iframe);
  await new Promise<void>((resolve) => iframe.addEventListener('load', () => resolve(), { once: true }));
  const doc = iframe.contentDocument!;
  await Promise.all(Array.from(doc.images).map((img) => img.decode().catch(() => undefined)));
  iframe.contentWindow!.focus();
  iframe.contentWindow!.print();
  // print() 在多数 WebView 中是阻塞的；稍后移除 iframe
  setTimeout(() => iframe.remove(), 60_000);
}

/** 画布转 data URL 用于打印。 */
export function canvasToDataUrl(c: HTMLCanvasElement): string {
  return c.toDataURL('image/png');
}

/**
 * 文档里的嵌入式图片：字节放在这个模块级仓库里，模型（Run.image）只记 id 和尺寸，
 * 这样撤销快照（structuredClone）不会反复复制图片数据。同样内容的图片共用一个 id。
 */

interface Entry {
  mime: string;
  data: Uint8Array;
  /** 解码完成的图片（浏览器里才有） */
  el?: HTMLImageElement;
  loading?: Promise<void>;
}

const store = new Map<string, Entry>();

/** 浏览器能直接显示的图片类型（按扩展名）；EMF / WMF / TIFF 等不支持 */
const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
};

export function imageMime(fileName: string): string | null {
  return MIME_BY_EXT[fileName.split('.').pop()?.toLowerCase() ?? ''] ?? null;
}

/** 写回 docx 时的扩展名 */
export function imageExt(mime: string): string {
  return mime === 'image/jpeg' ? 'jpeg' : mime.slice('image/'.length);
}

/** FNV-1a：同样内容得到同一个 id */
function hash(data: Uint8Array): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < data.length; i++) {
    h ^= data[i];
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16);
}

export function registerImage(data: Uint8Array, mime: string): string {
  const id = `img-${data.length}-${hash(data)}`;
  if (!store.has(id)) store.set(id, { mime, data });
  return id;
}

export function imageData(id: string): { mime: string; data: Uint8Array } | undefined {
  return store.get(id);
}

/** 已解码的图片；还没解码（或解码失败）返回 undefined，调用方画占位框 */
export function imageElement(id: string): HTMLImageElement | undefined {
  return store.get(id)?.el;
}

/** 解码这些图片（打开文档后、首次绘制前调用）；解码失败的保持未加载状态 */
export function loadImages(ids: Iterable<string>): Promise<void> {
  const jobs: Promise<void>[] = [];
  for (const id of new Set(ids)) {
    const e = store.get(id);
    if (!e || e.el) continue;
    e.loading ??= new Promise<void>((resolve) => {
      const url = URL.createObjectURL(new Blob([e.data as unknown as BlobPart], { type: e.mime }));
      const img = new Image();
      const done = () => {
        URL.revokeObjectURL(url);
        e.loading = undefined;
        resolve();
      };
      img.onload = () => {
        e.el = img;
        done();
      };
      img.onerror = done;
      img.src = url;
    });
    jobs.push(e.loading);
  }
  return Promise.all(jobs).then(() => undefined);
}

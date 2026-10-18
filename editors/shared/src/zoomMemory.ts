/**
 * 编辑器的缩放比例按“上次用的”打开（localStorage，各窗口共用）：
 * 悬浮文档窗口里再开一次同一个文件时，缩放才和主窗口一样，框选的内容才对得上。Word 用自己的 `word.zoom`，做法相同。
 */

/** 读上次的缩放；没有或不合法时用 fallback */
export function loadZoom<T extends number | string>(key: string, fallback: T, ok: (v: unknown) => boolean): T {
  try {
    const raw = localStorage.getItem(key);
    if (raw === null) return fallback;
    const v: unknown = JSON.parse(raw);
    return ok(v) ? (v as T) : fallback;
  } catch {
    return fallback;
  }
}

/** 记下缩放（存不了只是下次按默认打开） */
export function saveZoom(key: string, v: number | string): void {
  try {
    localStorage.setItem(key, JSON.stringify(v));
  } catch {
    /* 无痕模式等情况下忽略 */
  }
}

/** 数字缩放是否在 [min, max] 里 */
export const zoomIn = (min: number, max: number) => (v: unknown) => typeof v === 'number' && v >= min && v <= max;

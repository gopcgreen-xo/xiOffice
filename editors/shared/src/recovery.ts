/**
 * 自动保存与崩溃恢复：编辑器有未保存的修改时每隔一段时间把内容存成草稿（IndexedDB，在工作台自己的数据里，
 * 不碰原文件）；正常保存或关闭标签后删除草稿。下次打开同一个文件时如果还有草稿，说明上次没保存就退出了，提示恢复。
 */

export interface Draft {
  /** 文件路径（小写） */
  key: string;
  /** 保存草稿的时间（毫秒） */
  time: number;
  /** 草稿内容：文件字节（docx / xlsx）或文字（csv） */
  data: Uint8Array | string;
}

/** 自动保存的间隔 */
export const AUTOSAVE_MS = 2 * 60 * 1000;

const DB = 'wz-recovery';
const STORE = 'drafts';

function open(): Promise<IDBDatabase | null> {
  return new Promise((resolve) => {
    try {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'key' });
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function tx<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T | null> {
  const db = await open();
  if (!db) return null;
  return new Promise((resolve) => {
    try {
      const req = fn(db.transaction(STORE, mode).objectStore(STORE));
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  }).finally(() => db.close()) as Promise<T | null>;
}

export const draftKey = (path: string) => path.replace(/\//g, '\\').toLowerCase();

/**
 * 关掉草稿（只读浏览的窗口用，如悬浮文档）：不读、不写、不删草稿。
 * 草稿存在各窗口共用的 IndexedDB 里，只读窗口打开 / 关闭同一个文件时不能去恢复或清掉主窗口的草稿。
 */
let disabled = false;
export function disableDrafts() {
  disabled = true;
}

export async function saveDraft(path: string, data: Uint8Array | string): Promise<void> {
  if (disabled) return;
  await tx('readwrite', (s) => s.put({ key: draftKey(path), time: Date.now(), data } satisfies Draft));
}

export async function loadDraft(path: string): Promise<Draft | null> {
  if (disabled) return null;
  return (await tx<Draft | undefined>('readonly', (s) => s.get(draftKey(path)))) ?? null;
}

export async function clearDraft(path: string): Promise<void> {
  if (disabled) return;
  await tx('readwrite', (s) => s.delete(draftKey(path)));
}

/** 恢复提示里的时间：今天显示时分，其它日子带日期 */
export function draftTime(t: number, now = new Date()): string {
  const d = new Date(t);
  const hm = `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  return d.toDateString() === now.toDateString() ? `今天 ${hm}` : `${d.getMonth() + 1} 月 ${d.getDate()} 日 ${hm}`;
}

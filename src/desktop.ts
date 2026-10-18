/** Electron 预加载脚本（electron/preload.cjs）暴露的能力；在普通浏览器里打开时没有 */
import type { DirEntry } from '@wz/editor-shared';

export interface DesktopApi {
  platform: string;
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, data: Uint8Array): Promise<void>;
  listDir(dir: string): Promise<DirEntry[]>;
  /** 写到文件夹里，同名时自动改名；返回实际文件名 */
  exportFile(dir: string, name: string, data: Uint8Array): Promise<string>;
  showOpen(opts: { title: string; exts?: string[]; multiple?: boolean }): Promise<string[]>;
  /** filters：[名称, 扩展名列表]；取消返回 null */
  showSave(opts: { title: string; defaultPath: string; filters: [string, string[]][] }): Promise<string | null>;
  pickFolder(title: string): Promise<string | null>;
  openPath(path: string): Promise<void>;
  openUrl(url: string): Promise<void>;
  fontFaces(): Promise<Record<string, string>[]>;
  info(): Promise<{ documents: string; version: string }>;
  takePendingFiles(): Promise<string[]>;
  setDirty(dirty: boolean): void;
  pathOf(file: File): string;
  onOpenFiles(cb: (paths: string[]) => void): () => void;
  onMenu(cb: (cmd: string) => void): () => void;
}

declare global {
  interface Window {
    xioffice?: DesktopApi;
  }
}

export const desktop: DesktopApi | undefined = window.xioffice;

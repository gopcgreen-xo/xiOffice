/** 开发页：不启动 Tauri，在浏览器里直接跑 Word 编辑器（`npm run dev:word`）；文件系统是内存里的假实现，只用来看版面和交互 */
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { setEditorHost, type EditorHost } from '@wz/editor-shared';
import '../../../src/styles/themes.css';
import '../../../src/styles/app.css';
import { Editor } from '../src';

const noop = () => undefined;
// ?file=<本机完整路径>：打开这个文件（经 Vite 的 /@fs/ 读取），看真实文档的版面
const file = new URLSearchParams(location.search).get('file');
const host = {
  readBytes: async (p: string) => new Uint8Array(await (await fetch(`/@fs/${p.split('\\').join('/')}`)).arrayBuffer()),
  readText: async () => ({ content: '', encoding: 'UTF-8' }),
  writeBytes: async () => undefined,
  writeText: async () => undefined,
  listDir: async () => [],
  readXls: async () => [],
  openWithSystem: async () => undefined,
  openUrl: async () => undefined,
  workFolders: () => ['C:\dev'],
  toast: (m: string) => console.log('[toast]', m),
  askText: async (_t: string, v?: string) => v ?? '',
  confirm: async () => true,
  guard: async <T,>(p: Promise<T>) => p,
  pushPopup: noop,
  popPopup: noop,
  updateTab: noop,
  printImages: async () => undefined,
  log: noop,
  useWorkFolders: () => ['C:\dev'],
  useFsEvent: () => ({ tick: 0, dirs: [] }),
  useUserName: () => '开发者',
} as unknown as EditorHost;
setEditorHost(host);

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <Editor tab={file ? { id: 'dev', title: file.split(/[\\/]/).pop()!, path: file } : { id: 'dev', title: '开发页.docx' }} />
  </StrictMode>,
);

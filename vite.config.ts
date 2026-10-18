import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';
import type { Plugin } from 'vite';
import { fileURLToPath } from 'url';

const at = (p: string) => fileURLToPath(new URL(p, import.meta.url));

/** 生产构建的内容安全策略：只运行自带的脚本、不连外网（开发服务器的热更新需要内联脚本，不加） */
const CSP = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "font-src 'self' data:",
  "connect-src 'self' data: blob:",
  "worker-src 'self' blob:",
  "object-src 'none'",
].join('; ');

const csp = (): Plugin => ({
  name: 'xioffice-csp',
  apply: 'build',
  transformIndexHtml: (html) => html.replace('<meta charset="utf-8" />', `<meta charset="utf-8" />\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`),
});

export default defineConfig({
  plugins: [react(), csp()],
  // 编辑器公共件按包名引用
  resolve: {
    alias: [
      { find: /^@wz\/editor-shared$/, replacement: at('./editors/shared/src/index.ts') },
      { find: /^@wz\/editor-shared\/(.*)$/, replacement: at('./editors/shared/src/') + '$1' },
    ],
  },
  // 相对路径：Electron 从 app:// 加载，也能部署到任意子路径
  base: './',
  server: { port: 1430, watch: { ignored: ['**/editors/word/tools/data/**'] } },
  build: { target: 'es2022', chunkSizeWarningLimit: 4000 },
  // 表格的解析 Worker 用 ES 模块格式
  worker: { format: 'es' },
  test: {
    environment: 'node',
    include: ['src/**/*.test.ts', 'editors/{word,sheet,shared}/**/*.test.ts'],
    // CI 的 2 核机器上并行跑压力测试（如 2000 层嵌套表格）会超过默认的 5 秒
    testTimeout: 30_000,
  },
});

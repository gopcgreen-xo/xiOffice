import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';

// 开发页：单独的 Vite 服务（端口 1431），在普通浏览器里跑 Word 编辑器看版面和交互
export default defineConfig({
  root: __dirname,
  plugins: [react()],
  resolve: { alias: [{ find: /^@wz\/editor-shared\/(.*)$/, replacement: resolve(__dirname, '../../shared/src/') + '/$1' }] },
  server: { port: 1431, strictPort: true, fs: { allow: [resolve(__dirname, '../../..')] } },
});

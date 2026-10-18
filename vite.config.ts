import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'url';

const at = (p: string) => fileURLToPath(new URL(p, import.meta.url));

export default defineConfig({
  plugins: [react()],
  // 编辑器公共件按包名引用
  resolve: {
    alias: [
      { find: /^@wz\/editor-shared$/, replacement: at('./editors/shared/src/index.ts') },
      { find: /^@wz\/editor-shared\/(.*)$/, replacement: at('./editors/shared/src/') + '$1' },
    ],
  },
  // GitHub Pages 等子路径部署：BASE=/xiOffice/ npm run build
  base: process.env.BASE ?? '/',
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

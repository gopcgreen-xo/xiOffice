// 开发：启动 Vite 开发服务器，再用它的地址启动 Electron；关掉窗口时一起退出
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { createServer } from 'vite';

const require = createRequire(import.meta.url);
/** electron 包导出的是可执行文件路径 */
const electron = require('electron');

const server = await createServer();
await server.listen();
const url = server.resolvedUrls.local[0];
console.log(`开发服务器：${url}`);

const child = spawn(electron, ['.'], { stdio: 'inherit', env: { ...process.env, XIOFFICE_DEV_URL: url } });
child.on('exit', async (code) => {
  await server.close();
  process.exit(code ?? 0);
});

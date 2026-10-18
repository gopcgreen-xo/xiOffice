/** 在 Web Worker 里解析工作簿，大文件不卡界面 */
import ParseWorker from './parse.worker?worker';
import type { ParseRequest, ParseResponse } from './parse.worker';
import type { Workbook } from './workbook';

export function parseInWorker(req: ParseRequest): Promise<Workbook> {
  return new Promise((resolve, reject) => {
    const w = new ParseWorker();
    w.onmessage = (e: MessageEvent<ParseResponse>) => {
      w.terminate();
      if (e.data.ok) resolve(e.data.book);
      else reject(new Error(e.data.error));
    };
    w.onerror = (e) => {
      w.terminate();
      reject(new Error(e.message || '解析失败'));
    };
    w.postMessage(req, req.kind === 'xlsx' ? [req.data.buffer as ArrayBuffer] : []);
  });
}

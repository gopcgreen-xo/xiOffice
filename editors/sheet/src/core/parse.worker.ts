import { readCsv, readWorkbook, type Workbook } from './workbook';

export type ParseRequest = { kind: 'xlsx'; data: Uint8Array } | { kind: 'csv'; text: string; name: string };
export type ParseResponse = { ok: true; book: Workbook } | { ok: false; error: string };

self.onmessage = (e: MessageEvent<ParseRequest>) => {
  const req = e.data;
  let res: ParseResponse;
  try {
    res = { ok: true, book: req.kind === 'xlsx' ? readWorkbook(req.data) : readCsv(req.text, req.name) };
  } catch (err) {
    res = { ok: false, error: err instanceof Error ? err.message : String(err) };
  }
  self.postMessage(res);
};

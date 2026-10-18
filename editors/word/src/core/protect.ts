/**
 * 文档保护（限制编辑）：只读 / 只允许批注 / 只允许修订（修订不能关、不能接受拒绝），可选密码；
 * “允许编辑的区域”（Word 的例外项：permStart / permEnd，所有人）在只读或只批注时仍可以改。
 * 密码按 Word 2013+ 的方式（SHA-512 迭代，docx-engine 的 protection.ts）保存和校验。
 * 还没有支持的部分：按用户 / 用户组区分的例外项（只认“所有人”）、窗体保护下填写窗体（按只读处理）、
 * 页眉页脚和表格外之外的例外区域（只认正文）。
 */
import { hashProtectionPassword, verifyProtectionPassword } from '../../vendor/docx-engine/src/protection';
import type { DocProtection } from '../../vendor/docx-engine/src/types';
import { FIELD_CHAR, isAtom, type Doc, type Opaque, type Run } from './model';
import { comparePos, type Pos } from './layout';
import { getPara, paragraphPaths, runPropsAt } from './ops';

export type { DocProtection };
export type ProtectMode = 'readOnly' | 'comments' | 'trackedChanges' | 'forms';

export const PROTECT_LABEL: Record<string, string> = {
  readOnly: '只读',
  comments: '只能批注',
  trackedChanges: '只能修订',
  forms: '只能填写窗体',
};

/** 生效的保护（没有、没启用返回 null） */
export const activeProtection = (doc: Doc): DocProtection | null => (doc.protection?.enforced ? doc.protection : null);

export function protectLabel(p: DocProtection): string {
  return PROTECT_LABEL[p.edit] ?? p.edit;
}

/** 设置保护：有密码时按 Word 的方式算哈希（耗时约百毫秒） */
export async function makeProtection(mode: ProtectMode, password?: string, spinCount?: number): Promise<DocProtection> {
  if (!password) return { edit: mode, enforced: true };
  const h = await hashProtectionPassword(password, spinCount);
  return { edit: mode, enforced: true, hash: h.hash, salt: h.salt, spinCount: h.spinCount, algorithmSid: h.algorithmSid };
}

/** 停止保护要不要输入密码 */
export const needsPassword = (p: DocProtection): boolean => !!p.hash;

export async function checkPassword(p: DocProtection, password: string): Promise<boolean> {
  return verifyProtectionPassword(password, p);
}

// ---------------- 允许编辑的区域 ----------------

const permMark = (r: Run): { kind: 'start' | 'end'; id: string; everyone: boolean } | null => {
  const o = r.opaque;
  if (!o || o.level !== 'para') return null;
  const start = /^<w:permStart\b/.test(o.xml);
  const end = /^<w:permEnd\b/.test(o.xml);
  if (!start && !end) return null;
  const id = o.xml.match(/\bw:id="([^"]*)"/)?.[1] ?? '';
  // 没写 edGrp / ed 的按“所有人”；指定了具体用户的不算
  const everyone = start ? /\bw:edGrp="everyone"/.test(o.xml) || !/\bw:ed="/.test(o.xml) : true;
  return { kind: start ? 'start' : 'end', id, everyone };
};

export interface PermRange {
  id: string;
  start: Pos;
  end: Pos;
}

/** 正文里允许所有人编辑的区域（文档顺序） */
export function permRanges(doc: Doc): PermRange[] {
  const starts = new Map<string, Pos>();
  const out: PermRange[] = [];
  for (const path of paragraphPaths(doc, 'body')) {
    let acc = 0;
    for (const r of getPara(doc, { story: 'body', path }).runs) {
      const m = permMark(r);
      if (m?.kind === 'start' && m.everyone) starts.set(m.id, { story: 'body', path, offset: acc });
      else if (m?.kind === 'end' && starts.has(m.id)) {
        out.push({ id: m.id, start: starts.get(m.id)!, end: { story: 'body', path, offset: acc } });
        starts.delete(m.id);
      }
      acc += r.text.length;
    }
  }
  return out;
}

/** 选区（两端）整个落在同一个允许编辑的区域里 */
export function inPerm(doc: Doc, a: Pos, b: Pos): boolean {
  if (a.story !== 'body' || b.story !== 'body') return false;
  const [s, e] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  return permRanges(doc).some((r) => comparePos(r.start, s) <= 0 && comparePos(e, r.end) <= 0);
}

const nextPermId = (doc: Doc): number => {
  let max = 19999;
  for (const path of paragraphPaths(doc, 'body')) for (const r of getPara(doc, { story: 'body', path }).runs) {
    const m = permMark(r);
    if (m) max = Math.max(max, Number(m.id) || 0);
  }
  return max + 1;
};

const permRun = (xml: string, like: Run['props']): Run => ({ text: FIELD_CHAR, props: { ...like }, opaque: { xml, level: 'para', kind: 'other', w: 0, h: 0 } as Opaque });

/** 把选区标成允许编辑的区域 */
export function addPermRange(doc: Doc, a: Pos, b: Pos): void {
  if (a.story !== 'body' || b.story !== 'body') throw new Error('允许编辑的区域只能在正文里');
  [a, b] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  if (comparePos(a, b) === 0) throw new Error('请先选中要允许编辑的文字');
  const id = nextPermId(doc);
  const like = runPropsAt(doc, a);
  put(doc, b, permRun(`<w:permEnd w:id="${id}"/>`, like));
  put(doc, a, permRun(`<w:permStart w:id="${id}" w:edGrp="everyone"/>`, like));
}

function put(doc: Doc, pos: Pos, run: Run): void {
  const p = getPara(doc, pos);
  let acc = 0;
  let i = 0;
  for (; i < p.runs.length; i++) {
    const r = p.runs[i];
    if (acc === pos.offset) break;
    if (pos.offset < acc + r.text.length && !isAtom(r)) {
      const k = pos.offset - acc;
      p.runs.splice(i, 1, { ...r, text: r.text.slice(0, k) }, { ...r, props: { ...r.props }, text: r.text.slice(k) });
      i++;
      break;
    }
    acc += r.text.length;
  }
  p.runs.splice(i, 0, run);
}

/** 去掉一个允许编辑的区域（两端的标记） */
export function removePermRange(doc: Doc, id: string): void {
  for (const path of paragraphPaths(doc, 'body')) {
    const p = getPara(doc, { story: 'body', path });
    p.runs = p.runs.filter((r) => permMark(r)?.id !== id);
  }
}

export const clearPermRanges = (doc: Doc): void => permRanges(doc).forEach((r) => removePermRange(doc, r.id));

// ---------------- 编辑检查 ----------------

/** mutate 的各种操作归成几类 */
export type EditClass = 'content' | 'comment' | 'review' | 'document';

const DOCUMENT_KINDS = new Set(['style', 'watermark', 'page', 'setup', 'hf-link', 'section', 'toc', 'protect']);

export function editClass(kind: string): EditClass {
  if (kind === 'comment') return 'comment';
  if (kind === 'review') return 'review';
  return DOCUMENT_KINDS.has(kind) ? 'document' : 'content';
}

/**
 * 这次操作被保护挡住了吗：返回提示文字，没挡住返回 null。
 * 只读 / 只批注：只有允许编辑的区域里能改（批注模式下任何位置能加批注）；
 * 只修订：什么都能改（由编辑器强制打开修订），但不能接受 / 拒绝修订。文档级的设置（样式、水印、页面…）只有只修订时放行。
 */
export function editError(doc: Doc, kind: string, a: Pos, b: Pos): string | null {
  const p = activeProtection(doc);
  if (!p) return null;
  const hint = `文档受保护（${protectLabel(p)}），要修改请先“停止保护”`;
  const k = editClass(kind);
  if (p.edit === 'trackedChanges') return k === 'review' ? hint : null;
  if (k === 'comment') return p.edit === 'comments' || inPerm(doc, a, b) ? null : hint;
  if (k === 'content') return inPerm(doc, a, b) ? null : hint;
  return hint;
}

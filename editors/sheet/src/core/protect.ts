/**
 * 保护设置：工作表保护（可选密码、允许用户做的操作）、单元格锁定、工作簿结构保护。
 *
 * - 保护的数据是 `<sheetProtection>` / `<workbookProtection>` 的属性（`protection.attrs`，原样保留，没改的属性不动）。
 * - 密码：写文件时用 Excel 认识的旧式 16 位散列（`password="CC1A"`）；读到新式（SHA-512 加盐迭代，Excel 2013 起）的
 *   也能校验，取消保护时要输入正确的密码。密码不是加密，只是防误改（和 Excel 一样）。
 * - 本编辑器里受保护的工作表只能改“未锁定”的单元格内容（比文件里允许的更严格），见 ops.ts `violatesProtection`。
 */
import { applyStyle, type Tx } from './ops';
import { sha384, sha512 } from './sha512';
import type { Range, SheetData, Workbook } from './workbook';

// ---------- 密码 ----------

/** 旧式密码散列（ECMA-376 Part 4 §14.7.1）：16 位，写成 4 位十六进制大写 */
export function legacyHash(pw: string): string {
  if (!pw) return '';
  let h = 0;
  for (let i = pw.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= pw.charCodeAt(i);
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  h ^= pw.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
}

const b64 = (b: Uint8Array) => btoa(String.fromCharCode(...b));
const unb64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const ALGO: Record<string, string> = { 'SHA-512': 'SHA-512', 'SHA-384': 'SHA-384', 'SHA-256': 'SHA-256', 'SHA-1': 'SHA-1', SHA1: 'SHA-1', SHA512: 'SHA-512', SHA256: 'SHA-256' };

/** 校验密码：没设密码返回 true；用了不支持的算法返回 null（无法校验） */
export async function verifyPassword(attrs: Record<string, string>, pw: string, prefix = ''): Promise<boolean | null> {
  // 工作簿保护的属性名带 workbook 前缀（workbookPassword / workbookHashValue …）
  const key = (n: string) => (prefix ? prefix + n[0].toUpperCase() + n.slice(1) : n);
  const legacy = attrs[key('password')];
  const hash = attrs[key('hashValue')];
  if (hash) {
    const algo = ALGO[(attrs[key('algorithmName')] ?? '').toUpperCase()];
    if (!algo) return null;
    const salt = unb64(attrs[key('saltValue')] ?? '');
    const spin = Number(attrs[key('spinCount')] ?? 0);
    const pwBytes = new Uint8Array(pw.length * 2);
    for (let i = 0; i < pw.length; i++) {
      pwBytes[i * 2] = pw.charCodeAt(i) & 0xff;
      pwBytes[i * 2 + 1] = pw.charCodeAt(i) >> 8;
    }
    const concat = (...parts: Uint8Array[]) => {
      const all = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
      let at = 0;
      for (const p of parts) {
        all.set(p, at);
        at += p.length;
      }
      return all;
    };
    const idx = (i: number) => new Uint8Array([i & 0xff, (i >> 8) & 0xff, (i >> 16) & 0xff, (i >> 24) & 0xff]);
    // Excel 用的 SHA-512 / SHA-384 走同步实现（迭代 10 万次，异步接口太慢）；其它算法用 WebCrypto
    const sync = algo === 'SHA-512' ? sha512 : algo === 'SHA-384' ? sha384 : null;
    let h: Uint8Array;
    if (sync) {
      h = sync(concat(salt, pwBytes));
      for (let i = 0; i < spin; i++) h = sync(concat(h, idx(i)));
    } else {
      h = new Uint8Array(await crypto.subtle.digest(algo, concat(salt, pwBytes)));
      for (let i = 0; i < spin; i++) h = new Uint8Array(await crypto.subtle.digest(algo, concat(h, idx(i))));
    }
    return b64(h) === hash;
  }
  if (legacy) return legacy.toUpperCase() === legacyHash(pw);
  return true;
}

/** 有没有设置密码 */
export const hasPassword = (attrs: Record<string, string> | undefined, prefix = ''): boolean => {
  const key = (n: string) => (prefix ? prefix + n[0].toUpperCase() + n.slice(1) : n);
  return !!(attrs && (attrs[key('hashValue')] || attrs[key('password')]));
};

// ---------- 工作表保护的选项 ----------

/** 允许用户做的操作（保护后）。值 true = 允许。名字就是 `<sheetProtection>` 的属性名 */
export interface SheetAllow {
  selectLockedCells: boolean;
  selectUnlockedCells: boolean;
  formatCells: boolean;
  formatColumns: boolean;
  formatRows: boolean;
  insertColumns: boolean;
  insertRows: boolean;
  insertHyperlinks: boolean;
  deleteColumns: boolean;
  deleteRows: boolean;
  sort: boolean;
  autoFilter: boolean;
  pivotTables: boolean;
  objects: boolean;
}

/** 属性缺省时的值（true = 该操作被禁止）。Excel 保护工作表时默认只允许选定单元格 */
const LOCKED_BY_DEFAULT: Record<keyof SheetAllow, boolean> = {
  objects: false,
  formatCells: true,
  formatColumns: true,
  formatRows: true,
  insertColumns: true,
  insertRows: true,
  insertHyperlinks: true,
  deleteColumns: true,
  deleteRows: true,
  selectLockedCells: false,
  sort: true,
  autoFilter: true,
  pivotTables: true,
  selectUnlockedCells: false,
};

export const ALLOW_KEYS = Object.keys(LOCKED_BY_DEFAULT) as (keyof SheetAllow)[];

const flag = (v: string | undefined, dflt: boolean) => (v === undefined ? dflt : v === '1' || v === 'true');

/** 读出“允许用户做的操作” */
export function allowOf(attrs: Record<string, string> | undefined): SheetAllow {
  const out = {} as SheetAllow;
  for (const k of ALLOW_KEYS) out[k] = !flag(attrs?.[k], LOCKED_BY_DEFAULT[k]);
  return out;
}

/** Excel 保护工作表对话框的默认勾选：只允许选定锁定和未锁定的单元格 */
export const DEFAULT_ALLOW: SheetAllow = { ...allowOf({}), objects: false };

// ---------- 修改 ----------

/** 保护工作表的属性。password 为空 = 不设密码；allow 为允许的操作。保留原来属性里不认识的部分 */
export function protectSheetAttrs(prev: Record<string, string> | undefined, allow: SheetAllow, password: string): Record<string, string> {
  const out: Record<string, string> = {};
  // 密码属性放最前面（顺序同 Excel）
  if (password) out.password = legacyHash(password);
  out.sheet = '1';
  for (const k of ALLOW_KEYS) {
    // 只写和缺省值不同的；objects / scenarios 缺省是“不禁止”，但 Excel 保护时一定写 1（禁止编辑对象和方案）
    const locked = !allow[k];
    if (locked !== LOCKED_BY_DEFAULT[k] || (k === 'objects' && locked)) out[k] = locked ? '1' : '0';
    if (k === 'objects') out.scenarios = '1';
  }
  // 其它不认识的属性（如 algorithmName 以外的扩展）沿用
  for (const [k, v] of Object.entries(prev ?? {})) {
    if (!(k in out) && !['password', 'hashValue', 'saltValue', 'spinCount', 'algorithmName', 'sheet', ...ALLOW_KEYS, 'scenarios'].includes(k)) out[k] = v;
  }
  return out;
}

/** 保护（attrs 由 protectSheetAttrs 算好）/ 取消保护（attrs = null） */
export function setSheetProtection(tx: Tx, si: number, attrs: Record<string, string> | null): void {
  tx.bypass = true;
  tx.meta(si);
  const s: SheetData = tx.book.sheets[si];
  if (attrs) {
    s.protection = { attrs };
    s.protected = true;
  } else {
    delete s.protection;
    delete s.protected;
  }
}

/** 选中的单元格锁定 / 取消锁定（保护工作表后才有效果） */
export function setCellsLocked(tx: Tx, si: number, g: Range, locked: boolean): void {
  applyStyle(tx, si, g, { locked });
}

/** 工作簿结构保护：不能增删、改名、移动、隐藏工作表 */
export function setStructureLock(tx: Tx, lock: false | { password: string }): void {
  const book = tx.book;
  tx.wholeBook();
  book.namesDirty = true;
  if (!lock) delete book.wbProtection;
  else {
    const attrs: Record<string, string> = {};
    if (lock.password) attrs.workbookPassword = legacyHash(lock.password);
    attrs.lockStructure = '1';
    book.wbProtection = { attrs };
  }
}

export const structureLocked = (book: Workbook): boolean => !!book.wbProtection && flag(book.wbProtection.attrs.lockStructure, false);

/** 单元格是否锁定（缺省锁定）；用于界面上显示锁定状态 */
export function isLocked(book: Workbook, sheet: SheetData, r: number, c: number): boolean {
  const s = sheet.cells[r]?.[c]?.s ?? sheet.blankStyles?.[r]?.[c] ?? sheet.rowStyles?.[r] ?? sheet.colStyles?.[c];
  const xf = book.styleBook?.xfs[s ?? 0];
  return !/<protection\b[^>]*\blocked="(0|false)"/.test(xf?.inner ?? '');
}

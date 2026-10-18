/**
 * 加密工作包的钥匙（`meta/keys.json`，不加密；格式见 docs/WBP_FORMAT.md 7.2）。
 * 包里的部件用一个随机生成的“解压密钥”加密（WinZip AES-256，解压软件打开时输入它）；
 * 这个密钥再分别用管理密码、每个角色的密码、恢复码各包一份（PBKDF2-SHA256 → AES-GCM），
 * 所以任何一个能登录的密码或恢复码都能解开密钥，忘了密码可以用恢复码进入后重设。
 */
import { unzipSync } from 'fflate';
import type { Who } from './access';

export const KEYS_INNER = 'meta/keys.json';

export interface Wrap {
  /** admin / recovery / role:<角色 id> */
  who: string;
  salt: string;
  iv: string;
  iter: number;
  data: string;
}

export interface KeysFile {
  v: 1;
  wraps: Wrap[];
}

const ITER = 120_000;
const b64 = (u: Uint8Array): string => btoa(String.fromCharCode(...u));
const unb64 = (s: string): Uint8Array => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
/** 去掉容易认错的 0 O 1 I L */
const ALPHA = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';

/** 随机码：6 组 × 4 个字符（约 119 位），如 K7QM-2XPA-9RTV-HC4N-8WJD-3FBE。解压密钥和恢复码都用这种格式 */
export function randomCode(): string {
  const bytes = crypto.getRandomValues(new Uint8Array(24));
  const chars = Array.from(bytes, (b) => ALPHA[b % ALPHA.length]).join('');
  return chars.match(/.{4}/g)!.join('-');
}

/** 恢复码输入时不区分大小写、空格和横线可有可无 */
export const normalizeCode = (s: string): string => {
  const c = s.toUpperCase().replace(/[^A-Z0-9]/g, '');
  return c.length === 24 ? c.match(/.{4}/g)!.join('-') : s.trim();
};

async function kek(password: string, salt: Uint8Array, iter: number): Promise<CryptoKey> {
  const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', hash: 'SHA-256', salt: salt as BufferSource, iterations: iter }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}

/** 用密码包住解压密钥 */
export async function wrapKey(packKey: string, password: string, who: string, iter = ITER): Promise<Wrap> {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const k = await kek(password, salt, iter);
  const data = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, k, new TextEncoder().encode(packKey)));
  return { who, salt: b64(salt), iv: b64(iv), iter, data: b64(data) };
}

/** 用密码解开；密码不对返回 null（AES-GCM 校验失败） */
export async function unwrapKey(w: Wrap, password: string): Promise<string | null> {
  try {
    const k = await kek(password, unb64(w.salt), w.iter);
    const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(w.iv) as BufferSource }, k, unb64(w.data) as BufferSource);
    return new TextDecoder().decode(plain);
  } catch {
    return null;
  }
}

export const whoOfWrap = (w: string): Who | null => (w === 'admin' ? { kind: 'admin' } : w.startsWith('role:') ? { kind: 'role', id: w.slice(5) } : null);
export const wrapName = (who: Who): string => (who.kind === 'admin' ? 'admin' : who.kind === 'role' ? `role:${who.id}` : 'base');

/** 用登录密码解开：返回解压密钥和身份；都不对返回 null */
export async function unlockWith(keys: KeysFile, password: string): Promise<{ key: string; who: Who } | null> {
  for (const w of keys.wraps) {
    const who = whoOfWrap(w.who);
    if (!who) continue;
    const key = await unwrapKey(w, password);
    if (key) return { key, who };
  }
  return null;
}

/** 用恢复码解开（进入后是管理身份，应马上重设管理密码） */
export async function unlockWithRecovery(keys: KeysFile, code: string): Promise<string | null> {
  const w = keys.wraps.find((x) => x.who === 'recovery');
  return w ? unwrapKey(w, normalizeCode(code)) : null;
}

export function parseKeys(text: string | undefined | null): KeysFile | null {
  if (!text) return null;
  try {
    const raw = JSON.parse(text) as { wraps?: unknown };
    const wraps = (Array.isArray(raw.wraps) ? raw.wraps : []).filter(
      (w): w is Wrap => !!w && typeof w === 'object' && ['who', 'salt', 'iv', 'data'].every((k) => typeof (w as Record<string, unknown>)[k] === 'string') && Number.isInteger((w as Wrap).iter),
    );
    return { v: 1, wraps };
  } catch {
    return { v: 1, wraps: [] };
  }
}

export const keysText = (k: KeysFile): string => JSON.stringify(k, null, 2) + '\n';

/** 包字节是不是加密的工作包（含不加密的 meta/keys.json）；是的话返回钥匙文件 */
export function encryptedKeys(bytes: Uint8Array): KeysFile | null {
  try {
    const parts = unzipSync(bytes, { filter: (f) => f.name === KEYS_INNER });
    const k = parts[KEYS_INNER];
    return k ? parseKeys(new TextDecoder().decode(k)) : null;
  } catch {
    return null;
  }
}

/** 打开加密包时缺少密钥：界面据此显示密码登录 */
export class PackLockedError extends Error {
  constructor(
    readonly pack: string,
    readonly keys: KeysFile,
  ) {
    super('这个工作包已加密，请先输入密码');
  }
}

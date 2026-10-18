/**
 * 工作包的访问权限（`meta/access.json`，格式见 docs/WBP_FORMAT.md）：
 *   · 管理密码：能改权限设置、增删改包里的文件；
 *   · 角色：每个角色一个密码，登录时只输密码（不输账号）；
 *   · 基础权限：不输密码进入时的权限。
 * 每个角色（含基础权限）有一个默认级别，可以对单个文件另设：不可见 / 可见（只看到名字）/ 可读（打开、导出，不能保存）/ 可写。
 * 包里没有权限设置时谁都能完全使用（和以前一样）。
 *
 * 注意：这是本程序里的访问控制，不是加密——.wbp 是 zip，用解压软件仍能看到里面的文件。
 * 登录状态只在当前窗口内存里（每个窗口一份），关掉窗口要重新输入。
 */
import type { Bundle } from './container';

export const ACCESS_INNER = 'meta/access.json';

export type Level = 'hidden' | 'visible' | 'read' | 'write';
export const LEVELS: Level[] = ['hidden', 'visible', 'read', 'write'];
export const LEVEL_LABEL: Record<Level, string> = { hidden: '不可见', visible: '可见', read: '可读', write: '可写' };
const rank = (l: Level): number => LEVELS.indexOf(l);
export const atLeast = (l: Level, want: Level): boolean => rank(l) >= rank(want);

export interface Secret {
  salt: string;
  hash: string;
  iter: number;
}

export interface AccessRule {
  /** 默认级别 */
  level: Level;
  /** 条目 id → 级别（没写的用默认级别） */
  items?: Record<string, Level>;
}

export interface AccessRole extends AccessRule {
  id: string;
  name: string;
  secret: Secret;
}

export interface AccessFile {
  v: 1;
  admin: Secret;
  base: AccessRule;
  roles: AccessRole[];
}

/** 登录身份：管理员 / 某个角色 id / 基础权限 */
export type Who = { kind: 'admin' } | { kind: 'role'; id: string } | { kind: 'base' };

// ---------------------------------------------------------------- 密码

const ITER = 120_000;
const b64 = (u: Uint8Array): string => btoa(String.fromCharCode(...u));
const unb64 = (s: string): Uint8Array => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

async function derive(password: string, salt: Uint8Array, iter: number): Promise<string> {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: salt as BufferSource, iterations: iter }, key, 256);
  return b64(new Uint8Array(bits));
}

/** 加盐哈希一个密码（PBKDF2-SHA256） */
export async function makeSecret(password: string, iter = ITER): Promise<Secret> {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  return { salt: b64(salt), hash: await derive(password, salt, iter), iter };
}

export async function checkSecret(password: string, s: Secret): Promise<boolean> {
  try {
    const h = await derive(password, unb64(s.salt), s.iter);
    // 定长比较
    if (h.length !== s.hash.length) return false;
    let d = 0;
    for (let i = 0; i < h.length; i++) d |= h.charCodeAt(i) ^ s.hash.charCodeAt(i);
    return d === 0;
  } catch {
    return false;
  }
}

/** 用密码登录：返回身份；密码不对返回 null */
export async function loginWith(file: AccessFile, password: string): Promise<Who | null> {
  if (await checkSecret(password, file.admin)) return { kind: 'admin' };
  for (const r of file.roles) if (await checkSecret(password, r.secret)) return { kind: 'role', id: r.id };
  return null;
}

/** 这个密码是否已被管理员或别的角色用了（设置角色密码时防重复；except = 正在改的角色 id） */
export async function passwordTaken(file: AccessFile, password: string, except?: string): Promise<boolean> {
  if (await checkSecret(password, file.admin)) return true;
  for (const r of file.roles) if (r.id !== except && (await checkSecret(password, r.secret))) return true;
  return false;
}

// ---------------------------------------------------------------- 读写设置

const isLevel = (v: unknown): v is Level => typeof v === 'string' && (LEVELS as string[]).includes(v);
const cleanSecret = (v: unknown): Secret | null => {
  const o = v as Record<string, unknown> | null;
  if (!o || typeof o !== 'object' || typeof o.salt !== 'string' || typeof o.hash !== 'string') return null;
  const iter = Number(o.iter);
  return { salt: o.salt, hash: o.hash, iter: Number.isInteger(iter) && iter >= 1000 && iter <= 10_000_000 ? iter : ITER };
};
const cleanRule = (v: unknown, def: Level): AccessRule => {
  const o = (v && typeof v === 'object' ? v : {}) as Record<string, unknown>;
  const items: Record<string, Level> = {};
  if (o.items && typeof o.items === 'object') for (const [k, l] of Object.entries(o.items as object)) if (isLevel(l) && /^[A-Za-z0-9_-]{1,64}$/.test(k)) items[k] = l;
  return { level: isLevel(o.level) ? o.level : def, ...(Object.keys(items).length ? { items } : {}) };
};

/** 解析权限设置；坏掉的设置当作“谁都只能看到名字”（宁可严，不可松），没有管理密码的设置无效（返回 null = 不限制） */
export function parseAccess(text: string | undefined): AccessFile | null {
  if (!text) return null;
  let raw: Record<string, unknown>;
  try {
    raw = JSON.parse(text) as Record<string, unknown>;
  } catch {
    return { v: 1, admin: { salt: '', hash: '!', iter: ITER }, base: { level: 'visible' }, roles: [] };
  }
  const admin = cleanSecret(raw.admin);
  if (!admin) return null;
  const roles: AccessRole[] = [];
  for (const r of Array.isArray(raw.roles) ? raw.roles : []) {
    const o = r as Record<string, unknown>;
    const secret = cleanSecret(o?.secret);
    if (!secret || typeof o.id !== 'string' || !o.id) continue;
    roles.push({ id: o.id.slice(0, 40), name: typeof o.name === 'string' && o.name.trim() ? o.name.trim().slice(0, 40) : '角色', secret, ...cleanRule(o, 'read') });
  }
  return { v: 1, admin, base: cleanRule(raw.base, 'visible'), roles };
}

export const readAccess = (b: Bundle): AccessFile | null => parseAccess(b.readText(ACCESS_INNER));

export function writeAccess(b: Bundle, file: AccessFile | null): void {
  if (file) b.writeText(ACCESS_INNER, JSON.stringify(file, null, 2) + '\n');
  else b.remove(ACCESS_INNER);
}

// ---------------------------------------------------------------- 级别

/** 某个身份对某个条目的级别 */
export function levelFor(file: AccessFile | null, who: Who | null, itemId: string): Level {
  if (!file) return 'write';
  if (!who) return 'hidden';
  if (who.kind === 'admin') return 'write';
  const rule = who.kind === 'base' ? file.base : file.roles.find((r) => r.id === who.id);
  if (!rule) return 'hidden';
  return rule.items?.[itemId] ?? rule.level;
}

export const whoLabel = (file: AccessFile | null, who: Who | null): string =>
  !file ? '' : !who ? '未登录' : who.kind === 'admin' ? '管理' : who.kind === 'base' ? '基础权限' : (file.roles.find((r) => r.id === who.id)?.name ?? '角色');

// ---------------------------------------------------------------- 当前窗口的登录状态

const sessions = new Map<string, Who>();
const listeners = new Set<() => void>();
const notify = () => listeners.forEach((cb) => cb());

export function setAccessSession(pack: string, who: Who | null): void {
  if (who) sessions.set(pack, who);
  else sessions.delete(pack);
  notify();
}
export const accessSession = (pack: string): Who | null => sessions.get(pack) ?? null;
export function onAccessChange(cb: () => void): () => void {
  listeners.add(cb);
  return () => void listeners.delete(cb);
}

/** 当前窗口里这个包的权限视图（b 为已读入的包） */
export interface PackAccess {
  /** 包里有没有权限设置 */
  configured: boolean;
  who: Who | null;
  /** 能不能管理（改设置、增删改文件）：没有设置时谁都能 */
  manage: boolean;
  level(itemId: string): Level;
  label: string;
}

export function packAccess(b: Bundle, pack: string): PackAccess {
  const file = readAccess(b);
  const who = accessSession(pack);
  return {
    configured: !!file,
    who,
    manage: !file || who?.kind === 'admin',
    level: (id) => levelFor(file, who, id),
    label: whoLabel(file, who),
  };
}

/** 没有管理权限时抛错（增删改文件、改设置之前调用） */
export function requireManage(b: Bundle, pack: string): void {
  if (!packAccess(b, pack).manage) throw new Error('这个工作包设置了权限，只有用管理密码登录才能修改');
}

/** 级别不够时抛错 */
export function requireLevel(b: Bundle, pack: string, itemId: string, want: Level): void {
  const l = packAccess(b, pack).level(itemId);
  if (!atLeast(l, want)) throw new Error(want === 'write' ? '只读：没有修改这个文件的权限' : '没有打开这个文件的权限');
}

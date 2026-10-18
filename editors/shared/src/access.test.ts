import { strFromU8, strToU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  ACCESS_INNER,
  checkSecret,
  levelFor,
  loginWith,
  makeSecret,
  packAccess,
  parseAccess,
  passwordTaken,
  readAccess,
  setAccessSession,
  writeAccess,
  type AccessFile,
} from './access';
import { Bundle } from './container';
import type { EditorHost } from './host';
import { addFilesToBundle, flushBundle, itemPath, loadBundle, newBundleFile, withBundles } from './vfs';

// 测试里迭代次数取小，跑得快；算法相同
const ITER = 1000;

async function sample(): Promise<AccessFile> {
  return {
    v: 1,
    admin: await makeSecret('admin123', ITER),
    base: { level: 'visible', items: { i2: 'hidden' } },
    roles: [
      { id: 'r1', name: '采购员', secret: await makeSecret('buy456', ITER), level: 'read', items: { i1: 'write' } },
      { id: 'r2', name: '审核', secret: await makeSecret('chk789', ITER), level: 'write' },
    ],
  };
}

describe('权限：密码与级别', () => {
  it('密码只存加盐哈希；同一个密码两次的哈希不同，都能验证', async () => {
    const a = await makeSecret('同一个密码', ITER);
    const b = await makeSecret('同一个密码', ITER);
    expect(a.hash).not.toBe(b.hash);
    expect(JSON.stringify(a)).not.toContain('同一个密码');
    expect(await checkSecret('同一个密码', a)).toBe(true);
    expect(await checkSecret('同一个密码', b)).toBe(true);
    expect(await checkSecret('别的', a)).toBe(false);
  });

  it('只输密码登录：管理 / 角色 / 错误', async () => {
    const f = await sample();
    expect(await loginWith(f, 'admin123')).toEqual({ kind: 'admin' });
    expect(await loginWith(f, 'buy456')).toEqual({ kind: 'role', id: 'r1' });
    expect(await loginWith(f, 'wrong')).toBeNull();
    expect(await passwordTaken(f, 'chk789')).toBe(true);
    expect(await passwordTaken(f, 'chk789', 'r2')).toBe(false);
  });

  it('级别：默认 + 单个文件另设；管理全部可写；未登录看不到；没设权限谁都可写', async () => {
    const f = await sample();
    expect(levelFor(f, { kind: 'base' }, 'i1')).toBe('visible');
    expect(levelFor(f, { kind: 'base' }, 'i2')).toBe('hidden');
    expect(levelFor(f, { kind: 'role', id: 'r1' }, 'i1')).toBe('write');
    expect(levelFor(f, { kind: 'role', id: 'r1' }, 'i2')).toBe('read');
    expect(levelFor(f, { kind: 'role', id: 'r2' }, 'i2')).toBe('write');
    expect(levelFor(f, { kind: 'role', id: '已删除' }, 'i1')).toBe('hidden');
    expect(levelFor(f, { kind: 'admin' }, 'i2')).toBe('write');
    expect(levelFor(f, null, 'i1')).toBe('hidden');
    expect(levelFor(null, null, 'i1')).toBe('write');
  });

  it('读写设置：往返不变；坏值清洗；没有管理密码 = 不限制；JSON 坏了按最严处理', async () => {
    const f = await sample();
    const b = Bundle.create('t');
    writeAccess(b, f);
    expect(readAccess(Bundle.open(b.serialize()))).toEqual(f);
    expect(parseAccess(JSON.stringify({ v: 1, roles: [] }))).toBeNull();
    const bad = parseAccess(JSON.stringify({ admin: f.admin, base: { level: '随便', items: { i1: 'write', 'a/b': 'read', i2: 'x' } }, roles: [{ id: 'r', secret: f.admin, level: 'zzz' }, { name: '无密码' }] }))!;
    expect(bad.base).toEqual({ level: 'visible', items: { i1: 'write' } });
    expect(bad.roles).toEqual([{ id: 'r', name: '角色', secret: f.admin, level: 'read' }]);
    const broken = parseAccess('{不是 JSON')!;
    expect(broken.base.level).toBe('visible');
    expect(await loginWith(broken, '')).toBeNull();
    // 关闭权限 = 删掉设置
    writeAccess(b, null);
    expect(b.readText(ACCESS_INNER)).toBeUndefined();
  });
});

describe('权限：工作包读写时的限制', () => {
  function fakeHost(files: Map<string, Uint8Array>) {
    const base = {
      readBytes: async (p: string) => files.get(p)!,
      readText: async (p: string) => ({ content: strFromU8(files.get(p)!), encoding: 'UTF-8' }),
      writeBytes: async (p: string, d: Uint8Array) => void files.set(p, d),
      writeText: async (p: string, t: string) => void files.set(p, strToU8(t)),
      canWrite: async () => true,
      workFolders: () => ['D:\\w'],
      openWithSystem: async () => undefined,
    } as unknown as EditorHost;
    return withBundles(base);
  }

  it('按登录身份：不可见 / 可见读不了，可读能读不能写，可写都行；管理才能加文件', async () => {
    const files = new Map<string, Uint8Array>([
      ['D:\\w\\a.md', strToU8('A')],
      ['D:\\w\\b.md', strToU8('B')],
    ]);
    const host = fakeHost(files);
    const pack = 'D:\\w\\权限.wbp';
    await newBundleFile(pack, '权限');
    const [a, b] = await addFilesToBundle(pack, ['D:\\w\\a.md', 'D:\\w\\b.md']);
    const bundle = await loadBundle(pack);
    writeAccess(bundle, {
      v: 1,
      admin: await makeSecret('admin', ITER),
      base: { level: 'hidden', items: { [a!.id]: 'visible' } },
      roles: [{ id: 'r1', name: '读', secret: await makeSecret('r1pw', ITER), level: 'read', items: { [b!.id]: 'write' } }],
    });
    await flushBundle(pack);
    const pa = itemPath(pack, a!);
    const pb = itemPath(pack, b!);

    // 还没登录：什么都读不了，也不能加文件
    await expect(host.readBytes(pa)).rejects.toThrow('没有打开这个文件的权限');
    await expect(addFilesToBundle(pack, ['D:\\w\\a.md'])).rejects.toThrow('只有用管理密码登录');

    // 基础权限：a 只“可见”，仍然读不了
    setAccessSession(pack, { kind: 'base' });
    expect(packAccess(bundle, pack).level(a!.id)).toBe('visible');
    expect(packAccess(bundle, pack).level(b!.id)).toBe('hidden');
    await expect(host.readBytes(pa)).rejects.toThrow('没有打开这个文件的权限');

    // 角色：a 可读（能读、不能写、canWrite 为 false），b 可写
    setAccessSession(pack, { kind: 'role', id: 'r1' });
    expect(strFromU8(await host.readBytes(pa))).toBe('A');
    await expect(host.writeBytes(pa, strToU8('改'))).rejects.toThrow('只读');
    expect(await host.canWrite!(pa)).toBe(false);
    await host.writeBytes(pb, strToU8('B2'));
    expect(strFromU8(Bundle.open(files.get(pack)!).itemBytes(b!.id)!)).toBe('B2');
    expect(host.bundleOf!(pa)!.writable).toBe(false);
    expect(() => host.bundleOf!(pa)!.setColumns('S', null)).toThrow('只读');
    expect(packAccess(bundle, pack).manage).toBe(false);

    // 管理：全部可写，能加文件
    setAccessSession(pack, { kind: 'admin' });
    await host.writeBytes(pa, strToU8('A2'));
    expect(await addFilesToBundle(pack, ['D:\\w\\a.md'])).toHaveLength(1);
    setAccessSession(pack, null);
  });
});

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { packAccess, setAccessSession } from './access';
import { Bundle } from './container';
import type { EditorHost } from './host';
import { KEYS_INNER, PackLockedError, encryptedKeys, keysText, normalizeCode, parseKeys, randomCode, unlockWith, unlockWithRecovery, unwrapKey, wrapKey } from './packKeys';
import { addFilesToBundle, flushBundle, forgetPackKey, itemPath, loadBundle, newBundleFile, unlockPack, withBundles } from './vfs';

const ITER = 1000;

describe('加密包的钥匙', () => {
  it('随机码：6 组 4 位、不含易混字符；恢复码输入不区分大小写和横线', () => {
    const c = randomCode();
    expect(c).toMatch(/^([A-HJ-NP-Z2-9]{4}-){5}[A-HJ-NP-Z2-9]{4}$/);
    expect(randomCode()).not.toBe(c);
    expect(normalizeCode(c.toLowerCase().replace(/-/g, ' '))).toBe(c);
  });

  it('用密码包住解压密钥：对的密码解开，错的不行；按身份登录；恢复码单独解开', async () => {
    const key = randomCode();
    const rec = randomCode();
    const keys = {
      v: 1 as const,
      wraps: [await wrapKey(key, 'admin-pw', 'admin', ITER), await wrapKey(key, 'buyer-pw', 'role:r1', ITER), await wrapKey(key, rec, 'recovery', ITER)],
    };
    expect(JSON.stringify(keys)).not.toContain(key);
    expect(await unwrapKey(keys.wraps[0]!, 'admin-pw')).toBe(key);
    expect(await unwrapKey(keys.wraps[0]!, 'wrong')).toBeNull();
    expect(await unlockWith(keys, 'buyer-pw')).toEqual({ key, who: { kind: 'role', id: 'r1' } });
    expect(await unlockWith(keys, 'admin-pw')).toEqual({ key, who: { kind: 'admin' } });
    // 恢复码不能当普通密码用，只能走“用恢复码进入”
    expect(await unlockWith(keys, rec)).toBeNull();
    expect(await unlockWithRecovery(keys, rec.toLowerCase())).toBe(key);
    expect(await unlockWithRecovery(keys, randomCode())).toBeNull();
    expect(parseKeys(keysText(keys))).toEqual(keys);
  });

  it('识别加密包：只看不加密的 meta/keys.json', () => {
    const plain = Bundle.create('t').serialize();
    expect(encryptedKeys(plain)).toBeNull();
    const enc = zipSync({ mimetype: strToU8('x'), [KEYS_INNER]: strToU8('{"v":1,"wraps":[]}') });
    expect(encryptedKeys(enc)).toEqual({ v: 1, wraps: [] });
    expect(encryptedKeys(strToU8('不是 zip'))).toBeNull();
  });
});

describe('加密包的读写（假宿主模拟加密）', () => {
  /** 假宿主：writePack 把不加密的 zip 存在旁边，“磁盘”上只放 mimetype + keys.json + 一个看不懂的部件 */
  function fakeHost(files: Map<string, Uint8Array>) {
    const sealed = new Map<string, { key: string; plain: Uint8Array }>();
    const base = {
      readBytes: async (p: string) => files.get(p)!,
      readText: async (p: string) => ({ content: strFromU8(files.get(p)!), encoding: 'UTF-8' }),
      writeBytes: async (p: string, d: Uint8Array) => void files.set(p, d),
      writeText: async (p: string, t: string) => void files.set(p, strToU8(t)),
      canWrite: async () => true,
      workFolders: () => ['D:\\w'],
      openWithSystem: async () => undefined,
      readPack: async (p: string, key: string) => {
        const s = sealed.get(p);
        if (!s || s.key !== key) throw new Error('解压密钥不对');
        return s.plain;
      },
      writePack: async (p: string, plain: Uint8Array, key: string) => {
        sealed.set(p, { key, plain });
        const keys = unzipSync(plain)[KEYS_INNER]!;
        files.set(p, zipSync({ mimetype: strToU8('application/vnd.wz.workpackage+zip'), [KEYS_INNER]: keys, 'items/x': strToU8('乱码') }));
      },
    } as unknown as EditorHost;
    return withBundles(base);
  }

  it('加密后：磁盘上没有明文；不给密钥打不开（要求登录）；解开后正常读写；关闭加密写回普通 zip', async () => {
    const files = new Map<string, Uint8Array>([['D:\\w\\合同.md', strToU8('机密合同')]]);
    const host = fakeHost(files);
    const pack = 'D:\\w\\加密.wbp';
    await newBundleFile(pack, '加密');
    const [item] = await addFilesToBundle(pack, ['D:\\w\\合同.md']);
    const b = await loadBundle(pack);

    // 启用加密：写钥匙文件、记住密钥，写盘走 writePack
    const key = randomCode();
    b.writeText(KEYS_INNER, keysText({ v: 1, wraps: [await wrapKey(key, 'pw1234', 'admin', ITER)] }));
    unlockPack(pack, key);
    await flushBundle(pack);
    const onDisk = files.get(pack)!;
    expect(encryptedKeys(onDisk)).not.toBeNull();
    expect(strFromU8(onDisk)).not.toContain('机密合同');

    // 解开后可读写（“没有密钥打不开”见下一个测试）；写回仍是加密的
    setAccessSession(pack, { kind: 'admin' });
    expect(strFromU8(await host.readBytes(itemPath(pack, item!)))).toBe('机密合同');
    await host.writeBytes(itemPath(pack, item!), strToU8('改过的合同'));
    expect(strFromU8(files.get(pack)!)).not.toContain('改过的合同');
    expect(packAccess(b, pack).manage).toBe(true);

    // 关闭加密：删钥匙文件、忘掉密钥 → 普通 zip，明文可见
    b.remove(KEYS_INNER);
    forgetPackKey(pack);
    await flushBundle(pack);
    expect(encryptedKeys(files.get(pack)!)).toBeNull();
    expect(strFromU8(Bundle.open(files.get(pack)!).itemBytes(item!.id)!)).toBe('改过的合同');
    setAccessSession(pack, null);
  });

  it('没有密钥时打开加密包：抛 PackLockedError（带钥匙文件），密钥不对读不出', async () => {
    const files = new Map<string, Uint8Array>();
    fakeHost(files); // 安装宿主（加密包读写走它）
    const pack = 'D:\\w\\锁着.wbp';
    // 直接造一个“磁盘上的加密包”
    const keys = { v: 1 as const, wraps: [await wrapKey('KEY-0000-0000-0000', 'pw', 'admin', ITER)] };
    files.set(pack, zipSync({ mimetype: strToU8('x'), [KEYS_INNER]: strToU8(keysText(keys)) }));
    const err = await loadBundle(pack).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(PackLockedError);
    expect((err as PackLockedError).keys).toEqual(keys);
    unlockPack(pack, 'WRONG-KEY');
    await expect(loadBundle(pack)).rejects.toThrow('解压密钥不对');
    forgetPackKey(pack);
  });
});

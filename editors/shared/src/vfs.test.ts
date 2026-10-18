import { strFromU8, strToU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { Bundle } from './container';
import type { EditorHost } from './host';
import { addFilesToBundle, flushBundle, itemPath, loadBundle, newBundleFile, refreshBundle, withBundles } from './vfs';

/** 内存里的“磁盘”和一个只实现文件读写的假宿主 */
function fakeHost(files: Map<string, Uint8Array>, texts = new Map<string, string>()) {
  const writes: string[] = [];
  const base = {
    readBytes: async (p: string) => {
      const b = files.get(p);
      if (!b) throw new Error(`没有 ${p}`);
      return b;
    },
    readText: async (p: string) => ({ content: texts.get(p) ?? strFromU8(files.get(p)!), encoding: 'UTF-8' }),
    writeBytes: async (p: string, d: Uint8Array) => {
      writes.push(p);
      files.set(p, d);
    },
    writeText: async (p: string, t: string) => void files.set(p, strToU8(t)),
    canWrite: async () => true,
    workFolders: () => ['D:\\w'],
    openWithSystem: async () => undefined,
  } as unknown as EditorHost;
  return { host: withBundles(base), writes };
}

let n = 0;
const packPath = () => `D:\\w\\包${++n}.wbp`;

describe('虚拟路径', () => {
  it('新建包、加入文件、通过虚拟路径读写条目', async () => {
    const files = new Map<string, Uint8Array>([['D:\\w\\订单.xlsx', strToU8('XLSX')]]);
    const { host, writes } = fakeHost(files);
    const pack = packPath();
    await newBundleFile(pack, '测试');
    expect(files.has(pack)).toBe(true);
    const [it1] = await addFilesToBundle(pack, ['D:\\w\\订单.xlsx']);
    expect(it1).toMatchObject({ id: 'i1', kind: 'sheet', name: '订单.xlsx' });

    const vp = itemPath(pack, it1);
    expect(vp).toBe(`${pack}!/items/i1/订单.xlsx`);
    expect(strFromU8(await host.readBytes(vp))).toBe('XLSX');

    const before = writes.length;
    await host.writeBytes(vp, strToU8('XLSX2'));
    expect(writes.length).toBe(before + 1);
    expect(writes.at(-1)).toBe(pack);
    // 重新从“磁盘”读，内容已在包里
    expect(strFromU8(Bundle.open(files.get(pack)!).itemBytes('i1')!)).toBe('XLSX2');
  });

  it('包外路径原样转发', async () => {
    const files = new Map([['D:\\w\\a.txt', strToU8('hi')]]);
    const { host } = fakeHost(files);
    expect(strFromU8(await host.readBytes('D:\\w\\a.txt'))).toBe('hi');
    expect(host.bundleOf?.('D:\\w\\a.xlsx')).toBeNull();
  });

  it('文本条目：加入时转成 UTF-8；读回 UTF-8；写文本不管编码参数', async () => {
    const files = new Map([['D:\\w\\说明.md', strToU8('旧')]]);
    const texts = new Map([['D:\\w\\说明.md', '# 标题（来自 GBK 文件）']]);
    const { host } = fakeHost(files, texts);
    const pack = packPath();
    await newBundleFile(pack);
    const [item] = await addFilesToBundle(pack, ['D:\\w\\说明.md']);
    const vp = itemPath(pack, item);
    expect((await host.readText(vp)).content).toBe('# 标题（来自 GBK 文件）');
    await host.writeText(vp, '# 新', 'GBK');
    expect((await host.readText(vp)).content).toBe('# 新');
  });

  it('bundleOf：给出列属性 / 表单上下文，改动随条目保存写盘', async () => {
    const files = new Map([['D:\\w\\订单.xlsx', strToU8('X')]]);
    const { host, writes } = fakeHost(files);
    const pack = packPath();
    await newBundleFile(pack);
    const [item] = await addFilesToBundle(pack, ['D:\\w\\订单.xlsx']);
    const vp = itemPath(pack, item);
    const ctx = host.bundleOf!(vp)!;
    expect(ctx).toMatchObject({ pack, itemId: 'i1', itemName: '订单.xlsx' });

    let notified = 0;
    ctx.subscribe(() => notified++);
    ctx.setColumns('订单', { header: 1, columns: [{ id: 'c1', title: '单位', col: 0, type: 'select', options: [{ v: '个' }] }] });
    expect(notified).toBe(1);
    // 还没写盘
    const before = writes.length;
    expect(Bundle.open(files.get(pack)!).columns('i1').sheets).toEqual({});
    await host.writeBytes(vp, strToU8('X'));
    expect(writes.length).toBe(before + 1);
    expect(Bundle.open(files.get(pack)!).columns('i1').sheets['订单'].columns[0].id).toBe('c1');
    expect(ctx.columns().sheets['订单'].columns).toHaveLength(1);

    await ctx.saveForm({ id: ctx.newFormId(), title: '新增', target: { item: 'i1', sheet: '订单' }, fields: [{ column: 'c1' }] });
    expect(Bundle.open(files.get(pack)!).forms()).toHaveLength(1);
    expect(ctx.forms()).toHaveLength(1);
    await ctx.deleteForm('f1');
    expect(ctx.forms()).toHaveLength(0);

    ctx.setColumns('订单', null);
    await ctx.flush();
    expect(Bundle.open(files.get(pack)!).has('meta/i1.columns.json')).toBe(false);
  });

  it('改名后，旧路径（按 id 找）仍然能读写', async () => {
    const src = 'D:\\w\\订单.xlsx';
    const files = new Map([[src, strToU8('X')]]);
    const { host } = fakeHost(files);
    const pack = packPath();
    await newBundleFile(pack);
    const [item] = await addFilesToBundle(pack, [src]);
    const oldPath = itemPath(pack, item);
    const b = await loadBundle(pack);
    b.renameItem('i1', '采购订单.xlsx');
    expect(strFromU8(await host.readBytes(oldPath))).toBe('X');
    await host.writeBytes(oldPath, strToU8('Y'));
    expect(strFromU8(Bundle.open(files.get(pack)!).read('items/i1/采购订单.xlsx')!)).toBe('Y');
    expect(host.bundleOf!(oldPath)?.itemName).toBe('采购订单.xlsx');
  });

  it('写不存在的包内文件被拒绝；包内文件不能用系统程序打开', async () => {
    const { host } = fakeHost(new Map());
    const pack = packPath();
    await newBundleFile(pack);
    await expect(host.writeBytes(`${pack}!/items/i9/x.xlsx`, strToU8('x'))).rejects.toThrow(/没有/);
    await expect(host.openWithSystem(`${pack}!/items/i1/a.xlsx`)).rejects.toThrow(/导出/);
  });

  it('外部改了包文件：没有未保存修改时 refreshBundle 重新读入', async () => {
    const files = new Map<string, Uint8Array>();
    const { host } = fakeHost(files);
    const pack = packPath();
    await newBundleFile(pack);
    const b = await loadBundle(pack);
    // 外部程序加了一个条目
    const other = Bundle.open(files.get(pack)!);
    other.addItem('外来.md', strToU8('x'));
    files.set(pack, other.serialize());
    expect(await refreshBundle(pack)).toBe(true);
    expect(b.manifest.items).toHaveLength(1);
    // 有未保存修改时不刷新
    b.addItem('本地.md', strToU8('y'));
    expect(await refreshBundle(pack)).toBe(false);
    await flushBundle(pack);
    expect(host).toBeDefined();
  });
});

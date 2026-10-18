import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { beforeEach, describe, expect, it } from 'vitest';
import { Bundle, MIMETYPE, columnsInner, itemInner, kindOfName, safeInner, safeName, splitVPath } from './container';
import { checkColumnValue, cleanColumnsFile, cleanForm, splitMulti } from './schema';
import { writeXlsx } from './xlsx';

const bytes = (s: string) => strToU8(s);

beforeEach(() => {
  Bundle.now = () => '2026-10-06T00:00:00.000Z';
});

function sample(): Uint8Array {
  const b = Bundle.create('订单包');
  b.addItem('订单.xlsx', bytes('XLSX-BYTES'));
  b.addItem('说明.md', bytes('# 说明'));
  const bytes0 = b.serialize();
  b.markSaved(bytes0);
  return bytes0;
}

describe('路径', () => {
  it('虚拟路径拆分', () => {
    expect(splitVPath('D:\\w\\a.wbp!/items/i1/订单.xlsx')).toEqual({ pack: 'D:\\w\\a.wbp', inner: 'items/i1/订单.xlsx' });
    expect(splitVPath('D:\\w\\a.xlsx')).toBeNull();
    expect(splitVPath('a.wbp!/')).toBeNull();
  });
  it('拒绝不安全的包内路径和文件名', () => {
    for (const p of ['../a', '/a', 'a/../b', 'a\\b', 'C:/x', 'a//b', '']) expect(safeInner(p)).toBe(false);
    expect(safeInner('items/i1/订单.xlsx')).toBe(true);
    expect(safeName('a/b.xlsx')).toBe(false);
    expect(safeName('订单.xlsx')).toBe(true);
  });
  it('条目种类按扩展名', () => {
    expect(kindOfName('a.XLSX')).toBe('sheet');
    expect(kindOfName('a.wbd')).toBe('board');
    expect(kindOfName('a.ps1')).toBe('script');
    expect(kindOfName('a.bin')).toBe('file');
  });
});

describe('容器往返', () => {
  it('新建的包：mimetype 是第一个条目且不压缩，清单字段齐全', () => {
    const data = sample();
    expect(Object.keys(unzipSync(data))[0]).toBe('mimetype');
    expect(strFromU8(unzipSync(data).mimetype)).toBe(MIMETYPE);
    const b = Bundle.open(data);
    expect(b.manifest.format).toBe('wbp');
    expect(b.manifest.id).toMatch(/^[0-9a-f-]{36}$/);
    expect(b.manifest.created).toBe('2026-10-06T00:00:00.000Z');
    expect(b.manifest.items.map((i) => [i.id, i.kind, i.name, i.mediaType])).toEqual([
      ['i1', 'sheet', '订单.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
      ['i2', 'md', '说明.md', 'text/markdown'],
    ]);
  });

  it('打开即保存 = 原文件字节', () => {
    const data = sample();
    const b = Bundle.open(data);
    expect(b.dirty).toBe(false);
    expect(b.serialize()).toBe(data);
  });

  it('只改一个条目：其它部件字节不变', () => {
    const data = sample();
    const b = Bundle.open(data);
    b.write(itemInner(b.item('i2')!), bytes('# 改过'));
    const out = unzipSync(b.serialize());
    const before = unzipSync(data);
    expect(strFromU8(out[itemInner(b.item('i2')!)])).toBe('# 改过');
    expect(out[itemInner(b.item('i1')!)]).toEqual(before[itemInner(b.item('i1')!)]);
    expect(Object.keys(out)).toEqual(Object.keys(before));
  });

  it('写入相同内容不算改动', () => {
    const b = Bundle.open(sample());
    b.write(itemInner(b.item('i1')!), bytes('XLSX-BYTES'));
    expect(b.dirty).toBe(false);
  });

  it('不认识的部件和清单扩展字段原样保留', () => {
    const base = unzipSync(sample());
    const manifest = JSON.parse(strFromU8(base['manifest.json']));
    manifest['x-acme'] = { a: 1 };
    manifest.items[0]['x-acme-tag'] = 'k';
    manifest.bogus = 1;
    const data = zipSync({
      mimetype: base.mimetype,
      'manifest.json': strToU8(JSON.stringify(manifest)),
      ...Object.fromEntries(Object.entries(base).filter(([n]) => n !== 'mimetype' && n !== 'manifest.json')),
      'extra/unknown.bin': bytes('KEEP'),
    });
    const b = Bundle.open(data);
    expect(b.manifest.ext).toEqual({ 'x-acme': { a: 1 } });
    expect(b.item('i1')!.ext).toEqual({ 'x-acme-tag': 'k' });
    b.write(itemInner(b.item('i2')!), bytes('# 改过'));
    const again = Bundle.open(b.serialize());
    expect(strFromU8(again.read('extra/unknown.bin')!)).toBe('KEEP');
    expect(again.manifest.ext).toEqual({ 'x-acme': { a: 1 } });
    expect(again.item('i1')!.ext).toEqual({ 'x-acme-tag': 'k' });
    // 不认识的普通字段（bogus）不保留
    expect(JSON.parse(strFromU8(unzipSync(b.serialize())['manifest.json'])).bogus).toBeUndefined();
  });

  it('清单里写的种类和媒体类型不被信任；缺文件的条目被丢弃', () => {
    const base = unzipSync(sample());
    const manifest = JSON.parse(strFromU8(base['manifest.json']));
    manifest.items[0].kind = 'script';
    manifest.items[0].mediaType = 'text/evil';
    manifest.items.push({ id: 'i9', name: '不存在.docx' });
    const data = zipSync({ ...base, 'manifest.json': strToU8(JSON.stringify(manifest)) });
    const b = Bundle.open(data);
    expect(b.item('i1')).toMatchObject({ kind: 'sheet', mediaType: expect.stringContaining('spreadsheetml') });
    expect(b.item('i9')).toBeUndefined();
  });

  it('拒绝：不是 zip / 缺清单 / 路径穿越 / 错的 mimetype', () => {
    expect(() => Bundle.open(bytes('nope'))).toThrow(/有效/);
    expect(() => Bundle.open(zipSync({ a: bytes('x') }))).toThrow(/manifest/);
    const base = unzipSync(sample());
    expect(() => Bundle.open(zipSync({ ...base, '../evil': bytes('x') }))).toThrow(/不安全/);
    expect(() => Bundle.open(zipSync({ ...base, mimetype: bytes('application/zip') }))).toThrow(/不是工作包/);
  });

  it('更新版本只读', () => {
    const base = unzipSync(sample());
    const manifest = JSON.parse(strFromU8(base['manifest.json']));
    manifest.version = 99;
    const b = Bundle.open(zipSync({ ...base, 'manifest.json': strToU8(JSON.stringify(manifest)) }));
    expect(b.newer).toBe(true);
    expect(() => b.write('items/i1/订单.xlsx', bytes('x'))).toThrow(/只能查看/);
    expect(() => b.serialize()).toThrow();
  });
});

describe('条目操作', () => {
  it('调整条目顺序：保存后重新打开顺序不变，文件内容字节不变', () => {
    const b = Bundle.create('t');
    const a = b.addItem('a.md', bytes('A'));
    const c = b.addItem('b.md', bytes('B'));
    const d = b.addItem('c.md', bytes('C'));
    b.moveItem(d.id, a.id); // c 移到 a 前
    expect(b.manifest.items.map((i) => i.name)).toEqual(['c.md', 'a.md', 'b.md']);
    b.moveItem(d.id, c.id, true); // c 移到 b 后
    expect(b.manifest.items.map((i) => i.name)).toEqual(['a.md', 'b.md', 'c.md']);
    b.moveItem(a.id, d.id, true); // a 移到最后
    const again = Bundle.open(b.serialize());
    expect(again.manifest.items.map((i) => i.name)).toEqual(['b.md', 'c.md', 'a.md']);
    expect(strFromU8(again.itemBytes(a.id)!)).toBe('A');
  });

  it('同名自动改名；删除条目带走列属性和表单', () => {
    const b = Bundle.open(sample());
    const dup = b.addItem('订单.xlsx', bytes('2'));
    expect(dup.name).toBe('订单 (1).xlsx');
    b.setColumns('i1', { v: 1, sheets: { 订单: { header: 1, columns: [{ id: 'c1', title: '状态', col: 0, type: 'select', options: [{ v: '待发' }] }] } } });
    b.saveForm({ id: 'f1', title: '新增', target: { item: 'i1', sheet: '订单' }, fields: [{ column: 'c1' }] });
    expect(b.has(columnsInner('i1'))).toBe(true);
    expect(b.forms()).toHaveLength(1);
    b.removeItem('i1');
    expect(b.has(columnsInner('i1'))).toBe(false);
    expect(b.has(itemInner({ id: 'i1', name: '订单.xlsx' }))).toBe(false);
    expect(b.forms()).toHaveLength(0);
  });

  it('改名后 id 和列属性不变，不能改扩展名', () => {
    const b = Bundle.open(sample());
    b.setColumns('i1', { v: 1, sheets: { S: { header: 1, columns: [] } } });
    b.renameItem('i1', '采购订单.xlsx');
    expect(b.item('i1')!.name).toBe('采购订单.xlsx');
    expect(b.read('items/i1/采购订单.xlsx')).toBeDefined();
    expect(b.read('items/i1/订单.xlsx')).toBeUndefined();
    expect(() => b.renameItem('i1', '采购订单.docx')).toThrow(/扩展名/);
  });

  it('保存期间又有改动：markSaved 后仍是已修改', () => {
    const b = Bundle.open(sample());
    b.write('items/i2/说明.md', bytes('a'));
    const snap = b.snapshot();
    b.write('items/i2/说明.md', bytes('b'));
    b.markSaved(snap.bytes, snap.rev);
    expect(b.dirty).toBe(true);
  });

  it('外部改了文件：reloadFrom 就地更新并通知', () => {
    const b = Bundle.open(sample());
    let n = 0;
    b.subscribe(() => n++);
    const other = Bundle.open(sample());
    other.addItem('新.md', bytes('x'));
    b.reloadFrom(other.serialize());
    expect(b.manifest.items).toHaveLength(3);
    expect(n).toBe(1);
  });
});

describe('跨条目引用', () => {
  it('[条目名]工作表!区域 读出文字：去空去重、按出现顺序；引用不对返回空', () => {
    const b = Bundle.create('t');
    b.addItem('供应商.xlsx', writeXlsx([{ name: '名单', rows: [['名称'], ['供应商甲'], ['供应商乙'], ['供应商甲'], [null]] }, { name: '另一张', rows: [['x']] }]));
    expect(b.readRange('[供应商.xlsx]名单!A2:A5')).toEqual(['供应商甲', '供应商乙']);
    expect(b.readRange('[供应商.xlsx]A2:A3')).toEqual(['供应商甲', '供应商乙']);
    expect(b.readRange('[i1]另一张!A1')).toEqual(['x']);
    expect(b.readRange('[没有.xlsx]名单!A1')).toEqual([]);
    expect(b.readRange('[供应商.xlsx]没有的表!A1')).toEqual([]);
    expect(b.readRange('A2:A4')).toEqual([]);
  });
});

describe('列属性与表单清洗', () => {
  it('列属性：丢弃坏列、重复 id / 列号、坏颜色', () => {
    const f = cleanColumnsFile({
      sheets: {
        订单: {
          header: 2,
          columns: [
            { id: 'c1', title: '状态', col: 0, type: 'select', options: [{ v: '待发', color: '#f59e0b' }, { v: '已发', color: 'red' }, '待发'] },
            { id: 'c1', title: '重复 id', col: 1, type: 'text' },
            { id: 'c2', title: '重复列号', col: 0, type: 'text' },
            { id: 'bad id', title: 'x', col: 2, type: 'text' },
            { id: 'c3', title: '未知类型', col: 3, type: 'weird' },
            { id: 'c4', title: '金额', col: 4, type: 'money', min: 0, required: true, default: '0' },
          ],
        },
      },
    });
    const cols = f.sheets['订单'].columns;
    expect(f.sheets['订单'].header).toBe(2);
    // 未知类型（c3）当文本列保留，坏 id / 重复 id / 重复列号被丢弃
    expect(cols.map((c) => c.id)).toEqual(['c1', 'c3', 'c4']);
    expect(cols[1]).toMatchObject({ type: 'text', unknownType: 'weird' });
    expect(cols[0].options).toEqual([{ v: '待发', color: '#f59e0b' }, { v: '已发' }]);
    expect(cols[2]).toMatchObject({ required: true, default: '0', min: 0 });
  });

  it('未知列类型和 x- 扩展字段：读入后原样写回', () => {
    const b = Bundle.open(sample());
    const raw = {
      v: 1,
      'x-acme': { a: 1 },
      sheets: {
        S: {
          header: 1,
          'x-acme-sheet': true,
          columns: [
            { id: 'c1', title: '评分', col: 0, type: 'rating', max: 5, stars: true },
            { id: 'c2', title: '状态', col: 1, type: 'select', options: [{ v: 'a', 'x-acme-c': 1 }], 'x-acme-col': 'k' },
          ],
        },
      },
    };
    b.writeText(columnsInner('i1'), JSON.stringify(raw));
    const cols = b.columns('i1');
    expect(cols.sheets.S.columns[0]).toMatchObject({ type: 'text', unknownType: 'rating' });
    b.setColumns('i1', cols);
    expect(JSON.parse(b.readText(columnsInner('i1'))!)).toEqual(raw);
  });

  it('表单：x- 扩展字段保留', () => {
    const b = Bundle.open(sample());
    b.saveForm({ id: 'f1', title: 't', target: { item: 'i1', sheet: 'S' }, fields: [{ column: 'c1', ext: { 'x-a': 1 } }], ext: { 'x-b': 2 } });
    expect(b.form('f1')).toMatchObject({ ext: { 'x-b': 2 }, fields: [{ column: 'c1', ext: { 'x-a': 1 } }] });
  });

  it('表单：字段引用去重，坏表单为 null', () => {
    expect(cleanForm({ id: 'f1', title: 't', target: { item: 'i1', sheet: 'S' }, fields: [{ column: 'c1' }, { column: 'c1' }, { column: 'a b' }] })!.fields).toEqual([{ column: 'c1' }]);
    expect(cleanForm({ id: 'f1', target: {}, fields: [] })).toBeNull();
  });

  it('列值校验', () => {
    const num = { id: 'c', title: '数量', col: 0, type: 'number' as const, integer: true, min: 0 };
    expect(checkColumnValue(num, '8.5')).toBe('只允许整数');
    expect(checkColumnValue(num, '-1')).toMatch(/不能小于/);
    expect(checkColumnValue(num, '12')).toBeNull();
    expect(checkColumnValue({ ...num, required: true }, '')).toBe('必填');
    const sel = { id: 's', title: '单位', col: 1, type: 'select' as const, options: [{ v: '个' }, { v: '台' }] };
    expect(checkColumnValue(sel, '米')).toBe('不在选项里');
    expect(checkColumnValue({ ...sel, type: 'multi' }, '个,台')).toBeNull();
    expect(checkColumnValue({ ...sel, type: 'multi' }, '个,米')).toMatch(/米/);
    expect(checkColumnValue({ id: 'd', title: 'd', col: 2, type: 'date' }, '2026-02-30')).toMatch(/日期/);
    expect(checkColumnValue({ id: 'd', title: 'd', col: 2, type: 'date' }, '2026-10-12')).toBeNull();
    expect(splitMulti('a, b，c、d')).toEqual(['a', 'b', 'c', 'd']);
  });
});

import { describe, expect, it } from 'vitest';
import { applyEntry, insertDelete, setInput, Tx, type Entry } from './ops';
import { bindColumns, inferColumns, metaAt, missingRequired, optionValues, shiftColumnMeta } from './colmeta';
import { applyInferences, fillDefaults, makeMeta, rowIsBlank, setColumn, setColumnInput, typeFormat } from './columns';
import { blankWorkbook, type Workbook } from './workbook';

function edit(b: Workbook, fn: (tx: Tx) => void): Entry | null {
  const tx = new Tx(b, 'x');
  fn(tx);
  return tx.commit();
}

/** 采购明细样例：表头 + 数据 */
function sample(): Workbook {
  const b = blankWorkbook();
  const rows = [
    ['物资名称', '单位', '数量', '单价', '到货日期', '已验收', '资料'],
    ['法兰', '个', '24', '86.5', '2026-10-12', '是', '合格证,材质证明'],
    ['截止阀', '台', '6', '1250', '2026-10-15', '否', '合格证'],
    ['钢管', '米', '120', '145.5', '2026-10-18', '否', '材质证明'],
    ['电缆', '米', '800', '12.8', '2026-10-20', '是', '合格证,材质证明'],
    ['垫片', '个', '50', '6.5', '2026-10-12', '是', '合格证'],
    ['灯具', '台', '18', '320', '2026-10-25', '否', ''],
  ];
  edit(b, (tx) => rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, v))));
  return b;
}

describe('列属性：推断', () => {
  it('按内容推断类型、选项和依据', () => {
    const b = sample();
    const inf = inferColumns(b, b.sheets[0], 1);
    const by = Object.fromEntries(inf.map((x) => [x.head, x]));
    expect(by['物资名称'].type).toBe('text');
    expect(by['单位']).toMatchObject({ type: 'select', options: ['个', '台', '米'] });
    expect(by['数量']).toMatchObject({ type: 'number', integer: true });
    expect(by['单价']).toMatchObject({ type: 'money', fmt: '#,##0.00' });
    expect(by['到货日期'].type).toBe('date');
    expect(by['已验收'].type).toBe('check');
    expect(by['资料']).toMatchObject({ type: 'multi', options: ['合格证', '材质证明'] });
  });

  it('日期列里混了一个非日期：给出警告', () => {
    const b = sample();
    edit(b, (tx) => setInput(tx, 0, 6, 4, '待定'));
    const d = inferColumns(b, b.sheets[0], 1).find((x) => x.head === '到货日期')!;
    expect(d.type).toBe('date');
    expect(d.warn).toMatch(/E7.*待定/);
  });
});

describe('列属性：设置、撤销、位移', () => {
  it('设置一列：数据区套上格式；一步撤销；不会让工作表被标为已修改', () => {
    const b = sample();
    const wasDirty = b.sheets[0].dirty;
    b.sheets[0].dirty = false;
    const e = edit(b, (tx) => setColumn(tx, 0, 3, makeMeta(b.sheets[0], 3, 'money')))!;
    expect(e.cols).toHaveLength(1);
    expect(b.sheets[0].columns?.columns[0]).toMatchObject({ col: 3, type: 'money', title: '单价', id: 'c1' });
    // 单价列的已有数据显示成金额格式
    expect(b.sheets[0].cells[1]?.[3]?.t).toBe('86.50');
    // 只改列属性（没有设置格式的类型）不重写工作表
    const b2 = sample();
    b2.sheets[0].dirty = false;
    const e2 = edit(b2, (tx) => setColumn(tx, 0, 1, makeMeta(b2.sheets[0], 1, 'select', { options: [{ v: '个' }] })))!;
    expect(e2.cols).toHaveLength(1);
    expect(b2.sheets[0].dirty).toBeFalsy();
    // 撤销 / 重做
    applyEntry(b, e, -1);
    expect(b.sheets[0].columns).toBeUndefined();
    expect(b.sheets[0].cells[1]?.[3]?.t).toBe('86.5');
    applyEntry(b, e, 1);
    expect(b.sheets[0].columns?.columns).toHaveLength(1);
    void wasDirty;
  });

  it('清除一列的列属性', () => {
    const b = sample();
    edit(b, (tx) => setColumn(tx, 0, 1, makeMeta(b.sheets[0], 1, 'select')));
    edit(b, (tx) => setColumn(tx, 0, 1, null));
    expect(b.sheets[0].columns).toBeUndefined();
  });

  it('在前面插入 / 删除列：列属性跟着列走，被删的列去掉；撤销恢复', () => {
    const b = sample();
    edit(b, (tx) => {
      setColumn(tx, 0, 1, makeMeta(b.sheets[0], 1, 'select'));
      setColumn(tx, 0, 3, makeMeta(b.sheets[0], 3, 'money'));
    });
    const ins = edit(b, (tx) => insertDelete(tx, 0, 'col', 0, 2))!;
    expect(b.sheets[0].columns!.columns.map((c) => [c.title, c.col])).toEqual([['单位', 3], ['单价', 5]]);
    applyEntry(b, ins, -1);
    expect(b.sheets[0].columns!.columns.map((c) => c.col)).toEqual([1, 3]);
    const del = edit(b, (tx) => insertDelete(tx, 0, 'col', 1, -1))!;
    expect(b.sheets[0].columns!.columns.map((c) => [c.title, c.col])).toEqual([['单价', 2]]);
    applyEntry(b, del, -1);
    expect(b.sheets[0].columns!.columns).toHaveLength(2);
  });

  it('表头行上方插入 / 删除行：表头行号跟着变', () => {
    const sc = { header: 2, columns: [] };
    expect(shiftColumnMeta(sc, 'row', 0, 1).header).toBe(3);
    expect(shiftColumnMeta(sc, 'row', 5, 1).header).toBe(2);
    expect(shiftColumnMeta(sc, 'row', 0, -1).header).toBe(1);
  });
});

describe('列属性：绑定', () => {
  it('表头对得上原地不动；被移动的列按表头文字找回；找不到的标未绑定', () => {
    const b = sample();
    edit(b, (tx) => {
      setColumn(tx, 0, 1, makeMeta(b.sheets[0], 1, 'select'));
      setColumn(tx, 0, 3, makeMeta(b.sheets[0], 3, 'money'));
      setColumn(tx, 0, 5, makeMeta(b.sheets[0], 5, 'check'));
    });
    const saved = structuredClone(b.sheets[0].columns!);
    // 外部程序在最前面加了两列，并把“已验收”改名
    edit(b, (tx) => insertDelete(tx, 0, 'col', 0, 2));
    edit(b, (tx) => setInput(tx, 0, 0, 7, '验收'));
    b.sheets[0].columns = undefined;
    const r = bindColumns(b.sheets[0], saved);
    expect(r.moved).toEqual(['单位', '单价']);
    expect(r.unbound).toEqual(['已验收']);
    expect(r.sc.columns.map((c) => [c.title, c.col, !!c.unbound])).toEqual([['单位', 3, false], ['单价', 5, false], ['已验收', 5, true]]);
    b.sheets[0].columns = r.sc;
    expect(metaAt(b.sheets[0], 1, 3)?.type).toBe('select');
    // 未绑定的列不套用到任何列
    expect(metaAt(b.sheets[0], 1, 7)).toBeUndefined();
  });
});

describe('列属性：录入', () => {
  it('金额列输入数字套金额格式；多选存逗号分隔文字；日期按列格式', () => {
    const b = sample();
    edit(b, (tx) => {
      setColumn(tx, 0, 3, makeMeta(b.sheets[0], 3, 'money'));
      setColumn(tx, 0, 6, makeMeta(b.sheets[0], 6, 'multi', { options: [{ v: '合格证' }, { v: '材质证明' }] }));
      setColumn(tx, 0, 4, makeMeta(b.sheets[0], 4, 'date'));
    });
    const sc = b.sheets[0].columns!;
    edit(b, (tx) => {
      setColumnInput(tx, 0, 7, 3, sc.columns.find((c) => c.col === 3), '1200');
      setColumnInput(tx, 0, 7, 6, sc.columns.find((c) => c.col === 6), '合格证，材质证明');
      setColumnInput(tx, 0, 7, 4, sc.columns.find((c) => c.col === 4), '2026/10/30');
    });
    expect(b.sheets[0].cells[7]?.[3]?.t).toBe('1,200.00');
    expect(b.sheets[0].cells[7]?.[6]?.t).toBe('合格证,材质证明');
    expect(b.sheets[0].cells[7]?.[4]?.t).toBe('2026-10-30');
    expect(typeFormat({ type: 'date' })).toBe('yyyy-mm-dd');
  });

  it('默认值：新行开始录入时填其它空列，已有内容不覆盖；{{today+7}} 展开', () => {
    const b = sample();
    edit(b, (tx) => {
      setColumn(tx, 0, 1, makeMeta(b.sheets[0], 1, 'select', { options: [{ v: '个' }], default: '个' }));
      setColumn(tx, 0, 4, makeMeta(b.sheets[0], 4, 'date', { default: '{{today+7}}' }));
    });
    expect(rowIsBlank(b.sheets[0], 7, 0)).toBe(true);
    edit(b, (tx) => {
      setInput(tx, 0, 7, 0, '管卡');
      expect(fillDefaults(tx, 0, 7, 0, 'xiong', new Date(2026, 9, 6))).toBe(2);
    });
    expect(b.sheets[0].cells[7]?.[1]?.t).toBe('个');
    expect(b.sheets[0].cells[7]?.[4]?.t).toBe('2026-10-13');
    // 已有内容的格子不覆盖
    edit(b, (tx) => {
      setInput(tx, 0, 8, 0, '管夹');
      setInput(tx, 0, 8, 1, '台');
      fillDefaults(tx, 0, 8, 0, null, new Date(2026, 9, 6));
    });
    expect(b.sheets[0].cells[8]?.[1]?.t).toBe('台');
  });
});

describe('列属性：必填与选项', () => {
  it('必填列空着的格子：只看已有内容的行', () => {
    const b = sample();
    edit(b, (tx) => setColumn(tx, 0, 6, makeMeta(b.sheets[0], 6, 'multi', { required: true })));
    // 灯具那行资料为空；整行空的行不算
    expect(missingRequired(b.sheets[0])).toEqual([{ r: 6, c: 6 }]);
  });

  it('选项：手填、引用区域（同表 / 别的表）、复选框', () => {
    const b = sample();
    edit(b, (tx) => setInput(tx, 0, 0, 9, '甲'));
    edit(b, (tx) => setInput(tx, 0, 1, 9, '乙'));
    edit(b, (tx) => setInput(tx, 0, 2, 9, '甲'));
    const s = b.sheets[0];
    expect(optionValues(b, s, makeMeta(s, 1, 'select', { options: [{ v: 'a' }, { v: 'b' }] }))).toEqual(['a', 'b']);
    expect(optionValues(b, s, makeMeta(s, 1, 'select', { optionsRef: 'J1:J3' }))).toEqual(['甲', '乙']);
    expect(optionValues(b, s, makeMeta(s, 1, 'select', { optionsRef: 'Sheet1!J2:J3' }))).toEqual(['乙', '甲']);
    expect(optionValues(b, s, makeMeta(s, 1, 'check'))).toEqual(['是', '否']);
  });

  it('选项引用包里别的表格：由宿主提供读取，没有宿主时为空', () => {
    const b = sample();
    const s = b.sheets[0];
    const meta = makeMeta(s, 1, 'select', { optionsRef: '[供应商.xlsx]名单!A2:A9' });
    expect(optionValues(b, s, meta)).toEqual([]);
    b.externalValues = (ref) => (ref.startsWith('[供应商.xlsx]') ? ['供应商甲', '供应商乙'] : []);
    expect(optionValues(b, s, meta)).toEqual(['供应商甲', '供应商乙']);
  });

  it('应用推断结果：文本列不设置，一步撤销', () => {
    const b = sample();
    const inf = inferColumns(b, b.sheets[0], 1);
    let n = 0;
    const e = edit(b, (tx) => (n = applyInferences(tx, 0, 1, inf)))!;
    expect(n).toBe(6);
    expect(b.sheets[0].columns!.columns.map((c) => c.type)).toEqual(['select', 'number', 'money', 'date', 'check', 'multi']);
    applyEntry(b, e, -1);
    expect(b.sheets[0].columns).toBeUndefined();
  });
});

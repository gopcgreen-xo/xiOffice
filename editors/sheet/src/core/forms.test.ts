import { describe, expect, it } from 'vitest';
import type { FormDef } from '@wz/editor-shared';
import { applyInferences, setColumn, makeMeta } from './columns';
import { inferColumns } from './colmeta';
import { defaultFormFields, initialValues, nextAutoNumber, nextRow, resolveForm, rowValues, submitForm, validateValues } from './forms';
import { applyEntry, setInput, Tx, type Entry } from './ops';
import { setAutoFilter } from './sortFilter';
import { blankWorkbook, type Workbook } from './workbook';

function edit(b: Workbook, fn: (tx: Tx) => void): Entry | null {
  const tx = new Tx(b, 'x');
  fn(tx);
  return tx.commit();
}

function sample(): Workbook {
  const b = blankWorkbook();
  const rows = [
    ['编号', '物资名称', '单位', '数量', '到货日期', '已验收', '资料'],
    ['1', '法兰', '个', '24', '2026-10-12', '是', '合格证,材质证明'],
    ['2', '截止阀', '台', '6', '2026-10-15', '否', '合格证'],
    ['3', '钢管', '米', '120', '2026-10-18', '否', '材质证明'],
    ['4', '电缆', '个', '800', '2026-10-20', '是', '合格证,材质证明'],
    ['5', '垫片', '台', '50', '2026-10-12', '是', '合格证'],
    ['6', '灯具', '个', '18', '2026-10-25', '否', '材质证明'],
  ];
  edit(b, (tx) => rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, v))));
  edit(b, (tx) => {
    applyInferences(tx, 0, 1, inferColumns(b, b.sheets[0], 1));
  });
  // 编号列设成自动编号，物资名称必填
  edit(b, (tx) => {
    setColumn(tx, 0, 0, makeMeta(b.sheets[0], 0, 'autonum'));
    setColumn(tx, 0, 1, makeMeta(b.sheets[0], 1, 'text', { required: true }));
    setColumn(tx, 0, 2, { ...b.sheets[0].columns!.columns.find((c) => c.col === 2)!, options: [{ v: '个' }, { v: '台' }, { v: '米' }], default: '个' });
    setColumn(tx, 0, 4, { ...b.sheets[0].columns!.columns.find((c) => c.col === 4)!, default: '{{today+7}}' });
  });
  return b;
}

const NOW = new Date(2026, 9, 6, 9, 30, 0);

describe('表单：字段与默认值', () => {
  it('默认表单 = 已定义列属性的列，按列顺序；自动编号字段带 auto', () => {
    const b = sample();
    const f = defaultFormFields(b.sheets[0].columns);
    expect(f.map((x) => x.meta.title)).toEqual(['编号', '物资名称', '单位', '数量', '到货日期', '已验收', '资料']);
    expect(f[0].auto).toBe('autonum');
  });

  it('表单定义覆盖标签 / 顺序 / 隐藏；引用不到的列跳过', () => {
    const b = sample();
    const cols = b.sheets[0].columns!;
    const form: FormDef = {
      id: 'f1',
      title: '收货',
      target: { item: 'i1', sheet: 'Sheet1' },
      fields: [{ column: cols.columns[2].id, label: '计量单位', keep: true }, { column: 'nope' }, { column: cols.columns[1].id, hidden: true }],
    };
    const f = resolveForm(cols, form);
    expect(f.map((x) => [x.label, x.hidden, x.keep])).toEqual([['计量单位', false, true], ['物资名称', true, false]]);
  });

  it('初始值：默认值展开、沿用上一条的值', () => {
    const b = sample();
    const fields = defaultFormFields(b.sheets[0].columns).map((f) => (f.meta.col === 2 ? { ...f, keep: true } : f));
    const v1 = initialValues(fields, 'xiong', NOW);
    const unit = fields.find((f) => f.meta.col === 2)!.meta.id;
    const date = fields.find((f) => f.meta.col === 4)!.meta.id;
    expect(v1[unit]).toBe('个');
    expect(v1[date]).toBe('2026-10-13');
    expect(initialValues(fields, null, NOW, { [unit]: '台' })[unit]).toBe('台');
  });
});

describe('表单：校验', () => {
  it('必填、选项、整数；自动字段不检查', () => {
    const b = sample();
    const fields = defaultFormFields(b.sheets[0].columns);
    const id = (col: number) => fields.find((f) => f.meta.col === col)!.meta.id;
    const errs = validateValues(fields, { [id(0)]: '', [id(1)]: '', [id(2)]: '米桶', [id(3)]: '8.5', [id(4)]: '2026-02-30', [id(5)]: '也许', [id(6)]: '' });
    expect(errs[id(0)]).toBeUndefined();
    expect(errs[id(1)]).toBe('必填');
    expect(errs[id(2)]).toBe('不在选项里');
    expect(errs[id(3)]).toBe('只允许整数');
    expect(errs[id(4)]).toMatch(/日期/);
    expect(errs[id(5)]).toMatch(/是.*否/);
    expect(errs[id(6)]).toBeUndefined();
  });
});

describe('表单：提交', () => {
  it('新建：写到下一个空行，自动编号 + 1，一步撤销', () => {
    const b = sample();
    const fields = defaultFormFields(b.sheets[0].columns);
    const id = (col: number) => fields.find((f) => f.meta.col === col)!.meta.id;
    expect(nextRow(b.sheets[0])).toBe(7);
    expect(nextAutoNumber(b.sheets[0], 0)).toBe(7);
    let row = -1;
    const e = edit(b, (tx) => {
      row = submitForm(tx, 0, fields, { [id(1)]: '手动蝶阀', [id(2)]: '台', [id(3)]: '8', [id(4)]: '2026-10-28', [id(5)]: '否', [id(6)]: '合格证' }, { kind: 'new' }, 'xiong', NOW).row;
    })!;
    expect(row).toBe(7);
    expect(b.sheets[0].cells[7].map((c) => c?.t)).toEqual(['7', '手动蝶阀', '台', '8', '2026-10-28', '否', '合格证']);
    expect(nextRow(b.sheets[0])).toBe(8);
    applyEntry(b, e, -1);
    expect(nextRow(b.sheets[0])).toBe(7);
    expect(b.sheets[0].cells[7]?.some((c) => c)).toBeFalsy();
  });

  it('连续提交：编号依次增加；创建时间 / 当前用户自动填', () => {
    const b = sample();
    edit(b, (tx) => setColumn(tx, 0, 7, makeMeta(b.sheets[0], 7, 'created')));
    edit(b, (tx) => setInput(tx, 0, 0, 7, '登记时间'));
    const fields = defaultFormFields(b.sheets[0].columns);
    const name = fields.find((f) => f.meta.col === 1)!.meta.id;
    for (const n of ['A', 'B']) edit(b, (tx) => void submitForm(tx, 0, fields, { [name]: n }, { kind: 'new' }, 'xiong', NOW));
    expect(b.sheets[0].cells[7]?.[0]?.t).toBe('7');
    expect(b.sheets[0].cells[8]?.[0]?.t).toBe('8');
    expect(b.sheets[0].cells[8]?.[7]?.t).toBe('2026-10-06');
  });

  it('编辑当前行：只改非自动字段，编号不变', () => {
    const b = sample();
    const fields = defaultFormFields(b.sheets[0].columns);
    const id = (col: number) => fields.find((f) => f.meta.col === col)!.meta.id;
    const vals = rowValues(b.sheets[0], fields, 2);
    expect(vals[id(1)]).toBe('截止阀');
    edit(b, (tx) => void submitForm(tx, 0, fields, { ...vals, [id(3)]: '9' }, { kind: 'edit', row: 2 }, null, NOW));
    expect(b.sheets[0].cells[2].map((c) => c?.t)).toEqual(['2', '截止阀', '台', '9', '2026-10-15', '否', '合格证']);
  });

  it('提交在自动筛选区域的下一行：筛选范围跟着扩一行', () => {
    const b = sample();
    edit(b, (tx) => setAutoFilter(tx, 0, { r0: 0, c0: 0, r1: 6, c1: 6 }));
    const fields = defaultFormFields(b.sheets[0].columns);
    const name = fields.find((f) => f.meta.col === 1)!.meta.id;
    edit(b, (tx) => void submitForm(tx, 0, fields, { [name]: 'X' }, { kind: 'new' }, null, NOW));
    expect(b.sheets[0].autoFilter?.ref.r1).toBe(7);
  });
});

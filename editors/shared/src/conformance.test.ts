/**
 * 格式一致性：程序写出的清单 / 列属性 / 表单必须符合 docs/wbp/schema 里的 JSON Schema（标准的机器可读形式），
 * 反过来 Schema 也要能挡住明显不合规的内容。这里只带一个够用的 Schema 子集校验器（不引入依赖）。
 */
import { strFromU8, strToU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { Bundle } from './container';
import { load, validate } from '../tests/jsonSchema';

describe('JSON Schema 校验器自检', () => {
  it('能挡住不合规的内容', () => {
    const m = load('manifest.schema.json');
    expect(validate(m, { format: 'zip', version: 1, id: 'x', items: [] }).length).toBeGreaterThan(0);
    expect(validate(m, { format: 'wbp', version: 1, id: '123e4567-e89b-42d3-a456-426614174000', items: [{ id: 'i1', name: 'a/b.xlsx' }] }).length).toBeGreaterThan(0);
    expect(validate(m, { format: 'wbp', version: 1, id: '123e4567-e89b-42d3-a456-426614174000', items: [], bogus: 1 })).toEqual(['$: 不允许的字段 bogus']);
    expect(validate(m, { format: 'wbp', version: 1, id: '123e4567-e89b-42d3-a456-426614174000', items: [], 'x-ok': 1 })).toEqual([]);
  });
});

describe('程序写出的内容符合标准', () => {
  Bundle.now = () => '2026-10-06T08:30:00.000Z';
  const b = Bundle.create('订单包');
  b.addItem('订单.xlsx', strToU8('x'));
  b.addItem('脚本.js', strToU8('1'));
  b.setColumns('i1', {
    v: 1,
    sheets: {
      订单: {
        header: 1,
        columns: [
          { id: 'c1', title: '状态', col: 0, type: 'select', options: [{ v: '待发', color: '#f59e0b' }, { v: '已发' }], required: true, default: '待发' },
          { id: 'c2', title: '金额', col: 1, type: 'money', min: 0, fmt: '¥#,##0.00', note: '含税' },
          { id: 'c3', title: '标签', col: 2, type: 'multi', options: [{ v: 'a' }] },
        ],
      },
    },
  });
  b.saveForm({ id: 'f1', title: '新增订单', target: { item: 'i1', sheet: '订单' }, fields: [{ column: 'c1', keep: true }, { column: 'c2', label: '金额（元）', hint: '含税' }], after: 'continue', onValidate: 'i3' });
  b.saveFlow({ id: 'w1', title: '流程', description: '说明', steps: [{ kind: 'form', ref: 'f1', label: '录入' }, { kind: 'script', ref: 'i3', args: { 公司1: 'x' } }] });
  const files = unzipSync(b.serialize());
  const json = (n: string) => JSON.parse(strFromU8(files[n])) as unknown;

  it('manifest.json', () => expect(validate(load('manifest.schema.json'), json('manifest.json'))).toEqual([]));
  it('列属性', () => expect(validate(load('columns.schema.json'), json('meta/i1.columns.json'))).toEqual([]));
  it('表单', () => expect(validate(load('form.schema.json'), json('forms/f1.json'))).toEqual([]));
  it('联动流程', () => {
    expect(validate(load('flow.schema.json'), json('flows/w1.json'))).toEqual([]);
    expect(validate(load('flow.schema.json'), { id: 'w1', title: 'x', steps: [{ kind: 'bogus', ref: 'a' }] }).length).toBeGreaterThan(0);
  });
  it('提交记录的一行', () => {
    const line = { t: '2026-10-06T08:31:00Z', user: 'xiong', row: 14, mode: 'new', values: { c1: '待发', c2: '620.00' } };
    expect(validate(load('submission.schema.json'), line)).toEqual([]);
    expect(validate(load('submission.schema.json'), { ...line, row: 0 }).length).toBeGreaterThan(0);
  });
  it('离线填写结果', () => {
    const ok = { v: 1, form: 'f1', pack: '11111111-2222-4333-8444-555555555555', records: [{ c1: '法兰', c2: '24' }] };
    expect(validate(load('formresult.schema.json'), ok)).toEqual([]);
    expect(validate(load('formresult.schema.json'), { ...ok, records: [{ c1: 1 }] }).length).toBeGreaterThan(0);
    expect(validate(load('formresult.schema.json'), { v: 1, form: 'f1' }).length).toBeGreaterThan(0);
  });
  it('mimetype 是第一个条目', () => expect(Object.keys(files)[0]).toBe('mimetype'));
});

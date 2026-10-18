// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { applyInferences, makeMeta, setColumn } from './columns';
import { inferColumns } from './colmeta';
import { buildFormHtml, fillableFields } from './formHtml';
import { defaultFormFields, importFormResult, parseFormResult } from './forms';
import { setInput, Tx, type Entry } from './ops';
import { blankWorkbook, type Workbook } from './workbook';

function edit(b: Workbook, fn: (tx: Tx) => void): Entry | null {
  const tx = new Tx(b, 'x');
  fn(tx);
  return tx.commit();
}

function sample(): Workbook {
  const b = blankWorkbook();
  const rows = [
    ['物资名称', '单位', '数量', '已验收', '备注'],
    ['法兰', '个', '24', '是', ''],
    ['截止阀', '台', '6', '否', ''],
    ['钢管', '米', '120', '否', ''],
    ['电缆', '个', '800', '是', ''],
    ['垫片', '台', '50', '是', ''],
    ['灯具', '个', '18', '否', ''],
  ];
  edit(b, (tx) => rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, v))));
  edit(b, (tx) => void applyInferences(tx, 0, 1, inferColumns(b, b.sheets[0], 1)));
  edit(b, (tx) => {
    setColumn(tx, 0, 0, makeMeta(b.sheets[0], 0, 'text', { required: true }));
    setColumn(tx, 0, 4, makeMeta(b.sheets[0], 4, 'text'));
  });
  return b;
}

const html = (b: Workbook, title = '到货登记') => {
  const fields = defaultFormFields(b.sheets[0].columns);
  return buildFormHtml({ packTitle: '采购', packId: '11111111-2222-4333-8444-555555555555', formId: 'f1', title, fields, optionsOf: (f) => f.meta.options ?? [] });
};

/** 在 jsdom 里跑生成页的脚本；返回 document 和导出时下载的内容 */
function load(text: string) {
  const doc = document.implementation.createHTMLDocument('x');
  doc.documentElement.innerHTML = text.replace(/^<!doctype html>/i, '');
  const exported: string[] = [];
  const script = [...doc.querySelectorAll('script')].find((s) => !s.type)!.textContent!;
  class FakeBlob {
    constructor(public parts: string[]) {}
  }
  const fakeURL = { createObjectURL: (b: FakeBlob) => (exported.push(b.parts.join('')), 'blob:x') };
  const realCreate = doc.createElement.bind(doc);
  const fakeDoc = new Proxy(doc, {
    get: (t, k) => {
      if (k === 'createElement') return (tag: string) => (tag === 'a' ? { click: () => undefined, set href(_: string) {}, set download(_: string) {} } : realCreate(tag));
      const v = Reflect.get(t, k) as unknown;
      return typeof v === 'function' ? (v as (...a: unknown[]) => unknown).bind(t) : v;
    },
  });
  new Function('document', 'URL', 'Blob', script)(fakeDoc, fakeURL, FakeBlob);
  return { doc, exported };
}

describe('离线填写页', () => {
  it('只放能填的字段（自动编号 / 隐藏 / 自动字段不放）；标题等文字都被转义', () => {
    const b = sample();
    edit(b, (tx) => setColumn(tx, 0, 5, makeMeta(b.sheets[0], 5, 'autonum')));
    const fields = defaultFormFields(b.sheets[0].columns);
    expect(fillableFields(fields).map((f) => f.meta.title)).toEqual(['物资名称', '单位', '数量', '已验收', '备注']);
    const evil = html(b, '<script>alert(1)</script>"&');
    expect(evil).toContain('&lt;script&gt;alert(1)&lt;/script&gt;&quot;&amp;');
    expect(evil.match(/<script>alert/g)).toBeNull();
    expect(evil).toContain("default-src 'none'");
  });

  it('页面里能填：校验、连续添加、导出的 JSON 能被本程序读回', () => {
    const b = sample();
    const { doc, exported } = load(html(b));
    const q = <T extends Element>(s: string) => doc.querySelector(s) as T;
    const ids = Object.fromEntries(defaultFormFields(b.sheets[0].columns).map((f) => [f.meta.title, f.meta.id]));
    const set = (title: string, v: string) => {
      const el = q<HTMLInputElement>(`[data-id="${ids[title]}"]`);
      el.value = v;
    };
    // 必填没填 / 数量不是整数
    q<HTMLButtonElement>('#add').click();
    expect(q('#msg').textContent).toMatch(/没填对/);
    expect(doc.querySelector(`[data-field="${ids['物资名称']}"] .err`)!.textContent).toBe('必填');
    set('物资名称', '蝶阀');
    set('单位', '台');
    set('数量', '8.5');
    q<HTMLButtonElement>('#add').click();
    expect(doc.querySelector(`[data-field="${ids['数量']}"] .err`)!.textContent).toBe('只允许整数');
    set('数量', '8');
    q<HTMLButtonElement>('#add').click();
    expect(q('#n').textContent).toBe('1');
    // 第二条
    set('物资名称', '压力表');
    set('单位', '个');
    set('数量', '10');
    q<HTMLInputElement>(`[data-id="${ids['已验收']}"]`).checked = true;
    q<HTMLButtonElement>('#add').click();
    expect(q('#n').textContent).toBe('2');
    q<HTMLButtonElement>('#export').click();
    const result = parseFormResult(exported[0]!);
    expect(result).toMatchObject({ form: 'f1', pack: '11111111-2222-4333-8444-555555555555' });
    expect(result.records).toEqual([
      { [ids['物资名称']!]: '蝶阀', [ids['单位']!]: '台', [ids['数量']!]: '8', [ids['已验收']!]: '否', [ids['备注']!]: '' },
      { [ids['物资名称']!]: '压力表', [ids['单位']!]: '个', [ids['数量']!]: '10', [ids['已验收']!]: '是', [ids['备注']!]: '' },
    ]);
  });
});

describe('导入填写结果', () => {
  it('逐条按列属性校验写入；不合规的整条跳过并说明；= 开头的不当公式', () => {
    const b = sample();
    const fields = defaultFormFields(b.sheets[0].columns);
    const id = (t: string) => fields.find((f) => f.meta.title === t)!.meta.id;
    const result = parseFormResult(
      JSON.stringify({
        v: 1,
        form: 'f1',
        records: [
          { [id('物资名称')]: '蝶阀', [id('单位')]: '台', [id('数量')]: '8' },
          { [id('物资名称')]: '', [id('数量')]: '1' },
          { [id('物资名称')]: '管卡', [id('单位')]: '桶', [id('数量')]: '2' },
          { [id('物资名称')]: '=1+1', [id('单位')]: '个', [id('数量')]: '3' },
          { [id('物资名称')]: '压力表', [id('数量')]: 5 },
        ],
      }),
    );
    let out = { imported: 0, skipped: [] as { index: number; reason: string }[] };
    const e = edit(b, (tx) => void (out = importFormResult(tx, 0, fields, result, null)));
    expect(out.imported).toBe(3);
    expect(out.skipped).toEqual([
      { index: 2, reason: '物资名称：必填' },
      { index: 3, reason: '单位：不在选项里' },
    ]);
    expect(e).not.toBeNull();
    const rows = b.sheets[0].cells.slice(7).map((r) => r?.map((c) => c?.t));
    expect(rows[0]?.slice(0, 3)).toEqual(['蝶阀', '台', '8']);
    expect(rows[1]?.[0]).toBe('=1+1');
    expect(b.sheets[0].cells[8]?.[0]?.f).toBeUndefined();
    expect(rows[2]?.slice(0, 3)).toEqual(['压力表', undefined, '5']);
  });

  it('parseFormResult：不是填写结果 / 太大都拒绝', () => {
    expect(() => parseFormResult('不是 json')).toThrow(/无法解析/);
    expect(() => parseFormResult('{"a":1}')).toThrow(/不是离线填写页/);
    expect(() => parseFormResult(JSON.stringify({ form: 'f1', records: Array.from({ length: 5001 }, () => ({})) }))).toThrow(/最多/);
  });
});

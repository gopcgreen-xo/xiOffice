/**
 * 数据验证（同 Excel）：序列（下拉列表，写死的 "a,b,c" 或引用一片单元格 / 名称）、整数、小数、日期、文字长度
 * 及其比较方式（介于、大于等于……）。“出错警告”为停止时不让输入，警告 / 信息时提示后可继续。
 */
import { listRefs } from './formula/refs';
import type { Engine } from './formula/eval';
import { contains } from './grid';
import type { SheetData, Validation, Workbook } from './workbook';

export function validationAt(sheet: SheetData, r: number, c: number): Validation | undefined {
  return sheet.validations?.find((v) => v.ranges.some((g) => contains(g, r, c)));
}

/** 序列验证的选项 */
export function listItems(v: Validation, book: Workbook, si: number): string[] {
  const f = (v.formula1 ?? '').trim();
  if (!f) return [];
  // 写死的列表："是,否"
  if (f.startsWith('"')) {
    return f
      .replace(/^"|"$/g, '')
      .split(',')
      .map((x) => x.trim())
      .filter(Boolean);
  }
  // 名称：换成它的公式
  const name = book.names?.find((n) => n.name.toLowerCase() === f.toLowerCase());
  const refs = listRefs(name ? name.f : f);
  const out: string[] = [];
  for (const ref of refs) {
    const s = ref.sheet === undefined ? si : book.sheets.findIndex((x) => x.name.toLowerCase() === ref.sheet!.toLowerCase());
    const sheet = book.sheets[s];
    if (!sheet) continue;
    const a = ref.a;
    const b = ref.b ?? a;
    const r0 = Math.min(a.r ?? 0, b.r ?? 0);
    const r1 = Math.min(Math.max(a.r ?? sheet.rows - 1, b.r ?? sheet.rows - 1), sheet.rows - 1);
    const c0 = Math.min(a.c ?? 0, b.c ?? 0);
    const c1 = Math.min(Math.max(a.c ?? sheet.cols - 1, b.c ?? sheet.cols - 1), sheet.cols - 1);
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
      const t = sheet.cells[r]?.[c]?.t;
      if (t) out.push(t);
    }
  }
  return [...new Set(out)];
}

const OPS: Record<string, (x: number, a: number, b: number) => boolean> = {
  between: (x, a, b) => x >= Math.min(a, b) && x <= Math.max(a, b),
  notBetween: (x, a, b) => x < Math.min(a, b) || x > Math.max(a, b),
  equal: (x, a) => x === a,
  notEqual: (x, a) => x !== a,
  greaterThan: (x, a) => x > a,
  lessThan: (x, a) => x < a,
  greaterThanOrEqual: (x, a) => x >= a,
  lessThanOrEqual: (x, a) => x <= a,
};

const OP_TEXT: Record<string, string> = {
  between: '介于 {a} 和 {b} 之间',
  notBetween: '不在 {a} 和 {b} 之间',
  equal: '等于 {a}',
  notEqual: '不等于 {a}',
  greaterThan: '大于 {a}',
  lessThan: '小于 {a}',
  greaterThanOrEqual: '大于或等于 {a}',
  lessThanOrEqual: '小于或等于 {a}',
};

/**
 * 输入的值是否符合验证。value = 识别后的值（数字 / 文字 / 布尔 / null 为空）。符合返回 null，否则返回原因。
 */
export function checkValue(v: Validation, value: string | number | boolean | null, book: Workbook, si: number, r: number, c: number, engine: Engine | null): string | null {
  if (value === null || value === '') return v.allowBlank ? null : '这个格子不能为空';
  const num = (f: string | undefined): number => {
    if (f === undefined) return NaN;
    const n = Number(f);
    if (Number.isFinite(n)) return n;
    const x = engine?.evalFormula(f, si, r, c);
    return typeof x === 'number' ? x : NaN;
  };
  const op = v.operator ?? 'between';
  const a = num(v.formula1);
  const b = num(v.formula2);
  const opText = (OP_TEXT[op] ?? '').replace('{a}', String(a)).replace('{b}', String(b));
  switch (v.type) {
    case 'list': {
      const items = listItems(v, book, si);
      const t = String(typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value).toLowerCase();
      return items.some((x) => x.toLowerCase() === t) ? null : `只能从下拉列表里选：${items.slice(0, 8).join('、')}${items.length > 8 ? ' …' : ''}`;
    }
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time': {
      if (typeof value !== 'number') return v.type === 'date' ? '请输入日期' : '请输入数字';
      if (v.type === 'whole' && !Number.isInteger(value)) return '请输入整数';
      return OPS[op]?.(value, a, b) === false ? `数值应${opText}` : null;
    }
    case 'textLength': {
      const len = [...String(value)].length;
      return OPS[op]?.(len, a, b) === false ? `文字长度应${opText}` : null;
    }
    default:
      return null;
  }
}

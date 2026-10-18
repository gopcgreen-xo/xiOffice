/**
 * 表单的纯逻辑：把表单定义（工作包里的 forms/<id>.json）和列属性合成“字段”、算默认值、校验、
 * 提交时写进表格（新建一行或改一行）。界面在 ui/FormFill.tsx、ui/FormDesigner.tsx。
 */
import {
  checkColumnValue,
  formatDateTime,
  resolveDefault,
  type AutoValue,
  type ColumnMeta,
  type FormDef,
  type SheetColumns,
} from '@wz/editor-shared';
import { cellText } from './colmeta';
import { literalInput, setColumnInput } from './columns';
import type { Tx } from './ops';
import type { SheetData } from './workbook';

/** 表单里的一个字段：列属性 + 表单上对这个字段的覆盖 */
export interface ResolvedField {
  meta: ColumnMeta;
  label: string;
  hint?: string;
  hidden: boolean;
  auto?: AutoValue;
  /** 提交后保留这一项的值（沿用） */
  keep: boolean;
}

export type FormValues = Record<string, string>;

/** 没有保存过表单时用的默认表单：这张表所有已定义列属性的列，按列顺序 */
export function defaultFormFields(sc: SheetColumns | undefined): ResolvedField[] {
  return (sc?.columns ?? [])
    .filter((c) => !c.unbound)
    .sort((a, b) => a.col - b.col)
    .map((meta) => ({ meta, label: meta.title || `列 ${meta.col + 1}`, hidden: false, keep: false, ...(meta.type === 'autonum' ? { auto: 'autonum' as const } : meta.type === 'created' ? { auto: 'now' as const } : {}) }));
}

/** 表单定义 → 字段；引用不到列（被删 / 未绑定）的字段跳过 */
export function resolveForm(sc: SheetColumns | undefined, form: FormDef | null): ResolvedField[] {
  if (!form) return defaultFormFields(sc);
  const out: ResolvedField[] = [];
  for (const f of form.fields) {
    const meta = sc?.columns.find((c) => c.id === f.column && !c.unbound);
    if (!meta) continue;
    out.push({
      meta,
      label: f.label || meta.title || `列 ${meta.col + 1}`,
      ...(f.hint && { hint: f.hint }),
      hidden: !!f.hidden,
      ...(f.auto && { auto: f.auto }),
      keep: !!f.keep,
    });
  }
  return out;
}

/** 新记录写到哪一行：表头下面第一个空行（已用区域之后） */
export function nextRow(sheet: SheetData): number {
  const header = sheet.columns?.header ?? 1;
  return Math.max(header, sheet.rows);
}

/** 本列已有数值的最大值 + 1（自动编号）；没有数字时为 1 */
export function nextAutoNumber(sheet: SheetData, col: number): number {
  let max = 0;
  const header = sheet.columns?.header ?? 1;
  for (let r = header; r < sheet.rows; r++) {
    const v = sheet.cells[r]?.[col]?.v;
    if (typeof v === 'number' && v > max) max = v;
  }
  return Math.floor(max) + 1;
}

/** 新建记录时各字段的初始值：默认值（展开 {{today}} 等）；保留上一条里标了“沿用”的值 */
export function initialValues(fields: ResolvedField[], user: string | null, now: Date, previous?: FormValues): FormValues {
  const out: FormValues = {};
  for (const f of fields) {
    const id = f.meta.id;
    if (f.keep && previous && previous[id] !== undefined) out[id] = previous[id];
    else out[id] = f.meta.default ? resolveDefault(f.meta.default, now, user) : '';
  }
  return out;
}

/** 校验：返回 { 列 id: 原因 }；自动填的字段和隐藏字段不检查必填 */
export function validateValues(fields: ResolvedField[], values: FormValues): Record<string, string> {
  const errs: Record<string, string> = {};
  for (const f of fields) {
    if (f.auto) continue;
    const text = values[f.meta.id] ?? '';
    const msg = checkColumnValue(f.meta, text);
    if (msg && !(f.hidden && !text.trim())) errs[f.meta.id] = msg;
  }
  return errs;
}

/** 某一行里各字段现在的文字（编辑当前行时回填） */
export function rowValues(sheet: SheetData, fields: ResolvedField[], r: number): FormValues {
  const out: FormValues = {};
  for (const f of fields) out[f.meta.id] = cellText(sheet, r, f.meta.col);
  return out;
}

export type FormTarget = { kind: 'new' } | { kind: 'edit'; row: number };

/**
 * 提交：新建 = 写到下一个空行（自动字段取自动值，自动筛选范围跟着扩一行）；编辑 = 只改这一行的非自动字段。
 * 返回写入的行号。调用方用 run(...) 包住，所以一次提交是一步撤销。
 */
export function submitForm(tx: Tx, si: number, fields: ResolvedField[], values: FormValues, target: FormTarget, user: string | null, now = new Date()): { row: number; written: FormValues } {
  const sheet = tx.book.sheets[si];
  const isNew = target.kind === 'new';
  const r = target.kind === 'new' ? nextRow(sheet) : target.row;
  const written: FormValues = {};
  // 自动编号：同一次提交里按字段顺序取，互不影响
  for (const f of fields) {
    let text = values[f.meta.id] ?? '';
    if (f.auto) {
      if (!isNew) continue;
      text = f.auto === 'now' ? formatDateTime(now) : f.auto === 'user' ? (user ?? '') : String(nextAutoNumber(sheet, f.meta.col));
    } else if (f.meta.type === 'autonum' || f.meta.type === 'created') {
      if (!isNew) continue;
      text = f.meta.type === 'created' ? formatDateTime(now) : String(nextAutoNumber(sheet, f.meta.col));
    }
    written[f.meta.id] = text;
    setColumnInput(tx, si, r, f.meta.col, f.meta, text);
  }
  // 事务提交前先把已用区域撑大（同一个事务里连续提交多条时，下一条要接在这一条后面）
  sheet.rows = Math.max(sheet.rows, r + 1);
  for (const f of fields) sheet.cols = Math.max(sheet.cols, f.meta.col + 1);
  const af = sheet.autoFilter;
  if (isNew && af && r === af.ref.r1 + 1 && fields.some((f) => f.meta.col >= af.ref.c0 && f.meta.col <= af.ref.c1)) {
    tx.meta(si);
    sheet.autoFilter = { ...af, ref: { ...af.ref, r1: r } };
  }
  return { row: r, written };
}

/** 表单设计器里用：给没有列属性的列生成一个文本列属性（表单的字段必须引用列属性） */
export function textMetaFor(col: number, title: string, existing: { id: string }[]): ColumnMeta {
  const used = new Set(existing.map((c) => c.id));
  let n = existing.length + 1;
  while (used.has(`c${n}`)) n++;
  return { id: `c${n}`, title, col, type: 'text' };
}

// ---------- 离线填写结果的导入 ----------

export interface FormResult {
  v: number;
  form: string;
  pack?: string;
  records: Record<string, string>[];
}

export const MAX_IMPORT = 5000;

/** 读离线填写页导出的 .json；不是填写结果或太大时抛出能看懂的错误 */
export function parseFormResult(text: string): FormResult {
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    throw new Error('不是有效的填写结果文件（JSON 无法解析）');
  }
  const o = raw && typeof raw === 'object' && !Array.isArray(raw) ? (raw as Record<string, unknown>) : null;
  if (!o || typeof o.form !== 'string' || !Array.isArray(o.records)) throw new Error('不是离线填写页导出的填写结果');
  if (o.records.length > MAX_IMPORT) throw new Error(`一次最多导入 ${MAX_IMPORT} 条`);
  const records = o.records.map((r) => {
    const out: Record<string, string> = {};
    if (r && typeof r === 'object' && !Array.isArray(r)) {
      for (const [k, v] of Object.entries(r as Record<string, unknown>)) {
        if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') out[k] = String(v);
      }
    }
    return out;
  });
  return { v: typeof o.v === 'number' ? o.v : 1, form: o.form, ...(typeof o.pack === 'string' && { pack: o.pack }), records };
}

/**
 * 把填写结果逐条写进表格（新建记录）：每条按列属性校验，不合规的整条跳过并说明原因；
 * 结果文件来自别人，写进去的文字不当公式（以 = 开头的当普通文字）。
 */
export function importFormResult(
  tx: Tx,
  si: number,
  fields: ResolvedField[],
  result: FormResult,
  user: string | null,
  now = new Date(),
): { imported: number; skipped: { index: number; reason: string }[] } {
  const skipped: { index: number; reason: string }[] = [];
  let imported = 0;
  result.records.forEach((rec, i) => {
    const values = initialValues(fields, user, now);
    for (const f of fields) {
      const v = rec[f.meta.id];
      if (v !== undefined && !f.auto) values[f.meta.id] = v;
    }
    const errs = validateValues(fields, values);
    const first = fields.find((f) => errs[f.meta.id]);
    if (first) {
      skipped.push({ index: i + 1, reason: `${first.label}：${errs[first.meta.id]}` });
      return;
    }
    for (const k of Object.keys(values)) if (values[k]!.startsWith('=')) values[k] = literalInput(values[k]!);
    submitForm(tx, si, fields, values, { kind: 'new' }, user, now);
    imported++;
  });
  return { imported, skipped };
}

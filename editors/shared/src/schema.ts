/**
 * 工作包里的“列属性”和“表单”：类型、清洗（读入的 JSON 一律经过这里，不信任文件内容）和纯函数。
 * 存放位置见 docs/SHEET_COLUMNS.md：`meta/<条目id>.columns.json`、`forms/<表单id>.json`。
 * 表格编辑器和工作台都用这些类型；它们不依赖 DOM。
 */

export const COLUMN_TYPES = ['text', 'number', 'money', 'percent', 'date', 'select', 'multi', 'check', 'link', 'email', 'phone', 'autonum', 'created'] as const;
export type ColumnType = (typeof COLUMN_TYPES)[number];

export const COLUMN_TYPE_LABELS: Record<ColumnType, string> = {
  text: '文本',
  number: '数字',
  money: '金额',
  percent: '百分比',
  date: '日期',
  select: '单选',
  multi: '多选',
  check: '复选框',
  link: '链接',
  email: '邮箱',
  phone: '电话',
  autonum: '自动编号',
  created: '创建时间',
};

/** 表头上的类型小图标（和原型一致） */
export const COLUMN_TYPE_GLYPHS: Record<ColumnType, string> = {
  text: 'Aa',
  number: '#',
  money: '¥',
  percent: '%',
  date: '◷',
  select: '▾',
  multi: '☰',
  check: '✓',
  link: '↗',
  email: '@',
  phone: '☎',
  autonum: '№',
  created: '⏱',
};

/** 单选 / 多选的选项；color 是 #rrggbb */
export interface SelectOption {
  v: string;
  color?: string;
  /** 以 x- 开头的扩展字段（原样保留，见 docs/WBP_FORMAT.md 第 10 节） */
  ext?: Record<string, unknown>;
}

export interface ColumnMeta {
  /** 稳定的列 id（表头改名、插删列后不变） */
  id: string;
  /** 表头文字（打开文件时用它重新匹配列） */
  title: string;
  /** 当前列序号，0 起 */
  col: number;
  type: ColumnType;
  options?: SelectOption[];
  required?: boolean;
  /** 默认值（表单新建记录时预填；文字形式） */
  default?: string;
  /** 数字 / 金额 / 百分比 / 日期的显示格式（Excel 格式码），缺省按类型 */
  fmt?: string;
  min?: number;
  max?: number;
  /** 数字只允许整数 */
  integer?: boolean;
  /** 说明文字（表单里显示） */
  note?: string;
  /** 输入不合规时：stop 拒绝（缺省）/ warn 警告后可继续 / info 仅提示 */
  onInvalid?: 'stop' | 'warn' | 'info';
  /** 单选 / 多选的选项来自一个区域（如 `Sheet2!A2:A20` 或 `A2:A20`），有它时 options 只作显示色缓存 */
  optionsRef?: string;
  ext?: Record<string, unknown>;
  /** 运行期标记（不写进文件）：打开时按表头文字找不到对应的列，暂不套用到任何列 */
  unbound?: boolean;
  /** 本程序不认识的列类型（读入时当文本列处理，写回时原样恢复；用户改了类型就清掉它） */
  unknownType?: string;
  /** 不认识的类型的其它字段，写回时原样带上 */
  raw?: Record<string, unknown>;
}

export interface SheetColumns {
  /** 表头所在行，1 起 */
  header: number;
  columns: ColumnMeta[];
  ext?: Record<string, unknown>;
}

export interface ColumnsFile {
  v: 1;
  /** 工作表名 → 该表的列属性 */
  sheets: Record<string, SheetColumns>;
  ext?: Record<string, unknown>;
}

export type AutoValue = 'now' | 'user' | 'autonum';

export interface FormField {
  /** 引用的列 id */
  column: string;
  label?: string;
  hint?: string;
  hidden?: boolean;
  auto?: AutoValue;
  /** 提交后保留这一项的值（沿用） */
  keep?: boolean;
  ext?: Record<string, unknown>;
}

export interface FormDef {
  id: string;
  title: string;
  target: { item: string; sheet: string };
  fields: FormField[];
  /** 提交后：继续新建 / 关闭 */
  after?: 'continue' | 'close';
  /** 提交后运行的脚本条目 id（需要用户信任，见脚本一节） */
  onSubmit?: string;
  /** 提交前运行的校验脚本条目 id：不通过就不写入（见 docs/WBP_FORMAT.md 第 6.3 节） */
  onValidate?: string;
  ext?: Record<string, unknown>;
}

// ---------- 联动流程（首页） ----------

/** 流程里的一步：打开表单录入 / 打开某个条目 / 运行脚本 */
export interface FlowStep {
  kind: 'form' | 'open' | 'script';
  /** 表单 id（form）或条目 id（open / script） */
  ref: string;
  label?: string;
  /** script 步骤传给脚本的参数（脚本里用 wbp.input.args 读取） */
  args?: Record<string, string>;
  ext?: Record<string, unknown>;
}

/** 联动流程：工作包首页上的一张“流程卡”，把表单录入、校验、生成等按顺序摆出来，供使用者一步步点 */
export interface FlowDef {
  id: string;
  title: string;
  description?: string;
  steps: FlowStep[];
  ext?: Record<string, unknown>;
}

export const MAX_FLOW_STEPS = 30;

// ---------- 清洗 ----------

const isObj = (v: unknown): v is Record<string, unknown> => !!v && typeof v === 'object' && !Array.isArray(v);
const str = (v: unknown, max = 200): string => (typeof v === 'string' ? v.slice(0, max) : '');
const finite = (v: unknown): number | undefined => (typeof v === 'number' && Number.isFinite(v) ? v : undefined);
/** 颜色只认十六进制 */
export const safeOptionColor = (v: unknown): string | undefined => (typeof v === 'string' && /^#[0-9a-fA-F]{6}$/.test(v) ? v : undefined);
/** 取出以 x- 开头的扩展字段 */
function extOf(o: Record<string, unknown>): Record<string, unknown> | undefined {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(o)) if (/^x-[A-Za-z0-9_.-]{1,60}$/.test(k)) out[k] = v;
  return Object.keys(out).length ? out : undefined;
}

const safeId = (v: unknown): string | null => (typeof v === 'string' && /^[A-Za-z0-9_-]{1,40}$/.test(v) ? v : null);

export const MAX_OPTIONS = 200;
export const MAX_COLUMNS = 2000;
const MAX_FIELDS = 500;

export function cleanOptions(v: unknown): SelectOption[] {
  if (!Array.isArray(v)) return [];
  const seen = new Set<string>();
  const out: SelectOption[] = [];
  for (const o of v) {
    const text = typeof o === 'string' ? o : isObj(o) ? str(o.v, 100) : '';
    const trimmed = text.trim();
    if (!trimmed || seen.has(trimmed) || out.length >= MAX_OPTIONS) continue;
    seen.add(trimmed);
    const color = isObj(o) ? safeOptionColor(o.color) : undefined;
    const ext = isObj(o) ? extOf(o) : undefined;
    out.push({ v: trimmed, ...(color ? { color } : {}), ...(ext ? { ext } : {}) });
  }
  return out;
}

export function cleanColumn(v: unknown): ColumnMeta | null {
  if (!isObj(v)) return null;
  const id = safeId(v.id);
  const known = COLUMN_TYPES.find((t) => t === v.type);
  const col = finite(v.col);
  if (!id || typeof v.type !== 'string' || col === undefined || col < 0 || col > 16383 || !Number.isInteger(col)) return null;
  const type: ColumnType = known ?? 'text';
  const c: ColumnMeta = { id, title: str(v.title), col, type };
  // 不认识的列类型：当文本列显示，但记下原类型和其它字段，写回时原样恢复，不删除不改写
  if (!known) {
    c.unknownType = v.type.slice(0, 40);
    c.raw = Object.fromEntries(Object.entries(v).filter(([k]) => !['id', 'title', 'col', 'type'].includes(k)));
  }
  const ext = extOf(v);
  if (ext) c.ext = ext;
  if (type === 'select' || type === 'multi') c.options = cleanOptions(v.options);
  if (v.required === true) c.required = true;
  if (typeof v.default === 'string' && v.default !== '') c.default = v.default.slice(0, 500);
  if (typeof v.fmt === 'string' && v.fmt !== '') c.fmt = v.fmt.slice(0, 100);
  const min = finite(v.min);
  const max = finite(v.max);
  if (min !== undefined) c.min = min;
  if (max !== undefined) c.max = max;
  if (v.integer === true) c.integer = true;
  if (typeof v.note === 'string' && v.note !== '') c.note = v.note.slice(0, 300);
  if (v.onInvalid === 'stop' || v.onInvalid === 'warn' || v.onInvalid === 'info') c.onInvalid = v.onInvalid;
  if (typeof v.optionsRef === 'string' && v.optionsRef !== '') c.optionsRef = v.optionsRef.slice(0, 200);
  return c;
}

export function cleanSheetColumns(v: unknown): SheetColumns | null {
  if (!isObj(v) || !Array.isArray(v.columns)) return null;
  const header = finite(v.header);
  const ids = new Set<string>();
  const cols = new Set<number>();
  const columns: ColumnMeta[] = [];
  for (const raw of v.columns) {
    const c = cleanColumn(raw);
    // id、列序号都不能重复
    if (!c || ids.has(c.id) || cols.has(c.col) || columns.length >= MAX_COLUMNS) continue;
    ids.add(c.id);
    cols.add(c.col);
    columns.push(c);
  }
  const sc: SheetColumns = { header: header !== undefined && Number.isInteger(header) && header >= 1 ? header : 1, columns };
  const ext = extOf(v);
  if (ext) sc.ext = ext;
  return sc;
}

export function cleanColumnsFile(v: unknown): ColumnsFile {
  const out: ColumnsFile = { v: 1, sheets: {} };
  if (!isObj(v) || !isObj(v.sheets)) return out;
  const fileExt = extOf(v);
  if (fileExt) out.ext = fileExt;
  for (const [name, sc] of Object.entries(v.sheets)) {
    const clean = cleanSheetColumns(sc);
    if (clean && name.length <= 100) out.sheets[name] = clean;
  }
  return out;
}

export function cleanForm(v: unknown): FormDef | null {
  if (!isObj(v)) return null;
  const id = safeId(v.id);
  const t = v.target;
  if (!id || !isObj(t)) return null;
  const item = safeId(t.item);
  const sheet = str(t.sheet, 100);
  if (!item || !sheet || !Array.isArray(v.fields)) return null;
  const seen = new Set<string>();
  const fields: FormField[] = [];
  for (const f of v.fields) {
    if (!isObj(f) || fields.length >= MAX_FIELDS) continue;
    const column = safeId(f.column);
    if (!column || seen.has(column)) continue;
    seen.add(column);
    const field: FormField = { column };
    if (typeof f.label === 'string' && f.label) field.label = f.label.slice(0, 100);
    if (typeof f.hint === 'string' && f.hint) field.hint = f.hint.slice(0, 300);
    if (f.hidden === true) field.hidden = true;
    if (f.auto === 'now' || f.auto === 'user' || f.auto === 'autonum') field.auto = f.auto;
    if (f.keep === true) field.keep = true;
    const fext = extOf(f);
    if (fext) field.ext = fext;
    fields.push(field);
  }
  const form: FormDef = { id, title: str(v.title, 100) || '表单', target: { item, sheet }, fields };
  if (v.after === 'continue' || v.after === 'close') form.after = v.after;
  const onSubmit = safeId(v.onSubmit);
  if (onSubmit) form.onSubmit = onSubmit;
  const onValidate = safeId(v.onValidate);
  if (onValidate) form.onValidate = onValidate;
  const ext = extOf(v);
  if (ext) form.ext = ext;
  return form;
}

// ---------- 写出（把 ext 展开成 x- 字段，字段顺序固定） ----------

function optionJson(o: SelectOption): Record<string, unknown> {
  return { v: o.v, ...(o.color ? { color: o.color } : {}), ...o.ext };
}

function columnJson(c: ColumnMeta): Record<string, unknown> {
  const out: Record<string, unknown> = { id: c.id, title: c.title, col: c.col, type: c.unknownType ?? c.type };
  if (c.unknownType) Object.assign(out, c.raw);
  if (c.options && !c.unknownType) out.options = c.options.map(optionJson);
  if (!c.unknownType) {
    if (c.required) out.required = true;
    if (c.default !== undefined) out.default = c.default;
    if (c.fmt !== undefined) out.fmt = c.fmt;
    if (c.min !== undefined) out.min = c.min;
    if (c.max !== undefined) out.max = c.max;
    if (c.integer) out.integer = true;
    if (c.note !== undefined) out.note = c.note;
    if (c.onInvalid) out.onInvalid = c.onInvalid;
    if (c.optionsRef) out.optionsRef = c.optionsRef;
  }
  return { ...out, ...c.ext };
}

export function columnsFileJson(f: ColumnsFile): Record<string, unknown> {
  const sheets: Record<string, unknown> = {};
  for (const [name, sc] of Object.entries(f.sheets)) {
    sheets[name] = { header: sc.header, columns: sc.columns.map(columnJson), ...sc.ext };
  }
  return { v: 1, sheets, ...f.ext };
}

export function formJson(f: FormDef): Record<string, unknown> {
  return {
    id: f.id,
    title: f.title,
    target: { item: f.target.item, sheet: f.target.sheet },
    fields: f.fields.map((x) => ({
      column: x.column,
      ...(x.label ? { label: x.label } : {}),
      ...(x.hint ? { hint: x.hint } : {}),
      ...(x.hidden ? { hidden: true } : {}),
      ...(x.auto ? { auto: x.auto } : {}),
      ...(x.keep ? { keep: true } : {}),
      ...x.ext,
    })),
    ...(f.after ? { after: f.after } : {}),
    ...(f.onSubmit ? { onSubmit: f.onSubmit } : {}),
    ...(f.onValidate ? { onValidate: f.onValidate } : {}),
    ...f.ext,
  };
}

/** 清洗联动流程（文件里读出来的 JSON → FlowDef）；不合规的返回 null */
export function cleanFlow(v: unknown): FlowDef | null {
  if (!isObj(v)) return null;
  const id = safeId(v.id);
  if (!id || !Array.isArray(v.steps)) return null;
  const steps: FlowStep[] = [];
  for (const r of v.steps) {
    if (!isObj(r) || steps.length >= MAX_FLOW_STEPS) continue;
    const ref = safeId(r.ref);
    if (!ref || (r.kind !== 'form' && r.kind !== 'open' && r.kind !== 'script')) continue;
    const step: FlowStep = { kind: r.kind, ref };
    if (typeof r.label === 'string' && r.label) step.label = r.label.slice(0, 100);
    if (isObj(r.args)) {
      const args: Record<string, string> = {};
      for (const [k, x] of Object.entries(r.args)) {
        if (Object.keys(args).length >= 50) break;
        if (k && k.length <= 100 && (typeof x === 'string' || typeof x === 'number' || typeof x === 'boolean')) args[k] = String(x).slice(0, 2000);
      }
      if (Object.keys(args).length) step.args = args;
    }
    const ext = extOf(r);
    if (ext) step.ext = ext;
    steps.push(step);
  }
  const flow: FlowDef = { id, title: str(v.title, 100) || '流程', steps };
  if (typeof v.description === 'string' && v.description) flow.description = v.description.slice(0, 500);
  const ext = extOf(v);
  if (ext) flow.ext = ext;
  return flow;
}

export function flowJson(f: FlowDef): Record<string, unknown> {
  return {
    id: f.id,
    title: f.title,
    ...(f.description ? { description: f.description } : {}),
    steps: f.steps.map((x) => ({ kind: x.kind, ref: x.ref, ...(x.label ? { label: x.label } : {}), ...(x.args ? { args: x.args } : {}), ...x.ext })),
    ...f.ext,
  };
}

// ---------- 纯函数 ----------

export function newColumnId(existing: { id: string }[]): string {
  const used = new Set(existing.map((c) => c.id));
  for (let n = existing.length + 1; ; n++) if (!used.has(`c${n}`)) return `c${n}`;
}

/** 选项的颜色（没设置时按序号取一个固定色板，标签颜色稳定） */
export const OPTION_PALETTE = ['#3f5ef0', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#84cc16'];

export function optionColor(opts: SelectOption[] | undefined, value: string): string | undefined {
  const i = opts?.findIndex((o) => o.v === value) ?? -1;
  if (i < 0) return undefined;
  return opts![i].color ?? OPTION_PALETTE[i % OPTION_PALETTE.length];
}

/** 多选的单元格文字 ⇄ 选项数组（逗号分隔，中英文逗号、顿号都认） */
export function splitMulti(text: string): string[] {
  return text
    .split(/[,，、;；]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

export function joinMulti(values: string[]): string {
  return values.join(',');
}

export const CHECK_ON = '是';
export const CHECK_OFF = '否';

export function isCheckOn(text: string): boolean {
  return /^(是|true|yes|y|1|√|✓|☑)$/i.test(text.trim());
}

/** 一格的值是否合规；返回 null 表示通过，否则是原因 */
export function checkColumnValue(c: ColumnMeta, text: string): string | null {
  const t = text.trim();
  if (t === '') return c.required ? '必填' : null;
  switch (c.type) {
    case 'number':
    case 'money':
    case 'percent': {
      const n = parseNumberText(t, c.type === 'percent');
      if (n === null) return '请输入数字';
      if (c.integer && !Number.isInteger(n)) return '只允许整数';
      if (c.min !== undefined && n < c.min) return `不能小于 ${c.min}`;
      if (c.max !== undefined && n > c.max) return `不能大于 ${c.max}`;
      return null;
    }
    case 'date':
      return parseDateText(t) === null ? '请输入日期（如 2026-10-12）' : null;
    case 'select': {
      const opts = c.options ?? [];
      return opts.length && !opts.some((o) => o.v === t) ? '不在选项里' : null;
    }
    case 'multi': {
      const opts = c.options ?? [];
      const bad = opts.length ? splitMulti(t).find((v) => !opts.some((o) => o.v === v)) : undefined;
      return bad !== undefined ? `「${bad}」不在选项里` : null;
    }
    case 'check':
      return /^(是|否|true|false|yes|no|y|n|1|0|√|✓|☑|☐)$/i.test(t) ? null : '只能是「是」或「否」';
    case 'email':
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t) ? null : '邮箱格式不对';
    case 'phone':
      return /^[+\d][\d\s\-()]{4,}$/.test(t) ? null : '电话格式不对';
    case 'link':
      return /^(https?:\/\/|mailto:|tel:)\S+$/i.test(t) ? null : '链接要以 http:// 或 https:// 开头';
    default:
      return null;
  }
}

/** 数字文字：去千分位、货币符号；百分比允许 12% */
export function parseNumberText(text: string, percent = false): number | null {
  let t = text.trim().replace(/[¥￥$,，\s]/g, '');
  if (t === '') return null;
  let scale = 1;
  if (t.endsWith('%')) {
    t = t.slice(0, -1);
    scale = percent ? 0.01 : 1;
  }
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return null;
  const n = Number(t) * scale;
  return Number.isFinite(n) ? n : null;
}

/** 日期文字 → Date（本地零点）；认 2026-10-12、2026/10/12、2026年10月12日、20261012 */
export function parseDateText(text: string): Date | null {
  const t = text.trim();
  const m = /^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?$/.exec(t) ?? /^(\d{4})(\d{2})(\d{2})$/.exec(t);
  if (!m) return null;
  const y = +m[1];
  const mo = +m[2];
  const d = +m[3];
  const date = new Date(y, mo - 1, d);
  return date.getFullYear() === y && date.getMonth() === mo - 1 && date.getDate() === d ? date : null;
}

export function formatDate(d: Date): string {
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

export function formatDateTime(d: Date): string {
  const p = (n: number) => String(n).padStart(2, '0');
  return `${formatDate(d)} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

export function columnAt(sc: SheetColumns | undefined, col: number): ColumnMeta | undefined {
  return sc?.columns.find((c) => c.col === col && !c.unbound);
}

/** 列类型是否由用户输入（自动编号、创建时间由表单 / 提交时填） */
export const isAutoType = (t: ColumnType): boolean => t === 'autonum' || t === 'created';

/**
 * 默认值里的变量：`{{today}}` 今天，`{{today+7}}` / `{{today-1}}` 前后几天，`{{now}}` 当前时间，`{{user}}` 当前用户。
 * 其它文字原样返回。
 */
export function resolveDefault(text: string, now: Date, user: string | null): string {
  return text.replace(/\{\{\s*(today|now|user)\s*([+-]\s*\d+)?\s*\}\}/g, (_m, name: string, off?: string) => {
    if (name === 'user') return user ?? '';
    if (name === 'now') return formatDateTime(now);
    const d = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    if (off) d.setDate(d.getDate() + Number(off.replace(/\s+/g, '')));
    return formatDate(d);
  });
}

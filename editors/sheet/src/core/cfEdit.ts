/**
 * 条件格式的新建 / 删除（显示在 cf.ts）：新建的规则保存时才写进文件，文件里原有的规则（含图标集、x14 扩展等
 * 本编辑器不建模的）原样保留，只有被删掉的才从原 XML 里去掉（见 save.ts 的 removeCfRules / addCfRules）。
 */
import { colName } from '@wz/editor-shared/xlsx';
import type { Tx } from './ops';
import { dxfForStyle } from './styles';
import type { CfRule, Range, Workbook } from './workbook';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const argb = (c: string) => `FF${c.slice(1).toUpperCase()}`;

export const cellRef = (r: number, c: number) => `${colName(c)}${r + 1}`;
export const rangeRef = (g: Range) => (g.r0 === g.r1 && g.c0 === g.c1 ? cellRef(g.r0, g.c0) : `${cellRef(g.r0, g.c0)}:${cellRef(g.r1, g.c1)}`);

const cfvoXml = (v: { type: string; val?: string }) => `<cfvo type="${v.type}"${v.val !== undefined ? ` val="${esc(v.val)}"` : ''}/>`;

/** 一条规则 → <cfRule> */
export function cfRuleXml(rule: CfRule): string {
  const first = rule.ranges[0];
  const topLeft = first ? cellRef(first.r0, first.c0) : 'A1';
  let a = `type="${rule.type}"`;
  if (rule.dxf !== undefined) a += ` dxfId="${rule.dxf}"`;
  a += ` priority="${rule.priority}"`;
  if (rule.stopIfTrue) a += ' stopIfTrue="1"';
  if (rule.type === 'aboveAverage' && rule.above === false) a += ' aboveAverage="0"';
  if (rule.percent) a += ' percent="1"';
  if (rule.bottom) a += ' bottom="1"';
  if (rule.rank !== undefined) a += ` rank="${rule.rank}"`;
  const textOp: Record<string, string> = { containsText: 'containsText', notContainsText: 'notContains', beginsWith: 'beginsWith', endsWith: 'endsWith' };
  const operator = rule.operator ?? textOp[rule.type];
  if (operator) a += ` operator="${operator}"`;
  if (rule.text !== undefined) a += ` text="${esc(rule.text)}"`;
  if (rule.type === 'colorScale') {
    const cols = rule.colors ?? [];
    return `<cfRule ${a}><colorScale>${(rule.cfvo ?? []).map(cfvoXml).join('')}${cols.map((c) => `<color rgb="${argb(c)}"/>`).join('')}</colorScale></cfRule>`;
  }
  if (rule.type === 'dataBar') {
    return `<cfRule ${a}><dataBar>${(rule.cfvo ?? []).map(cfvoXml).join('')}<color rgb="${argb(rule.colors?.[0] ?? '#638ec6')}"/></dataBar></cfRule>`;
  }
  let formulas = rule.formulas;
  if (rule.text !== undefined && rule.type in textOp) {
    const t = `"${rule.text.replace(/"/g, '""')}"`;
    const f: Record<string, string> = {
      containsText: `NOT(ISERROR(SEARCH(${t},${topLeft})))`,
      notContainsText: `ISERROR(SEARCH(${t},${topLeft}))`,
      beginsWith: `LEFT(${topLeft},LEN(${t}))=${t}`,
      endsWith: `RIGHT(${topLeft},LEN(${t}))=${t}`,
    };
    formulas = [f[rule.type]];
  }
  return formulas.length ? `<cfRule ${a}>${formulas.map((f) => `<formula>${esc(f)}</formula>`).join('')}</cfRule>` : `<cfRule ${a}/>`;
}

// ---------- 预设 ----------

export interface CfLook {
  label: string;
  fill?: string;
  color?: string;
  bold?: boolean;
}

/** 突出显示用的格式（同 Excel 的预设） */
export const CF_LOOKS: CfLook[] = [
  { label: '浅红色填充深红色文本', fill: '#FFC7CE', color: '#9C0006' },
  { label: '黄填充深黄色文本', fill: '#FFEB9C', color: '#9C5700' },
  { label: '绿填充深绿色文本', fill: '#C6EFCE', color: '#006100' },
  { label: '浅红色填充', fill: '#FFC7CE' },
  { label: '红色文本', color: '#9C0006' },
  { label: '加粗红色文本', color: '#C00000', bold: true },
];

/** 色阶（从小到大的颜色） */
export const CF_SCALES: { label: string; colors: string[] }[] = [
  { label: '绿 - 黄 - 红', colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
  { label: '红 - 黄 - 绿', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  { label: '白 - 红', colors: ['#FCFCFF', '#F8696B'] },
  { label: '白 - 绿', colors: ['#FCFCFF', '#63BE7B'] },
  { label: '白 - 蓝', colors: ['#FCFCFF', '#5A8AC6'] },
];

export const CF_BAR_COLORS: { label: string; color: string }[] = [
  { label: '蓝色', color: '#638EC6' },
  { label: '绿色', color: '#63C384' },
  { label: '红色', color: '#FF555A' },
  { label: '橙色', color: '#FFB628' },
];

export const CF_OPERATORS: { op: string; label: string }[] = [
  { op: 'greaterThan', label: '大于' },
  { op: 'lessThan', label: '小于' },
  { op: 'between', label: '介于' },
  { op: 'equal', label: '等于' },
  { op: 'greaterThanOrEqual', label: '大于或等于' },
  { op: 'lessThanOrEqual', label: '小于或等于' },
  { op: 'notEqual', label: '不等于' },
  { op: 'notBetween', label: '不介于' },
];

/** 规则的一句话说明（规则列表里显示） */
export function describeRule(rule: CfRule): string {
  const op = CF_OPERATORS.find((x) => x.op === rule.operator)?.label ?? '';
  switch (rule.type) {
    case 'cellIs':
      return `单元格值${op} ${rule.formulas.join(' 和 ')}`;
    case 'expression':
      return `公式：${rule.formulas[0] ?? ''}`;
    case 'containsText':
      return `包含文本“${rule.text ?? ''}”`;
    case 'notContainsText':
      return `不包含文本“${rule.text ?? ''}”`;
    case 'beginsWith':
      return `开头是“${rule.text ?? ''}”`;
    case 'endsWith':
      return `结尾是“${rule.text ?? ''}”`;
    case 'duplicateValues':
      return '重复值';
    case 'uniqueValues':
      return '唯一值';
    case 'top10':
      return `${rule.bottom ? '后' : '前'} ${rule.rank ?? 10} ${rule.percent ? '%' : '项'}`;
    case 'aboveAverage':
      return rule.above === false ? '低于平均值' : '高于平均值';
    case 'colorScale':
      return `色阶（${(rule.colors ?? []).length} 色）`;
    case 'dataBar':
      return '数据条';
    case 'iconSet':
      return '图标集';
    case 'containsBlanks':
      return '空值';
    case 'notContainsBlanks':
      return '非空值';
    case 'containsErrors':
      return '错误值';
    case 'notContainsErrors':
      return '无错误值';
    default:
      return rule.type || '条件格式';
  }
}

export const rangesText = (rule: CfRule) => rule.ranges.map(rangeRef).join(' ');

// ---------- 修改（事务内） ----------

const intersects = (a: Range, b: Range) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** 新建一条规则（优先级排在最后）。差异格式由 look 决定（需要时追加到样式表） */
export function addCfRule(tx: Tx, si: number, rule: Omit<CfRule, 'priority' | 'idx' | 'dxf'> & { look?: CfLook }): void {
  const book: Workbook = tx.book;
  const sheet = book.sheets[si];
  tx.meta(si);
  const { look, ...rest } = rule;
  const next: CfRule = { ...rest, priority: Math.max(0, ...(sheet.cf ?? []).map((x) => x.priority)) + 1 };
  if (look && book.styleBook) next.dxf = dxfForStyle(book.styleBook, look);
  sheet.cf = [...(sheet.cf ?? []), next];
}

/** 删一条规则：文件里原有的记下序号，保存时从 XML 去掉；新建还没保存的直接丢掉 */
export function removeCfRule(tx: Tx, si: number, rule: CfRule): void {
  const sheet = tx.book.sheets[si];
  const cf = sheet.cf ?? [];
  const at = cf.indexOf(rule);
  if (at < 0) return;
  tx.meta(si);
  if (rule.idx !== undefined) sheet.cfRemoved = [...(sheet.cfRemoved ?? []), rule.idx];
  const rest = cf.filter((_, i) => i !== at);
  if (rest.length) sheet.cf = rest;
  else delete sheet.cf;
}

/** 清除与区域有重叠的规则（range 为 null：整个工作表）。返回清除的条数 */
export function clearCfRules(tx: Tx, si: number, range: Range | null): number {
  const hit = (tx.book.sheets[si].cf ?? []).filter((r) => !range || r.ranges.some((g) => intersects(g, range)));
  for (const r of hit) removeCfRule(tx, si, r);
  return hit.length;
}

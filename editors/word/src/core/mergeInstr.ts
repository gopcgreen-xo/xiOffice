/**
 * Word 域指令的解析：MERGEFIELD（字段名 + 开关）和规则域 IF / NEXT / NEXTIF / SKIPIF（可以嵌套 MERGEFIELD）。
 * 纯函数。规则域的指令保存成“带花括号的整串”：嵌套的域写成 { MERGEFIELD 名称 }，如
 * ` IF { MERGEFIELD 金额 } > 1000 "大额" "小额" `。
 */
import type { Opaque } from './model';
import { formatValue } from './mergeFormat';

export type MergeSpec = NonNullable<Opaque['merge']>;

/**
 * MERGEFIELD 字段名 [\b "前缀"] [\f "后缀"] [\# "数字图片"] [\@ "日期图片"] [\* Upper|Lower|FirstCap|Caps|MERGEFORMAT…]。
 * 字段名可以带引号。其它开关忽略。
 */
export function parseMergeInstr(instr: string): MergeSpec | null {
  const m = instr.match(/^\s*MERGEFIELD\s+(?:"([^"]*)"|(\S+))([\s\S]*)$/i);
  if (!m) return null;
  const name = (m[1] ?? m[2]).trim();
  if (!name) return null;
  const spec: MergeSpec = { name };
  for (const s of m[3].matchAll(/\\([#@*bf])\s*(?:"([^"]*)"|([^\s\\]+))?/gi)) {
    const arg = s[2] ?? s[3];
    if (arg === undefined) continue;
    switch (s[1].toLowerCase()) {
      case 'b':
        spec.before = arg;
        break;
      case 'f':
        spec.after = arg;
        break;
      case '#':
        spec.numFmt = arg;
        break;
      case '@':
        spec.dateFmt = arg;
        break;
      case '*':
        if (/^(upper|lower|firstcap|caps)$/i.test(arg)) spec.caseFmt = arg;
        break;
    }
  }
  return spec;
}

/** 取值：按开关格式化，值非空时加前缀 / 后缀 */
export function mergeText(spec: MergeSpec, raw: string): string {
  const v = formatValue(raw, spec);
  return v ? `${spec.before ?? ''}${v}${spec.after ?? ''}` : '';
}

// ---------------- 规则域 ----------------

export type RuleOp = '=' | '<>' | '<' | '>' | '<=' | '>=';
export type RuleTok = { t: 'str'; v: string } | { t: 'field'; m: MergeSpec } | { t: 'word'; v: string } | { t: 'op'; v: RuleOp };
export type RuleKind = 'if' | 'next' | 'nextif' | 'skipif';

export interface RuleField {
  kind: RuleKind;
  /** 关键字后面的各项 */
  toks: RuleTok[];
}

/** 把关键字后面的部分切成项：引号字符串、嵌套的 { MERGEFIELD … }、比较符、其它连续字符 */
function tokenize(s: string): RuleTok[] {
  const out: RuleTok[] = [];
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (/\s/.test(c)) i++;
    else if (c === '"' || c === '“' || c === '”') {
      const end = s.slice(i + 1).search(/["“”]/);
      const stop = end < 0 ? s.length : i + 1 + end;
      out.push({ t: 'str', v: s.slice(i + 1, stop) });
      i = stop + 1;
    } else if (c === '{') {
      let depth = 1;
      let j = i + 1;
      while (j < s.length && depth) {
        if (s[j] === '{') depth++;
        else if (s[j] === '}') depth--;
        j++;
      }
      const inner = s.slice(i + 1, depth ? s.length : j - 1);
      const m = parseMergeInstr(inner);
      out.push(m ? { t: 'field', m } : { t: 'word', v: '' });
      i = j;
    } else if (c === '=' || c === '<' || c === '>') {
      const two = s.slice(i, i + 2);
      const op = (['<>', '<=', '>='] as const).find((o) => o === two) ?? (c as RuleOp);
      out.push({ t: 'op', v: op });
      i += op.length;
    } else {
      let j = i;
      while (j < s.length && !/[\s"“”{=<>]/.test(s[j])) j++;
      out.push({ t: 'word', v: s.slice(i, j) });
      i = j;
    }
  }
  return out;
}

/** 规则域指令（嵌套域写成花括号）：IF a op b "真" "假" / NEXT / NEXTIF a op b / SKIPIF a op b；不是返回 null */
export function parseRule(full: string): RuleField | null {
  const m = full.match(/^\s*(IF|NEXTIF|SKIPIF|NEXT)\b([\s\S]*)$/i);
  if (!m) return null;
  const kind = m[1].toLowerCase() as RuleKind;
  const toks = tokenize(m[2]);
  if (kind === 'next') return { kind, toks: [] };
  // IF 要有 a op b；NEXTIF / SKIPIF 也一样
  if (toks.length < 3 || toks[1].t !== 'op') return null;
  return { kind, toks };
}

/** 规则域在编辑器里显示的文字 */
export function ruleLabel(r: RuleField): string {
  if (r.kind === 'next') return '«下一条记录»';
  const show = (t: RuleTok | undefined) => (!t ? '' : t.t === 'field' ? t.m.name : t.t === 'op' ? t.v : t.v);
  const cond = `${show(r.toks[0])}${show(r.toks[1])}${show(r.toks[2])}`;
  return `«${r.kind === 'if' ? '如果' : r.kind === 'nextif' ? '满足则下一条' : '满足则跳过'} ${cond}»`;
}

const numOf = (s: string): number | null => {
  const t = s.trim().replace(/[,，]/g, '');
  return /^[+-]?(\d+\.?\d*|\.\d+)$/.test(t) ? Number(t) : null;
};

/** Word 的比较：两边都是数字按数值，否则按文字（不分大小写；= / <> 时右边的 * ? 是通配符） */
export function compare(a: string, op: RuleOp, b: string): boolean {
  const na = numOf(a);
  const nb = numOf(b);
  if (na !== null && nb !== null) {
    switch (op) {
      case '=': return na === nb;
      case '<>': return na !== nb;
      case '<': return na < nb;
      case '>': return na > nb;
      case '<=': return na <= nb;
      default: return na >= nb;
    }
  }
  if (op === '=' || op === '<>') {
    const re = new RegExp(`^${b.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '[\\s\\S]')}$`, 'i');
    return re.test(a) === (op === '=');
  }
  const c = a.toLowerCase().localeCompare(b.toLowerCase());
  return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
}

export type RuleResult = { kind: 'text'; text: string } | { kind: 'next' } | { kind: 'skip' };

/** 计算规则域：get 返回 MERGEFIELD 的最终文字（已格式化） */
export function evalRule(r: RuleField, get: (m: MergeSpec) => string): RuleResult {
  const val = (t: RuleTok | undefined): string => (!t ? '' : t.t === 'field' ? get(t.m) : t.t === 'op' ? t.v : t.v);
  if (r.kind === 'next') return { kind: 'next' };
  const yes = compare(val(r.toks[0]), (r.toks[1] as { v: RuleOp }).v, val(r.toks[2]));
  if (r.kind === 'if') return { kind: 'text', text: yes ? val(r.toks[3]) : val(r.toks[4]) };
  if (!yes) return { kind: 'text', text: '' };
  return r.kind === 'nextif' ? { kind: 'next' } : { kind: 'skip' };
}

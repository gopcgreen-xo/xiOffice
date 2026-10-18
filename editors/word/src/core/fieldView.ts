/**
 * 域的结构与显示（只依赖模型，排版、读写和 fields.ts 都用）：
 * 段落里成对的开始 / 结束标记（Run.fc）、“显示域代码”的状态、域代码里嵌套花括号的拆分。
 */
import type { Paragraph, Run } from './model';

let seq = 0;
/** 新的域标识（读文件、插入域时用；同一个域的开始和结束标记共用） */
export const newFieldId = (): string => `fld${++seq}`;

// ---------- 显示域代码（Shift+F9 / Alt+F9）：只是显示状态，不进模型，不算修改 ----------

const codeShown = new Set<string>();
let codeEpoch = 0;

export const showsCode = (id: string): boolean => codeShown.has(id);

export function setShowCode(ids: Iterable<string>, on: boolean): void {
  for (const id of ids) {
    if (on) codeShown.add(id);
    else codeShown.delete(id);
  }
  codeEpoch++;
}

/** 显示状态的版本号（排版缓存用） */
export const fieldCodeEpoch = (): number => codeEpoch;

/** 显示域代码时开始标记显示的文字（Word 的样子：{ 域代码 }） */
export const codeText = (instr: string): string => `{ ${instr} }`;

// ---------- 段落里的域 ----------

export interface FieldSpan {
  id: string;
  instr: string;
  locked: boolean;
  /** 开始 / 结束标记在 runs 里的下标 */
  begin: number;
  end: number;
  /** 开始标记所在的字符位置、结束标记之后的字符位置 */
  from: number;
  to: number;
  /** 套在别的域的结果里 */
  depth: number;
}

/** 段落里成对的域（按开始位置排序）。按出现顺序配对（同 Word），配不上对的标记不算 */
export function fieldSpans(p: Paragraph): FieldSpan[] {
  const out: FieldSpan[] = [];
  const stack: { i: number; from: number }[] = [];
  let acc = 0;
  p.runs.forEach((r, i) => {
    if (r.fc?.kind === 'begin') stack.push({ i, from: acc });
    else if (r.fc?.kind === 'end' && stack.length) {
      const b = stack.pop()!;
      const fc = p.runs[b.i].fc!;
      out.push({ id: fc.id, instr: fc.instr ?? '', locked: !!fc.locked, begin: b.i, end: i, from: b.from, to: acc + r.text.length, depth: stack.length });
    }
    acc += r.text.length;
  });
  return out.sort((a, b) => a.begin - b.begin);
}

/** 光标位置 offset 所在的域（最里层）：在开始标记之后、结束标记之后（含）之间 */
export function fieldAt(p: Paragraph, offset: number): FieldSpan | null {
  let best: FieldSpan | null = null;
  for (const f of fieldSpans(p)) if (f.from < offset && offset <= f.to && (!best || f.depth > best.depth)) best = f;
  return best;
}

/** 成对的域标记（写文件时只写这些；删掉一半的域，剩下的结果当普通文字） */
export function pairedMarkers(runs: Run[]): Set<Run> {
  const ok = new Set<Run>();
  for (const f of fieldSpans({ runs } as Paragraph)) {
    ok.add(runs[f.begin]);
    ok.add(runs[f.end]);
  }
  return ok;
}

// ---------- 域代码里的嵌套域 ----------

export type InstrPiece = { t: 'text'; v: string } | { t: 'begin' } | { t: 'end' };

/** 域代码按花括号拆开：嵌套的域写成 {…}（读文件时这样记）；花括号不配对时整串当文字 */
export function instrPieces(instr: string): InstrPiece[] {
  let depth = 0;
  for (const ch of instr) {
    if (ch === '{') depth++;
    else if (ch === '}' && --depth < 0) break;
  }
  if (depth !== 0) return instr ? [{ t: 'text', v: instr }] : [];
  const out: InstrPiece[] = [];
  let buf = '';
  for (const ch of instr) {
    if (ch === '{' || ch === '}') {
      if (buf) out.push({ t: 'text', v: buf });
      buf = '';
      out.push({ t: ch === '{' ? 'begin' : 'end' });
    } else buf += ch;
  }
  if (buf) out.push({ t: 'text', v: buf });
  return out;
}

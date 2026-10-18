/**
 * 名称管理器：定义名称的校验、增删改（带撤销）。
 *
 * 规则同 Excel：名称以字母 / 下划线 / 反斜杠开头（汉字算字母），后面是字母、数字、下划线、点；不能有空格，不能像单元格地址
 * （A1、$B$2、R1C1、R、C），不超过 255 个字；同一范围（整个工作簿 / 某张工作表）内不分大小写不能重名。
 * 改名时，公式和其它名称里对它的引用一起改。名称的“引用位置”是任意公式（区域、常量、公式都可以）。
 */
import { parseFormula } from './formula/parse';
import { parseRef, type Workbook } from './workbook';
import { eachFormula } from './sheets';
import type { Tx } from './ops';

export interface DefinedName {
  name: string;
  /** 引用位置（公式，不带 =） */
  f: string;
  /** 只在这张工作表里有效（工作表序号）；不写 = 整个工作簿 */
  local?: number;
}

const LETTER = /[\p{L}_\\]/u;
const BODY = /^[\p{L}\p{N}_.\\]+$/u;

/** 名称不合规的原因（合规返回 null）。except = 正在编辑的那一项（不和自己比重名） */
export function nameProblem(book: Workbook, name: string, local: number | undefined, except?: DefinedName): string | null {
  if (!name) return '名称不能为空';
  if (name.length > 255) return '名称不能超过 255 个字';
  if (!LETTER.test(name[0])) return '名称要以字母、下划线或汉字开头';
  if (!BODY.test(name)) return '名称里只能有字母、数字、下划线和点（不能有空格）';
  if (parseRef(name) || /^[Rr]\d*[Cc]\d*$/.test(name) || /^[RrCc]$/.test(name)) return '名称不能和单元格地址一样';
  if (/^_xlnm\./i.test(name)) return '名称不能以 _xlnm. 开头';
  const dup = (book.names ?? []).find((n) => n !== except && n.local === local && n.name.toLowerCase() === name.toLowerCase());
  if (dup) return local === undefined ? '这个名称已经存在' : '这张工作表里已经有这个名称';
  return null;
}

/** 引用位置能不能当公式用（语法对、不为空） */
export function refersToProblem(f: string): string | null {
  const t = f.trim().replace(/^=/, '');
  if (!t) return '引用位置不能为空';
  try {
    parseFormula(t);
  } catch (e) {
    return `引用位置的公式不对：${e instanceof Error ? e.message : '语法错误'}`;
  }
  return null;
}

/** 公式里把名称 from 改成 to（跳过字符串、带引号的表名；函数名、单元格引用不动） */
export function renameNameInFormula(f: string, from: string, to: string): string {
  const low = from.toLowerCase();
  let out = '';
  let i = 0;
  while (i < f.length) {
    const q = f[i];
    if (q === '"' || q === "'") {
      let j = i + 1;
      while (j < f.length && !(f[j] === q && f[j + 1] !== q)) j += f[j] === q ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const w = f.slice(i).match(/^[\p{L}_\\][\p{L}\p{N}_.]*/u);
    if (w && !/[\p{L}\p{N}_.!$]/u.test(f[i - 1] ?? '')) {
      const next = f.slice(i + w[0].length).match(/^\s*(.)/)?.[1];
      out += w[0].toLowerCase() === low && next !== '(' && next !== '!' ? to : w[0];
      i += w[0].length;
      continue;
    }
    out += f[i++];
  }
  return out;
}

/** 新增或修改一个名称。replace = 被修改的那一项（改名时公式里的引用一起改）。返回 null 成功，否则是原因 */
export function saveName(tx: Tx, def: DefinedName, replace?: DefinedName): string | null {
  const book = tx.book;
  const f = def.f.trim().replace(/^=/, '');
  const bad = nameProblem(book, def.name, def.local, replace) ?? refersToProblem(f);
  if (bad) return bad;
  tx.wholeBook();
  book.namesDirty = true;
  const next: DefinedName = { name: def.name, f, ...(def.local !== undefined && { local: def.local }) };
  const list = book.names ?? (book.names = []);
  const at = replace ? list.findIndex((n) => n === replace || (n.name === replace.name && n.local === replace.local)) : -1;
  if (at >= 0) {
    const old = list[at];
    list[at] = next;
    if (old.name.toLowerCase() !== next.name.toLowerCase() || old.name !== next.name) {
      // 改名：引用它的公式和别的名称一起改（在它的有效范围里）
      eachFormula(book, (g) => renameNameInFormula(g, old.name, next.name));
    }
  } else list.push(next);
  return null;
}

export function deleteName(tx: Tx, def: DefinedName): void {
  const book = tx.book;
  tx.wholeBook();
  book.namesDirty = true;
  book.names = (book.names ?? []).filter((n) => !(n.name === def.name && n.local === def.local));
}

/** 名称指向一个连续区域（Sheet1!$A$1:$B$5）时返回它的位置，用来在名称框里跳转 / 显示 */
export function nameTarget(book: Workbook, name: string, fromSheet: number): { si: number; range: ReturnType<typeof parseRef> } | null {
  const n = (book.names ?? []).find((x) => x.name.toLowerCase() === name.toLowerCase() && x.local === fromSheet) ?? (book.names ?? []).find((x) => x.name.toLowerCase() === name.toLowerCase() && x.local === undefined);
  if (!n) return null;
  const m = n.f.trim().match(/^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/);
  if (!m) return null;
  const sheetName = (m[1] ?? m[2])?.replace(/''/g, "'");
  const si = sheetName === undefined ? fromSheet : book.sheets.findIndex((s) => s.name.toLowerCase() === sheetName.toLowerCase());
  const range = parseRef(m[3].replace(/\$/g, ''));
  return si < 0 || !range ? null : { si, range };
}

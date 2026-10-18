/**
 * EQ 域（Word 的公式域）的显示：没有缓存结果时 Word 按域代码当场排出来。
 * 目前支持 \o（重叠：各部分叠在一起，宽度取最宽的；\al / \ac / \ar 左 / 中 / 右对齐，缺省居中），
 * 常见的“带圈字符” eq \o\ac(○,1)、“打勾的方框” eq \o\ac(□,√) 就是它。
 * 其它开关（\f 分数、\r 根号、\s 上下标、\b 括号、\x 方框…）暂按参数文字依次排开显示。
 */
import type { EqDisplay, RunProps } from './model';

export interface EqChar {
  ch: string;
  props: RunProps;
}

/** 把一串带格式的字符按连续同格式合并成文字段 */
function segments(cs: EqChar[]): { text: string; props: RunProps }[] {
  const out: { text: string; props: RunProps }[] = [];
  for (const c of cs) {
    const last = out[out.length - 1];
    if (last && last.props === c.props) last.text += c.ch;
    else out.push({ text: c.ch, props: c.props });
  }
  return out;
}

/** 域代码（逐字符带格式）→ 显示；不是 EQ 域时返回 null */
export function parseEq(code: EqChar[]): EqDisplay | null {
  const s = code.map((c) => c.ch).join('');
  const head = /^\s*eq\s+/i.exec(s);
  if (!head) return null;
  let i = head[0].length;
  let overstrike = false;
  let align: EqDisplay['align'] = 'center';
  // 开关：\o 及其选项（\al \ac \ar）；其它开关和它们的选项跳过
  while (i < s.length && s[i] === '\\') {
    const m = /^\\([a-zA-Z]+)(\d*)/.exec(s.slice(i));
    if (!m) break;
    const name = m[1].toLowerCase();
    if (name === 'o') overstrike = true;
    else if (overstrike && (name === 'al' || name === 'ac' || name === 'ar')) align = name === 'al' ? 'left' : name === 'ar' ? 'right' : 'center';
    i += m[0].length;
    while (s[i] === ' ') i++;
  }
  // 参数：括号里按逗号 / 分号分开（括号可以嵌套）
  const open = s.indexOf('(', i);
  const close = s.lastIndexOf(')');
  const inner = open >= 0 && close > open ? code.slice(open + 1, close) : code.slice(i);
  const parts: EqChar[][] = [[]];
  let depth = 0;
  for (const c of inner) {
    if (c.ch === '(') depth++;
    else if (c.ch === ')') depth--;
    if (depth === 0 && (c.ch === ',' || c.ch === ';')) {
      parts.push([]);
      continue;
    }
    // 嵌套的开关和括号本身不显示
    if (c.ch === '(' || c.ch === ')') continue;
    parts[parts.length - 1].push(c);
  }
  const clean = parts.map((p) => {
    // 去掉嵌套开关（\s\up8 之类）
    const t = p.map((c) => c.ch).join('');
    const keep: EqChar[] = [];
    for (let k = 0; k < p.length; k++) {
      const m = t[k] === '\\' ? /^\\[a-zA-Z]+\d*/.exec(t.slice(k)) : null;
      if (m) {
        k += m[0].length - 1;
        continue;
      }
      keep.push(p[k]);
    }
    return keep;
  });
  const nonEmpty = clean.filter((p) => p.length);
  if (!nonEmpty.length) return null;
  return overstrike ? { overstrike, align, parts: nonEmpty.map(segments) } : { overstrike: false, align, parts: [segments(nonEmpty.flat())] };
}

/** 显示用的纯文字（导出、占位） */
export function eqText(e: EqDisplay): string {
  return e.parts.map((p) => p.map((x) => x.text).join('')).join(e.overstrike ? '' : '');
}

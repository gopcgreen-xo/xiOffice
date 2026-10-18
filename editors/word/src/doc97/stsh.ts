/**
 * 样式表（STSH）：每个样式的名字、基样式、段落 / 字符属性修改。[MS-DOC] 2.9.271 STSH、2.9.253 STD。
 * 样式的属性是 SPRM 列表，按基样式链从根到叶依次叠加。
 */
import { parseGrpprl, u16, type Sprm } from './sprm';

export interface StyleDef {
  istd: number;
  name: string;
  /** 内置样式编号（Normal = 0，标题 1–9 = 1–9） */
  sti: number;
  /** 1 段落、2 字符、3 表格、4 列表 */
  type: number;
  base: number;
  /** 段落样式的段落属性 / 字符属性；字符样式只有 chp */
  pap: Sprm[];
  chp: Sprm[];
}

export interface StyleSheet {
  styles: (StyleDef | null)[];
  /** 默认字体（FFN 下标）：西文、东亚、其它 */
  ftc: { ascii: number; fe: number; other: number };
}

const NIL = 0x0fff;

export function parseStyleSheet(t: Uint8Array, fc: number, lcb: number): StyleSheet {
  const out: StyleSheet = { styles: [], ftc: { ascii: 0, fe: 0, other: 0 } };
  if (!lcb) return out;
  const s = t.subarray(fc, fc + lcb);
  const cbStshi = u16(s, 0);
  const cstd = u16(s, 2);
  const cbBase = u16(s, 4);
  out.ftc = { ascii: u16(s, 14), fe: u16(s, 16), other: u16(s, 18) };
  let p = 2 + cbStshi;
  for (let i = 0; i < cstd && p + 2 <= s.length; i++) {
    const cb = u16(s, p);
    const start = p + 2;
    p = start + cb;
    if (!cb) {
      out.styles.push(null);
      continue;
    }
    const std = s.subarray(start, start + cb);
    const w1 = u16(std, 2);
    const w2 = u16(std, 4);
    const sti = u16(std, 0) & 0x0fff;
    const type = w1 & 0x0f;
    const base = (w1 >> 4) & 0x0fff;
    const cupx = w2 & 0x0f;
    // 名字：xstzName = cch(2) + UTF-16 + 结尾 0
    let o = cbBase;
    const cch = u16(std, o);
    let name = '';
    for (let k = 0; k < cch; k++) name += String.fromCharCode(u16(std, o + 2 + 2 * k));
    o += 2 + 2 * cch + 2;
    if (o & 1) o++;
    const upx: Uint8Array[] = [];
    for (let k = 0; k < cupx && o + 2 <= std.length; k++) {
      const n = u16(std, o);
      upx.push(std.subarray(o + 2, o + 2 + n));
      o += 2 + n;
      if (o & 1) o++;
    }
    let pap: Sprm[] = [];
    let chp: Sprm[] = [];
    if (type === 1) {
      // 段落样式：upx[0] = istd(2) + 段落 grpprl，upx[1] = 字符 grpprl
      if (upx[0]) pap = parseGrpprl(upx[0].subarray(2));
      if (upx[1]) chp = parseGrpprl(upx[1]);
    } else if (type === 2) {
      if (upx[0]) chp = parseGrpprl(upx[0]);
    }
    out.styles.push({ istd: i, name, sti, type, base: base === NIL ? -1 : base, pap, chp });
  }
  return out;
}

/** 从根样式到本样式的基样式链（防环） */
export function styleChain(sheet: StyleSheet, istd: number): StyleDef[] {
  const chain: StyleDef[] = [];
  const seen = new Set<number>();
  let cur: number = istd;
  while (cur >= 0 && !seen.has(cur)) {
    seen.add(cur);
    const st = sheet.styles[cur];
    if (!st) break;
    chain.unshift(st);
    cur = st.base;
  }
  return chain;
}

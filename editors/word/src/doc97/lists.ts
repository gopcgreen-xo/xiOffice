/**
 * 列表（自动编号）定义：PlfLst（列表 = 一组最多 9 级的 LVL）和 PlfLfo（列表实例，可以改起始编号或整级格式）。
 * [MS-DOC] 2.9.147 PlfLst、2.9.150 PlfLfo、2.9.149 LVL。段落用 sprmPIlfo（实例，1 起）和 sprmPIlvl（级别）引用。
 * 这里只负责读成原始结构，换成编辑器的 NumList 在 read.ts。
 */
import { i32, parseGrpprl, u16, u32, type Sprm } from './sprm';

export interface RawLevel {
  start: number;
  /** 编号格式 MSONFC：0 阿拉伯数字，1 / 2 大小写罗马，3 / 4 大小写字母，23 项目符号，255 无 … */
  nfc: number;
  /** 编号后面：0 制表符，1 空格，2 无 */
  follow: number;
  papx: Sprm[];
  chpx: Sprm[];
  /** 编号文字：值小于 9 的字符是级别占位符（0 = 第 1 级编号） */
  text: string;
}

export interface RawList {
  lsid: number;
  levels: RawLevel[];
}

export interface RawLfo {
  lsid: number;
  overrides: { ilvl: number; start: number | null; level: RawLevel | null }[];
}

/** 从 p 读一个 LVL，返回它和结束位置 */
function readLvl(t: Uint8Array, p: number): { lvl: RawLevel; end: number } {
  const cbChpx = t[p + 24] ?? 0;
  const cbPapx = t[p + 25] ?? 0;
  let o = p + 28;
  const papx = parseGrpprl(t.subarray(o, o + cbPapx));
  o += cbPapx;
  const chpx = parseGrpprl(t.subarray(o, o + cbChpx));
  o += cbChpx;
  const cch = u16(t, o);
  o += 2;
  let text = '';
  for (let k = 0; k < cch; k++) text += String.fromCharCode(u16(t, o + 2 * k));
  o += 2 * cch;
  return { lvl: { start: i32(t, p), nfc: t[p + 4] ?? 0, follow: t[p + 15] ?? 0, papx, chpx, text }, end: o };
}

export function parsePlfLst(t: Uint8Array, fc: number, lcb: number): RawList[] {
  if (lcb < 2) return [];
  // lcb 只覆盖列表头（LSTF）；各级 LVL 紧跟在后面，不在 lcb 里
  const s = t.subarray(fc);
  const n = u16(s, 0);
  const heads: { lsid: number; simple: boolean }[] = [];
  let p = 2;
  for (let i = 0; i < n; i++) {
    heads.push({ lsid: i32(s, p), simple: ((s[p + 26] ?? 0) & 1) === 1 });
    p += 28;
  }
  const out: RawList[] = [];
  for (const h of heads) {
    const levels: RawLevel[] = [];
    for (let k = 0; k < (h.simple ? 1 : 9) && p + 28 <= s.length; k++) {
      const r = readLvl(s, p);
      levels.push(r.lvl);
      p = r.end;
    }
    out.push({ lsid: h.lsid, levels });
  }
  return out;
}

export function parsePlfLfo(t: Uint8Array, fc: number, lcb: number): RawLfo[] {
  if (lcb < 4) return [];
  const s = t.subarray(fc, fc + lcb);
  const n = u32(s, 0);
  const heads: { lsid: number; count: number }[] = [];
  for (let i = 0; i < n && 4 + 16 * (i + 1) <= s.length; i++) heads.push({ lsid: i32(s, 4 + 16 * i), count: s[4 + 16 * i + 12] ?? 0 });
  let p = 4 + 16 * heads.length;
  const out: RawLfo[] = [];
  for (const h of heads) {
    p += 4; // LFOData.cp
    const overrides: RawLfo['overrides'] = [];
    for (let k = 0; k < h.count && p + 8 <= s.length; k++) {
      const start = i32(s, p);
      const flags = u32(s, p + 4);
      p += 8;
      let level: RawLevel | null = null;
      if (flags & 0x20) {
        const r = readLvl(s, p);
        level = r.lvl;
        p = r.end;
      }
      overrides.push({ ilvl: flags & 0xf, start: flags & 0x10 ? start : null, level });
    }
    out.push({ lsid: h.lsid, overrides });
  }
  return out;
}

/**
 * Word 97-2003（.doc）的属性修改（SPRM）：grpprl 是一串 [sprm(2 字节) + 操作数]，操作数大小由 sprm 的 spra 位决定。
 * 规范：[MS-DOC] 2.6 / 2.9 Sprm。这里只负责把 grpprl 拆成 { sprm, 操作数 } 列表，含义由 read.ts 解释。
 */
export interface Sprm {
  /** 完整的 sprm 值（含 sgc、spra 位），如 0x4A43 = sprmCHps */
  op: number;
  /** 操作数（不含变长的长度前缀） */
  data: Uint8Array;
}

/** sprm 属于哪一类：1 段落、2 字符、3 图片、4 节、5 表格 */
export const sgc = (op: number): number => (op >> 10) & 7;

export function u16(d: Uint8Array, o: number): number {
  return (d[o] ?? 0) | ((d[o + 1] ?? 0) << 8);
}
export function i16(d: Uint8Array, o: number): number {
  const v = u16(d, o);
  return v > 0x7fff ? v - 0x10000 : v;
}
export function u32(d: Uint8Array, o: number): number {
  return ((d[o] ?? 0) | ((d[o + 1] ?? 0) << 8) | ((d[o + 2] ?? 0) << 16) | ((d[o + 3] ?? 0) << 24)) >>> 0;
}
export function i32(d: Uint8Array, o: number): number {
  return u32(d, o) | 0;
}

/** sprmTDefTable：操作数前 2 字节是 cb，整个操作数共 cb + 1 字节（规范里的特例） */
const SPRM_TDEFTABLE = 0xd608;
/** sprmPChgTabsPapx 等 spra = 6 的变长 sprm：第一个字节是后面的字节数 */

export function parseGrpprl(buf: Uint8Array): Sprm[] {
  const out: Sprm[] = [];
  let p = 0;
  while (p + 2 <= buf.length) {
    const op = u16(buf, p);
    p += 2;
    const spra = op >> 13;
    let size: number;
    let start = p;
    switch (spra) {
      case 0:
      case 1:
        size = 1;
        break;
      case 2:
      case 4:
      case 5:
        size = 2;
        break;
      case 3:
        size = 4;
        break;
      case 7:
        size = 3;
        break;
      default:
        if (op === SPRM_TDEFTABLE) {
          size = u16(buf, p) + 1;
        } else {
          size = buf[p] ?? 0;
          start = p + 1;
        }
    }
    const end = (op === SPRM_TDEFTABLE ? p + size : start + size);
    if (end > buf.length) break;
    out.push({ op, data: buf.subarray(op === SPRM_TDEFTABLE ? p + 2 : start, end) });
    p = end;
  }
  return out;
}

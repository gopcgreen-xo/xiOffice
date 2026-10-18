/**
 * 字体文件度量（需求 W19：行高按字体文件算，而不是画布的近似值）。
 *
 * 读 TrueType / OpenType（.ttf / .otf）和字体集（.ttc）里的 name、hhea、OS/2、head 表，得到每个字体的族名（含中文名）
 * 和 Word 排版用的度量。Word（GDI）的单倍行高 = winAscent + winDescent + 外部行距，
 * 外部行距 = max(0, hhea.lineGap − ((winAscent + winDescent) − (hhea.ascender − hhea.descender)))；
 * 中日韩字体（OS/2 代码页含 GB2312 / Big5 / JIS / 韩文）不计外部行距、乘 1.3（Word 实测：宋体、仿宋、黑体、楷体 1.3 em，
 * 微软雅黑 1.32 × 1.3 = 1.716 em；Times New Roman 1.149、Calibri 1.221、Arial 1.149 em，与 Word 导出 PDF 的行距一致）。
 */

export type FaceTables = Record<string, Uint8Array>;

/** 一个字体的度量（单位 em）。ea = 中日韩字体 */
export interface FontMetrics {
  ascent: number;
  descent: number;
  lineGap: number;
  ea: boolean;
}

export interface FaceInfo {
  families: string[];
  metrics: FontMetrics;
}

const TAGS = ['name', 'hhea', 'OS/2', 'head'];

function tag(b: Uint8Array, o: number): string {
  return String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);
}

/** .ttf / .otf / .ttc 的字节 → 每个字体需要的表（只截取用到的几张表） */
export function sfntFaces(buf: Uint8Array, tags = TAGS): FaceTables[] {
  if (buf.length < 12) return [];
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const offsets: number[] = [];
  if (tag(buf, 0) === 'ttcf') {
    const n = v.getUint32(8);
    for (let i = 0; i < n && 12 + i * 4 + 4 <= buf.length; i++) offsets.push(v.getUint32(12 + i * 4));
  } else offsets.push(0);
  const out: FaceTables[] = [];
  for (const off of offsets) {
    if (off + 12 > buf.length) continue;
    const n = v.getUint16(off + 4);
    const t: FaceTables = {};
    for (let i = 0; i < n; i++) {
      const rec = off + 12 + i * 16;
      if (rec + 16 > buf.length) break;
      const name = tag(buf, rec);
      if (!tags.includes(name)) continue;
      const start = v.getUint32(rec + 8);
      const len = v.getUint32(rec + 12);
      if (start + len <= buf.length) t[name] = buf.subarray(start, start + len);
    }
    out.push(t);
  }
  return out;
}

function dv(b: Uint8Array): DataView {
  return new DataView(b.buffer, b.byteOffset, b.byteLength);
}

/** name 表里的族名（nameID 1 和 16，Windows 平台的 UTF-16，含中文名） */
function familyNames(name: Uint8Array): string[] {
  const v = dv(name);
  if (name.length < 6) return [];
  const count = v.getUint16(2);
  const strings = v.getUint16(4);
  const out = new Set<string>();
  for (let i = 0; i < count; i++) {
    const r = 6 + i * 12;
    if (r + 12 > name.length) break;
    const platform = v.getUint16(r);
    const id = v.getUint16(r + 6);
    if (platform !== 3 || (id !== 1 && id !== 16)) continue;
    const len = v.getUint16(r + 8);
    const at = strings + v.getUint16(r + 10);
    if (at + len > name.length) continue;
    let s = '';
    for (let k = 0; k + 1 < len; k += 2) s += String.fromCharCode(v.getUint16(at + k));
    if (s.trim()) out.add(s.trim());
  }
  return [...out];
}

/** 一个字体的族名和度量；缺表时 null */
export function faceInfo(t: FaceTables): FaceInfo | null {
  const { name, hhea, head } = t;
  const os2 = t['OS/2'];
  if (!name || !hhea || !head || head.length < 20 || hhea.length < 10) return null;
  const upm = dv(head).getUint16(18) || 1000;
  const h = dv(hhea);
  const hAsc = h.getInt16(4);
  const hDesc = h.getInt16(6);
  const hGap = h.getInt16(8);
  let winAsc = hAsc;
  let winDesc = -hDesc;
  let ea = false;
  if (os2 && os2.length >= 78) {
    const o = dv(os2);
    winAsc = o.getUint16(74);
    winDesc = o.getUint16(76);
    // ulCodePageRange1：17 JIS、18 GB2312、19 韩文、20 Big5、21 Johab
    if (o.getUint16(0) >= 1 && os2.length >= 82) ea = (o.getUint32(78) & (0x1f << 17)) !== 0;
  }
  const external = Math.max(0, hGap - (winAsc + winDesc - (hAsc - hDesc)));
  const families = familyNames(name);
  if (!families.length) return null;
  return { families, metrics: { ascent: winAsc / upm, descent: winDesc / upm, lineGap: external / upm, ea } };
}

/**
 * 单倍行距下的上伸和下伸（em）：西文字体外部行距算在上面；中日韩字体不计外部行距、上伸下伸各乘 1.3
 * （宋体的 hhea.lineGap 有 0.14 em，Word 里宋体的行高仍是 1.3 em）
 */
export function lineAscentDescent(m: FontMetrics): { ascent: number; descent: number } {
  if (!m.ea) return { ascent: m.ascent + m.lineGap, descent: m.descent };
  // 多出的 0.3 倍上下各一半（Word 实测：表格第一行的基线在行顶下 a + 0.15 × (a + d) em）
  const extra = (m.ascent + m.descent) * 0.15;
  return { ascent: m.ascent + extra, descent: m.descent + extra };
}

// ---------------- 已安装字体的度量表 ----------------

const table = new Map<string, FontMetrics>();
let epoch = 0;

/** 登记字体（宿主读到的系统字体）；同名的先登记的优先 */
export function registerFonts(faces: FaceInfo[]) {
  for (const f of faces) for (const n of f.families) if (!table.has(n.toLowerCase())) table.set(n.toLowerCase(), f.metrics);
  epoch++;
}

/** 字体度量变化的次数（排版缓存的键里要带上） */
export function fontEpoch(): number {
  return epoch;
}

/**
 * 按字体名查度量。没装的字体按 Word 的替换习惯找：去掉“_GB2312”等后缀（仿宋_GB2312 → 仿宋）；还找不到返回 null。
 */
export function fontMetrics(family: string): FontMetrics | null {
  const key = family.trim().toLowerCase();
  return table.get(key) ?? table.get(key.replace(/_gb2312$/, '')) ?? null;
}

/** 宿主给的字体表（表名 → base64）→ 登记度量。返回登记的字体数 */
export function registerFontTables(list: Record<string, string>[]): number {
  const faces: FaceInfo[] = [];
  for (const t of list) {
    const tables: FaceTables = {};
    for (const [k, b64] of Object.entries(t)) tables[k] = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const f = faceInfo(tables);
    if (f) faces.push(f);
  }
  if (faces.length) registerFonts(faces);
  return faces.length;
}

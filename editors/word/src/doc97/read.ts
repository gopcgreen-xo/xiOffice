/**
 * 读取 Word 97-2003 二进制文档（.doc / .dot）成编辑器的文档模型。规范：[MS-DOC]、[MS-CFB]、[MS-ODRAW]。
 *
 * 读取流程：OLE 复合文件 → WordDocument 流的 FIB → 表流（0Table / 1Table）里的 Clx（文本片段表）→ 文本；
 * 字符属性（CHPX）和段落属性（PAPX）在 WordDocument 流的 FKP 页里，按文件位置（FC）索引；样式表（STSH）提供基础格式；
 * 节属性（SEPX）给纸张、页边距、分栏；页眉页脚文字在正文之后的同一条文本流里（PlcfHdd 分段）。
 *
 * 规则（Word 的通用规则，不针对某份文档）：
 * - 字符格式 = 默认 ← 段落样式链 ← 字符样式 ← 直接格式；段落格式 = 默认 ← 样式链 ← 直接格式；
 * - 开关类属性（粗体、斜体…）取值 0 关、1 开、0x80 同样式、0x81 与样式相反；
 * - 表格：段落以 0x07 结束 = 单元格结束，带 fTtp 的 0x07 段落 = 行结束，行的列边界在 sprmTDefTable；
 * - 分页符和分节符都是 0x0C，位于节末尾的才是分节符；
 * - 页眉页脚的某个分区为空且不是第一节 = “与上一节相同”。
 *
 * 文本框、形状和浮动图片（OfficeArt 绘图）见 art.ts：读成 Run.box（组合拆成各个成员）。
 *
 * 还没有支持（读成空白或丢掉，并在 doc.lossy 里说明）：脚注 / 尾注 / 批注正文、自由曲线 / 艺术字等形状、OLE 对象、公式、
 * 嵌套表格、wmf / emf / tiff 图片、修订、加密文档、Word 95 及更早版本。保存始终是另存为 .docx。
 */
import * as CFB from 'cfb';
import {
  FIELD_CHAR,
  defaultParaProps,
  defaultRunProps,
  defaultSetup,
  newDocument,
  newHFParagraph,
  newParagraph,
  type Align,
  type Block,
  type Cell,
  type CellEdge,
  type CellEdges,
  type NumFmt,
  type NumLevel,
  type NumList,
  type Columns,
  type Doc,
  type PageGeom,
  type PageSetup,
  type Paragraph,
  type ParaProps,
  type Row,
  type Run,
  type RunProps,
  type SectionBreak,
  type SectionHF,
  type Table,
  type TabStop,
  type TextBox,
  sameProps,
} from '../core/model';
import { registerImage } from '../core/images';
import { bulletChar } from '../ooxml/docx';
import { parseOfficeArt, parsePlcSpa, parseTxbxPlc, shapeBoxes } from './art';
import { parsePlfLfo, parsePlfLst, type RawLevel } from './lists';
import { i16, i32, parseGrpprl, u16, u32, type Sprm } from './sprm';
import { parseStyleSheet, styleChain, type StyleSheet } from './stsh';

export function isDoc(bytes: Uint8Array): boolean {
  return bytes.length > 8 && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
}

const CP1252: Record<number, string> = {
  0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…', 0x86: '†', 0x87: '‡', 0x88: 'ˆ', 0x89: '‰', 0x8a: 'Š', 0x8b: '‹', 0x8c: 'Œ',
  0x8e: 'Ž', 0x91: '‘', 0x92: '’', 0x93: '“', 0x94: '”', 0x95: '•', 0x96: '–', 0x97: '—', 0x98: '˜', 0x99: '™', 0x9a: 'š', 0x9b: '›',
  0x9c: 'œ', 0x9e: 'ž', 0x9f: 'Ÿ',
};

/** 16 色表（sprmCIco / sprmCHighlight 的 ico） */
const ICO = ['', '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];

interface Piece {
  cp0: number;
  cp1: number;
  /** 第一个字符在 WordDocument 流里的字节位置 */
  fc: number;
  uni: boolean;
}

interface FkpRun {
  fc0: number;
  fc1: number;
  sprms: Sprm[];
  /** 段落：样式号 */
  istd: number;
}

/** 同一属性集的按 cp 索引：-1 = 没有直接格式 */
function indexRuns(runs: FkpRun[], pieces: Piece[], total: number): Int32Array {
  const idx = new Int32Array(total).fill(-1);
  runs.forEach((r, k) => {
    for (const p of pieces) {
      const w = p.uni ? 2 : 1;
      const a = Math.max(r.fc0, p.fc);
      const b = Math.min(r.fc1, p.fc + (p.cp1 - p.cp0) * w);
      if (a >= b) continue;
      const c0 = p.cp0 + Math.floor((a - p.fc) / w);
      const c1 = p.cp0 + Math.ceil((b - p.fc) / w);
      for (let c = c0; c < c1 && c < total; c++) idx[c] = k;
    }
  });
  return idx;
}

function parseFkps(wd: Uint8Array, t: Uint8Array, fc: number, lcb: number, kind: 'chp' | 'pap'): FkpRun[] {
  const out: FkpRun[] = [];
  if (!lcb) return out;
  const n = (lcb - 4) / 8;
  for (let i = 0; i < n; i++) {
    const pn = u32(t, fc + 4 * (n + 1) + 4 * i) & 0x3fffff;
    const page = wd.subarray(pn * 512, pn * 512 + 512);
    if (page.length < 512) continue;
    const count = page[511]!;
    for (let k = 0; k < count; k++) {
      const fc0 = u32(page, 4 * k);
      const fc1 = u32(page, 4 * (k + 1));
      if (kind === 'chp') {
        const off = (page[4 * (count + 1) + k] ?? 0) * 2;
        if (!off) {
          out.push({ fc0, fc1, sprms: [], istd: 0 });
          continue;
        }
        const cb = page[off] ?? 0;
        out.push({ fc0, fc1, sprms: parseGrpprl(page.subarray(off + 1, off + 1 + cb)), istd: 0 });
      } else {
        const off = (page[4 * (count + 1) + 13 * k] ?? 0) * 2;
        if (!off) {
          out.push({ fc0, fc1, sprms: [], istd: 0 });
          continue;
        }
        let cb = page[off] ?? 0;
        let start = off + 1;
        let size: number;
        if (cb === 0) {
          cb = page[off + 1] ?? 0;
          start = off + 2;
          size = cb * 2;
        } else size = cb * 2 - 1;
        const body = page.subarray(start, start + size);
        out.push({ fc0, fc1, sprms: parseGrpprl(body.subarray(2)), istd: u16(body, 0) });
      }
    }
  }
  return out.sort((a, b) => a.fc0 - b.fc0);
}

function parseFonts(t: Uint8Array, fc: number, lcb: number): string[] {
  const names: string[] = [];
  if (!lcb) return names;
  const count = u16(t, fc);
  let p = fc + 4;
  for (let i = 0; i < count && p < fc + lcb; i++) {
    const cb = t[p] ?? 0;
    let name = '';
    for (let o = p + 1 + 39; o + 1 < p + 1 + cb; o += 2) {
      const c = u16(t, o);
      if (!c) break;
      name += String.fromCharCode(c);
    }
    names.push(name);
    p += cb + 1;
  }
  return names;
}

/** 字符属性解析结果（带直接格式以外的特殊标记） */
interface CharInfo {
  rp: RunProps;
  spec: boolean;
  picLoc: number | null;
  symbol: { font: number; ch: number } | null;
}

const toggle = (op: number, cur: boolean): boolean => (op === 0x80 ? cur : op === 0x81 ? !cur : op !== 0);

function applyChp(sprms: Sprm[], rp: RunProps, info: Pick<CharInfo, 'spec' | 'picLoc' | 'symbol'>, fonts: string[]) {
  for (const { op, data: d } of sprms) {
    switch (op) {
      case 0x0835:
        rp.bold = toggle(d[0]!, rp.bold);
        break;
      case 0x0836:
        rp.italic = toggle(d[0]!, rp.italic);
        break;
      case 0x0837:
        rp.strike = toggle(d[0]!, !!rp.strike);
        break;
      case 0x2a3e:
        rp.underline = d[0] !== 0;
        break;
      case 0x4a43:
        rp.size = u16(d, 0) / 2;
        break;
      case 0x8840: {
        // sprmCDxaSpace：字符间距（twips，有符号）
        const v = (u16(d, 0) << 16) >> 16;
        if (v) rp.spacing = v / 20;
        else delete rp.spacing;
        break;
      }
      case 0x2a48:
        rp.vertAlign = d[0] === 1 ? 'superscript' : d[0] === 2 ? 'subscript' : undefined;
        if (!rp.vertAlign) delete rp.vertAlign;
        break;
      case 0x2a42:
        rp.color = ICO[d[0]!] || 'auto';
        break;
      case 0x6870:
        rp.color = d[3] ? 'auto' : [d[0], d[1], d[2]].map((v) => (v ?? 0).toString(16).padStart(2, '0')).join('').toUpperCase();
        break;
      case 0x2a0c:
        if (d[0]) rp.highlight = ICO[d[0]];
        else delete rp.highlight;
        break;
      case 0x4a4f:
        if (fonts[u16(d, 0)]) rp.fontLatin = fonts[u16(d, 0)]!;
        break;
      case 0x4a50:
        if (fonts[u16(d, 0)]) rp.fontEA = fonts[u16(d, 0)]!;
        break;
      case 0x0855:
        info.spec = d[0] === 1;
        break;
      case 0x6a03:
        info.picLoc = i32(d, 0);
        break;
      case 0x6a09:
        info.symbol = { font: u16(d, 0), ch: u16(d, 2) };
        break;
      default:
    }
  }
}

interface ParaInfo {
  props: ParaProps;
  istd: number;
  inTable: boolean;
  ttp: boolean;
  itap: number;
  /** 列表实例（sprmPIlfo，1 起；0 = 没有）和级别 */
  ilfo: number;
  ilvl: number;
  tdef: Sprm | null;
  rowHeight: { v: number; rule: Row['heightRule'] } | null;
  header: boolean;
  tableSprms: Sprm[];
}

/** MSONFC（编号格式）→ 编辑器的编号格式；没有对应的按阿拉伯数字 */
const NUM_FMT: Record<number, NumFmt> = {
  0: 'decimal',
  1: 'upperRoman',
  2: 'lowerRoman',
  3: 'upperLetter',
  4: 'lowerLetter',
  10: 'chineseCounting',
  11: 'chineseCounting',
  18: 'decimalEnclosedCircle',
  22: 'decimal',
  23: 'bullet',
  30: 'ideographTraditional',
  33: 'chineseCounting',
  34: 'chineseCounting',
  35: 'chineseCounting',
  36: 'chineseCounting',
  37: 'chineseCounting',
  38: 'chineseCounting',
  39: 'chineseCounting',
  255: 'none',
};

const JC: Align[] = ['left', 'center', 'right', 'both', 'distribute', 'both'];

/** 图文框（[MS-DOC] sprmPPc / sprmPDxaAbs / sprmPDyaAbs / sprmPWr / sprmPDxaWidth）：出现任何一个就是图文框段落 */
const FRAME_X_ALIGN: Record<number, string> = { [-4]: 'center', [-8]: 'right', [-12]: 'inside', [-16]: 'outside' };
function frameOf(p: ParaInfo['props']): NonNullable<ParaInfo['props']['frame']> {
  return (p.frame ??= { hAnchor: 'text', vAnchor: 'margin', x: 0, y: 0 });
}

/** 一段段落属性 grpprl 叠加到缺省段落属性上的结果（测试用：Word 新版本存不出只有旧版 sprm 的文件） */
export function papOf(grpprl: Uint8Array): ParaProps {
  const info: ParaInfo = { props: { ...defaultParaProps(), align: 'left' }, istd: 0, inTable: false, ttp: false, itap: 1, ilfo: 0, ilvl: 0, tdef: null, rowHeight: null, header: false, tableSprms: [] };
  applyPap(parseGrpprl(grpprl), info);
  return info.props;
}

function applyPap(sprms: Sprm[], info: ParaInfo) {
  const p = info.props;
  for (const s of sprms) {
    const { op, data: d } = s;
    switch (op) {
      case 0x261b: {
        // PositionCodeOperand：第 4–5 位 pcVert（0 页边距、1 页面、2 段落），第 6–7 位 pcHorz（0 栏、1 页边距、2 页面）
        const f = frameOf(p);
        f.vAnchor = ['margin', 'page', 'text'][(d[0]! >> 4) & 3] ?? 'margin';
        f.hAnchor = ['text', 'margin', 'page'][(d[0]! >> 6) & 3] ?? 'text';
        break;
      }
      case 0x8418: {
        const v = (u16(d, 0) << 16) >> 16;
        const f = frameOf(p);
        if (FRAME_X_ALIGN[v]) f.xAlign = FRAME_X_ALIGN[v];
        else {
          delete f.xAlign;
          f.x = v / 20;
        }
        break;
      }
      case 0x8419: {
        const v = (u16(d, 0) << 16) >> 16;
        // 负数是对齐方式（顶端、居中…），本编辑器按 0 处理
        frameOf(p).y = v > 0 ? v / 20 : 0;
        break;
      }
      case 0x2423:
      case 0x841a:
        frameOf(p);
        break;
      // sprmPDylBefore / sprmPDylAfter：以 1/100 行为单位的段前 / 段后（同 docx 的 beforeLines / afterLines，0 = 按磅值）
      case 0x4458: {
        const v = (u16(d, 0) << 16) >> 16;
        p.spaceBeforeLines = v ? v / 100 : null;
        break;
      }
      case 0x4459: {
        const v = (u16(d, 0) << 16) >> 16;
        p.spaceAfterLines = v ? v / 100 : null;
        break;
      }
      // sprmPDxcLeft1：首行缩进字符数（1/100 字符，负数为悬挂；0 = 按磅值）
      case 0x4457: {
        const v = (u16(d, 0) << 16) >> 16;
        p.firstLineChars = v ? v / 100 : null;
        break;
      }
      // sprmPFDyaBeforeAuto / sprmPFDyaAfterAuto：段前 / 段后“自动”
      case 0x245b:
        if (d[0]) p.spaceBeforeAuto = true;
        else delete p.spaceBeforeAuto;
        break;
      case 0x245c:
        if (d[0]) p.spaceAfterAuto = true;
        else delete p.spaceAfterAuto;
        break;
      case 0x2403:
      case 0x2461:
        p.align = JC[d[0]!] ?? 'left';
        break;
      // [MS-DOC]：0x840E = sprmPDxaRight80，0x840F = sprmPDxaLeft80（旧版，只有它们的 .doc 靠这两个）；0x845D / 0x845E 是新版
      case 0x840e:
      case 0x845d:
        p.indentRight = i16(d, 0) / 20;
        break;
      case 0x840f:
      case 0x845e:
        p.indentLeft = i16(d, 0) / 20;
        break;
      case 0x8411:
      case 0x8460:
        p.firstLine = i16(d, 0) / 20;
        // 同时写了字符数（sprmPDxcLeft1）时字符数优先（同 docx 的 firstLineChars）
        if (!sprms.some((x) => x.op === 0x4457)) p.firstLineChars = null;
        break;
      case 0xa413:
        p.spaceBefore = u16(d, 0) / 20;
        break;
      case 0xa414:
        p.spaceAfter = u16(d, 0) / 20;
        break;
      case 0x6412: {
        const line = i16(d, 0);
        const mult = u16(d, 2);
        p.line = mult ? { rule: 'auto', value: Math.round((line / 240) * 100) / 100 || 1 } : line > 0 ? { rule: 'atLeast', value: line / 20 } : line < 0 ? { rule: 'exact', value: -line / 20 } : { rule: 'auto', value: 1 };
        break;
      }
      case 0x2405:
        p.keepLines = d[0] === 1;
        break;
      case 0x2406:
        p.keepNext = d[0] === 1;
        break;
      case 0x2407:
        p.pageBreakBefore = d[0] === 1;
        break;
      case 0x2431:
        p.widowControl = d[0] === 1;
        break;
      case 0x2640:
        if (d[0]! < 9) p.outline = d[0];
        else delete p.outline;
        break;
      case 0x2416:
        info.inTable = d[0] === 1;
        break;
      case 0x2417:
        info.ttp = d[0] === 1;
        break;
      case 0x6649:
        info.itap = u32(d, 0);
        break;
      case 0x460b:
        // 0x0001–0x07FE 是 PlfLfo 里的下标；更大的值是旧式编号 / 特殊值，不支持
        info.ilfo = u16(d, 0) <= 0x7fe ? u16(d, 0) : 0;
        break;
      case 0x260a:
        info.ilvl = Math.min(8, d[0] ?? 0);
        break;
      case 0xc60d:
      case 0xc615:
        applyTabs(d, op === 0xc615, p);
        break;
      case 0xd608:
        info.tdef = s;
        break;
      case 0x9407: {
        const v = i16(d, 0);
        info.rowHeight = v > 0 ? { v: v / 20, rule: 'atLeast' } : v < 0 ? { v: -v / 20, rule: 'exact' } : null;
        break;
      }
      case 0x3404:
        info.header = d[0] === 1;
        break;
      default:
        if (((op >> 10) & 7) === 5) info.tableSprms.push(s);
    }
  }
}

function applyTabs(d: Uint8Array, withClose: boolean, p: ParaProps) {
  let o = 0;
  const nDel = d[o++] ?? 0;
  const del: number[] = [];
  for (let i = 0; i < nDel; i++) del.push(i16(d, o + 2 * i) / 20);
  o += 2 * nDel + (withClose ? 2 * nDel : 0);
  const nAdd = d[o++] ?? 0;
  const tabs: TabStop[] = (p.tabs ?? []).filter((t) => !del.some((x) => Math.abs(x - t.pos) < 0.5));
  for (let i = 0; i < nAdd; i++) {
    const pos = i16(d, o + 2 * i) / 20;
    const tbd = d[o + 2 * nAdd + i] ?? 0;
    const jc = tbd & 7;
    if (jc > 2) continue; // 小数点对齐、竖线：暂不支持
    const tlc = (tbd >> 3) & 7;
    const leader: TabStop['leader'] = tlc === 1 || tlc === 5 ? 'dot' : tlc === 2 ? 'hyphen' : tlc === 3 ? 'underscore' : 'none';
    tabs.push({ pos, align: jc === 1 ? 'center' : jc === 2 ? 'right' : 'left', leader });
  }
  tabs.sort((a, b) => a.pos - b.pos);
  if (tabs.length) p.tabs = tabs;
  else delete p.tabs;
}

// ---------------- 图片 ----------------

function readArtRecords(d: Uint8Array, from: number, to: number, out: { mime: string; data: Uint8Array }[]) {
  let p = from;
  while (p + 8 <= to) {
    const verInst = u16(d, p);
    const type = u16(d, p + 2);
    const len = u32(d, p + 4);
    const body = p + 8;
    const end = Math.min(to, body + len);
    if ((verInst & 0xf) === 0xf) readArtRecords(d, body, end, out);
    else if (type === 0xf007) {
      // BSE：36 字节头 + 名字 + 内嵌的 blip
      const cbName = d[body + 33] ?? 0;
      readArtRecords(d, body + 36 + cbName, end, out);
    } else if (type === 0xf01d || type === 0xf01e) {
      const inst = verInst >> 4;
      const twoUid = type === 0xf01d ? inst === 0x46b || inst === 0x6e3 : inst === 0x6e1;
      const start = body + 16 + (twoUid ? 16 : 0) + 1;
      if (start < end) out.push({ mime: type === 0xf01d ? 'image/jpeg' : 'image/png', data: d.slice(start, end) });
    }
    p = end;
  }
}

function readPicture(data: Uint8Array | null, loc: number): { id: string; w: number; h: number } | null {
  if (!data || loc < 0 || loc + 68 > data.length) return null;
  const lcb = i32(data, loc);
  const cbHeader = u16(data, loc + 4);
  const mm = i16(data, loc + 6);
  if (mm === 0x66) return null; // 外部文件链接的图片（MM_SHAPEFILE）；0x64 = MM_SHAPE 是内嵌的 OfficeArt 图
  const dxaGoal = i16(data, loc + 28);
  const dyaGoal = i16(data, loc + 30);
  const mx = u16(data, loc + 32) || 1000;
  const my = u16(data, loc + 34) || 1000;
  const found: { mime: string; data: Uint8Array }[] = [];
  readArtRecords(data, loc + cbHeader, Math.min(data.length, loc + lcb), found);
  const img = found[0];
  if (!img || dxaGoal <= 0 || dyaGoal <= 0) return null;
  return { id: registerImage(img.data, img.mime), w: Math.round(((dxaGoal / 20) * mx) / 1000 * 10) / 10, h: Math.round(((dyaGoal / 20) * my) / 1000 * 10) / 10 };
}

// ---------------- 表格边框 ----------------

/** 线宽 dpt 是 1/8 磅；类型 0 = 无，255 = 未设置；ico 0 = 自动（黑） */
function edgeOf(width: number, type: number, color: string): CellEdge | null {
  if (type === 0 || type === 0xff) return null;
  return { w: Math.max(0.25, width / 8), color };
}

/** Brc80（4 字节）：线宽、类型、颜色序号、间距；全 1 = 未设置 */
function brc80(d: Uint8Array, o: number): CellEdge | null {
  if (u32(d, o) === 0xffffffff) return null;
  return edgeOf(d[o] ?? 0, d[o + 1] ?? 0, ICO[d[o + 2] ?? 0] || '000000');
}

/** Brc（8 字节）：颜色 RGB + 自动标志，线宽，类型，间距 */
function brc(d: Uint8Array, o: number): CellEdge | null {
  const auto = (d[o + 3] ?? 0) === 0xff;
  const color = auto ? '000000' : [d[o], d[o + 1], d[o + 2]].map((v) => (v ?? 0).toString(16).padStart(2, '0')).join('').toUpperCase();
  return edgeOf(d[o + 4] ?? 0, d[o + 5] ?? 0, color);
}

/** Shd80（2 字节）：背景色序号在第 5–9 位，前景色在 0–4 位，图案在 10–15 位（1 = 实心，用前景色） */
function ipatIco(v: number): string {
  const fore = v & 0x1f;
  const back = (v >> 5) & 0x1f;
  const ipat = v >> 10;
  return (ipat === 1 ? ICO[fore] : ICO[back]) ?? '';
}

// ---------------- 主流程 ----------------

interface Sect {
  /** 节末尾字符的 cp（分节符 / 文末） */
  end: number;
  setup: PageSetup;
  bkc: number;
  cols?: Columns;
}

export function readDoc(bytes: Uint8Array): Doc {
  let cfb: CFB.CFB$Container;
  try {
    cfb = CFB.parse(bytes);
  } catch {
    throw new Error('不是有效的 Word 文档（无法读取文件结构）');
  }
  const entry = (name: string): Uint8Array | undefined => {
    const c = CFB.find(cfb, name)?.content as ArrayLike<number> | ArrayBufferView | undefined;
    if (!c) return undefined;
    // cfb 在不同环境下给 Array / Buffer / Uint8Array，统一成独立的 Uint8Array
    return ArrayBuffer.isView(c) ? new Uint8Array(c.buffer, c.byteOffset, c.byteLength) : Uint8Array.from(c as ArrayLike<number>);
  };
  const wd = entry('WordDocument');
  if (!wd || wd.length < 160) throw new Error('不是有效的 Word 97-2003 文档（缺少 WordDocument）');
  if (u16(wd, 0) !== 0xa5ec) throw new Error('不是有效的 Word 97-2003 文档');
  const nFib = u16(wd, 2);
  if (nFib < 0xc1) throw new Error('不支持 Word 95 及更早版本的 .doc 文件，请先在 Word 里另存为 .docx');
  const flags = u16(wd, 10);
  if (flags & 0x0100) throw new Error('文档已加密，请先在 Word 里去掉密码后再打开');
  const t = entry(flags & 0x0200 ? '1Table' : '0Table');
  if (!t) throw new Error('文档损坏（缺少表流）');
  const data = entry('Data') ?? null;

  const csw = u16(wd, 32);
  const lwBase = 34 + csw * 2 + 2;
  const ccp = (i: number) => i32(wd, lwBase + 12 + 4 * i);
  const ccpText = ccp(0);
  const ccpFtn = ccp(1);
  const ccpHdd = ccp(2);
  const ccpAtn = ccp(4);
  const ccpEdn = ccp(5);
  const ccpTxbx = ccp(6);
  const fcBase = lwBase + 88 + 2;
  const pair = (i: number): [number, number] => [u32(wd, fcBase + 8 * i), u32(wd, fcBase + 8 * i + 4)];

  // ---- 文本片段表 ----
  const [fcClx, lcbClx] = pair(33);
  const clx = t.subarray(fcClx, fcClx + lcbClx);
  let cp = 0;
  while (cp < clx.length && clx[cp] === 1) cp += 3 + u16(clx, cp + 1);
  if (clx[cp] !== 2) throw new Error('文档损坏（找不到文本片段表）');
  const lcbPcd = u32(clx, cp + 1);
  const plc = clx.subarray(cp + 5, cp + 5 + lcbPcd);
  const nPieces = (lcbPcd - 4) / 12;
  const pieces: Piece[] = [];
  for (let i = 0; i < nPieces; i++) {
    const raw = u32(plc, 4 * (nPieces + 1) + 8 * i + 2);
    const compressed = (raw & 0x40000000) !== 0;
    const fcv = raw & 0x3fffffff;
    pieces.push({ cp0: u32(plc, 4 * i), cp1: u32(plc, 4 * (i + 1)), fc: compressed ? fcv / 2 : fcv, uni: !compressed });
  }
  const total = pieces.length ? pieces[pieces.length - 1]!.cp1 : 0;
  const chars: string[] = new Array(total).fill('');
  for (const p of pieces) {
    for (let c = p.cp0; c < p.cp1; c++) {
      if (p.uni) chars[c] = String.fromCharCode(u16(wd, p.fc + (c - p.cp0) * 2));
      else {
        const b = wd[p.fc + (c - p.cp0)] ?? 0;
        chars[c] = CP1252[b] ?? String.fromCharCode(b);
      }
    }
  }

  // ---- 格式 ----
  const fonts = parseFonts(t, ...pair(15));
  const sheet: StyleSheet = parseStyleSheet(t, ...pair(1));
  const chpRuns = parseFkps(wd, t, ...pair(12), 'chp');
  const papRuns = parseFkps(wd, t, ...pair(13), 'pap');
  // 段落属性太大放不进 FKP 页时（多列表格的表格定义常见）：sprmPHugePapx 指向 Data 流里的 PrcData（长度 + grpprl），并入原来的属性列表
  for (const r of papRuns) {
    const huge = r.sprms.find((x) => x.op === 0x6646);
    if (!huge || !data) continue;
    const at = u32(huge.data, 0);
    if (at + 2 > data.length) continue;
    const extra = parseGrpprl(data.subarray(at + 2, at + 2 + u16(data, at)));
    r.sprms = r.sprms.filter((x) => x !== huge).concat(extra);
  }
  const chpIdx = indexRuns(chpRuns, pieces, total);
  const papIdx = indexRuns(papRuns, pieces, total);

  const baseRun = (): RunProps => {
    const rp = defaultRunProps();
    rp.size = 10;
    const a = fonts[sheet.ftc.ascii];
    const fe = fonts[sheet.ftc.fe];
    if (a) rp.fontLatin = a;
    if (fe) rp.fontEA = fe;
    return rp;
  };

  const styleRun = (istd: number): CharInfo => {
    const info: CharInfo = { rp: baseRun(), spec: false, picLoc: null, symbol: null };
    for (const st of styleChain(sheet, istd)) applyChp(st.chp, info.rp, info, fonts);
    return info;
  };
  const styleCache = new Map<number, CharInfo>();
  const styleCharInfo = (istd: number) => {
    let v = styleCache.get(istd);
    if (!v) styleCache.set(istd, (v = styleRun(istd)));
    return v;
  };

  const charCache = new Map<string, CharInfo>();
  const charInfo = (cpos: number, paraIstd: number): CharInfo => {
    const k = chpIdx[cpos]!;
    const key = `${paraIstd}:${k}`;
    let v = charCache.get(key);
    if (v) return v;
    const base = styleCharInfo(paraIstd);
    v = { rp: { ...base.rp }, spec: false, picLoc: null, symbol: null };
    if (k >= 0) {
      const sprms = chpRuns[k]!.sprms;
      // 字符样式（sprmCIstd）先于直接格式
      const cs = sprms.find((s) => s.op === 0x4a30);
      if (cs) for (const st of styleChain(sheet, u16(cs.data, 0))) applyChp(st.chp, v.rp, v, fonts);
      applyChp(sprms, v.rp, v, fonts);
    }
    charCache.set(key, v);
    return v;
  };

  const lossy = new Set<string>();
  const paraInfo = (cpos: number): ParaInfo => {
    const k = papIdx[cpos]!;
    const run = k >= 0 ? papRuns[k]! : null;
    const istd = run ? run.istd : 0;
    const info: ParaInfo = { props: { ...defaultParaProps(), align: 'left' }, istd, inTable: false, ttp: false, itap: 1, ilfo: 0, ilvl: 0, tdef: null, rowHeight: null, header: false, tableSprms: [] };
    for (const st of styleChain(sheet, istd)) applyPap(st.pap, info);
    // 标题样式：编辑器按样式名认标题
    const st = sheet.styles[istd];
    const m = st && (/^(?:heading|标题)\s*([1-9])$/i.exec(st.name) ?? (st.sti >= 1 && st.sti <= 9 ? [st.name, String(st.sti)] : null));
    if (m) {
      info.props.style = `Heading${m[1]}`;
      info.props.outline = Number(m[1]) - 1;
    }
    // 列表级别自己的缩进在样式之后、段落直接格式之前生效（和 Word 一样，段落自己写了缩进就以段落为准）
    const direct = run?.sprms ?? [];
    const lastOf = (op: number) => direct.filter((x) => x.op === op).pop();
    const dl = lastOf(0x460b);
    const dv = lastOf(0x260a);
    const ilfo = dl ? (u16(dl.data, 0) <= 0x7fe ? u16(dl.data, 0) : 0) : info.ilfo;
    const ilvl = dv ? Math.min(8, dv.data[0] ?? 0) : info.ilvl;
    const level = ilfo ? lists[ilfo - 1]?.levels[ilvl] : undefined;
    const ind = level && levelInd.get(level);
    if (ind) {
      if (ind.left !== undefined) info.props.indentLeft = ind.left;
      if (ind.first !== undefined) {
        info.props.firstLine = ind.first;
        info.props.firstLineChars = null;
      }
    }
    if (run) applyPap(direct, info);
    if (ilfo && level) info.props.num = { id: ilfo, level: ilvl };
    else if (ilfo) lossy.add('自动编号（找不到列表定义）');
    info.props.snapToGrid = true;
    return info;
  };

  // ---- 列表（自动编号）----
  const fmtOf = (nfc: number): NumFmt => NUM_FMT[nfc] ?? 'decimal';
  const levelInd = new Map<NumLevel, { left?: number; first?: number }>();
  const toLevel = (r: RawLevel): NumLevel => {
    const last = (ops: number[]) => r.papx.filter((x) => ops.includes(x.op)).pop();
    const left = last([0x840f, 0x845e]);
    const first = last([0x8411, 0x8460]);
    const fmt = fmtOf(r.nfc);
    let text = '';
    for (const ch of r.text) text += ch.charCodeAt(0) < 9 ? `%${ch.charCodeAt(0) + 1}` : ch;
    // 编号自己的字符格式：和正文默认格式不同的项
    const base = baseRun();
    const rp = { ...base };
    applyChp(r.chpx, rp, { spec: false, picLoc: null, symbol: null }, fonts);
    const label: Partial<RunProps> = {};
    if (rp.bold !== base.bold) label.bold = rp.bold;
    if (rp.italic !== base.italic) label.italic = rp.italic;
    if (rp.size !== base.size) label.size = rp.size;
    if (rp.color !== base.color) label.color = rp.color;
    // Symbol / Wingdings 的项目符号已换成通用字符，不再带这两个字体
    if (fmt !== 'bullet') {
      if (rp.fontLatin !== base.fontLatin) label.fontLatin = rp.fontLatin;
      if (rp.fontEA !== base.fontEA) label.fontEA = rp.fontEA;
    }
    const level: NumLevel = {
      fmt,
      text: fmt === 'bullet' ? bulletChar(text) : text,
      start: r.start,
      indentLeft: left ? i16(left.data, 0) / 20 : 0,
      firstLine: first ? i16(first.data, 0) / 20 : 0,
      firstLineChars: null,
      suffix: r.follow === 1 ? 'space' : r.follow === 2 ? 'nothing' : 'tab',
      ...(Object.keys(label).length ? { label } : {}),
    };
    if (left || first) levelInd.set(level, { left: left ? level.indentLeft : undefined, first: first ? level.firstLine : undefined });
    return level;
  };
  const fill = (levels: NumLevel[]): NumLevel[] => {
    for (let i = 0; i < 9; i++) levels[i] ??= { fmt: 'decimal', text: `%${i + 1}.`, start: 1, indentLeft: 21 * (i + 1), firstLine: -21, firstLineChars: null, suffix: 'tab' };
    return levels;
  };
  const rawLists = parsePlfLst(t, ...pair(73));
  const lists: NumList[] = [];
  parsePlfLfo(t, ...pair(74)).forEach((lfo, i) => {
    const base = rawLists.find((l) => l.lsid === lfo.lsid);
    if (!base) {
      lists.push({ id: i + 1, preset: `doc-${i + 1}`, levels: fill([]) });
      return;
    }
    const levels = fill(base.levels.map(toLevel));
    for (const o of lfo.overrides) {
      if (o.level) levels[o.ilvl] = toLevel(o.level);
      if (o.start !== null && levels[o.ilvl]) levels[o.ilvl] = { ...levels[o.ilvl]!, start: o.start };
    }
    lists.push({ id: i + 1, preset: `doc-${i + 1}`, levels });
  });

  // ---- 节 ----
  const [fcSed, lcbSed] = pair(6);
  const nSect = lcbSed ? (lcbSed - 4) / 16 : 0;
  const sects: Sect[] = [];
  for (let i = 0; i < nSect; i++) {
    const end = u32(t, fcSed + 4 * (i + 1));
    const fcSepx = i32(t, fcSed + 4 * (nSect + 1) + 12 * i + 2);
    const setup = defaultSetup();
    let bkc = 2;
    let cols: Columns | undefined;
    let grid = { clm: 0, pitch: 0, charSpace: 0 };
    if (fcSepx >= 0 && fcSepx + 2 <= wd.length) {
      const cb = u16(wd, fcSepx);
      let nc = 1;
      let space = 36;
      let sep = false;
      for (const { op, data: d } of parseGrpprl(wd.subarray(fcSepx + 2, fcSepx + 2 + cb))) {
        switch (op) {
          case 0xb01f:
            setup.width = u16(d, 0) / 20;
            break;
          case 0xb020:
            setup.height = u16(d, 0) / 20;
            break;
          case 0xb021:
            setup.marginLeft = u16(d, 0) / 20;
            break;
          case 0xb022:
            setup.marginRight = u16(d, 0) / 20;
            break;
          case 0x9023:
            setup.marginTop = i16(d, 0) / 20;
            break;
          case 0x9024:
            setup.marginBottom = i16(d, 0) / 20;
            break;
          case 0xb017:
            setup.headerDistance = u16(d, 0) / 20;
            break;
          case 0xb018:
            setup.footerDistance = u16(d, 0) / 20;
            break;
          case 0x300a:
            setup.titlePg = d[0] === 1;
            break;
          case 0x3009:
            bkc = d[0]!;
            break;
          case 0x500b:
            nc = u16(d, 0) + 1;
            break;
          case 0x900c:
            space = u16(d, 0) / 20;
            break;
          case 0x3019:
            sep = d[0] === 1;
            break;
          case 0x5032:
            grid.clm = u16(d, 0);
            break;
          case 0x9031:
            grid.pitch = u16(d, 0) / 20;
            break;
          case 0x7030:
            grid.charSpace = i32(d, 0);
            break;
          default:
        }
      }
      if (nc > 1) cols = { num: nc, space, sep };
      if (grid.clm) {
        setup.grid = {
          type: grid.clm === 1 ? 'linesAndChars' : grid.clm === 3 ? 'snapToChars' : 'lines',
          linePitch: grid.pitch || setup.grid.linePitch,
          charSpace: grid.clm === 1 || grid.clm === 3 ? grid.charSpace : 0,
        };
      } else setup.grid = { type: 'none', linePitch: grid.pitch || 18, charSpace: 0 };
    }
    sects.push({ end, setup, bkc, cols });
  }
  if (!sects.length) sects.push({ end: ccpText, setup: defaultSetup(), bkc: 2 });
  const sectEnds = new Map<number, number>();
  sects.slice(0, -1).forEach((s, i) => sectEnds.set(s.end - 1, i));

  // ---- 绘图对象（文本框、形状、浮动图片）----
  const hddStart0 = ccpText + ccpFtn;
  const txbxBase = ccpText + ccpFtn + ccpHdd + ccp(3) + ccpAtn + ccpEdn;
  const [fcDgg, lcbDgg] = pair(50);
  const art = parseOfficeArt(lcbDgg && fcDgg + lcbDgg <= t.length ? t.subarray(fcDgg, fcDgg + lcbDgg) : new Uint8Array(0));
  const spaMain = parsePlcSpa(t, ...pair(40));
  const spaHdr = parsePlcSpa(t, ...pair(41));
  const txbxMain = parseTxbxPlc(t, ...pair(56));
  const txbxHdr = parseTxbxPlc(t, ...pair(58));
  const blipIds = new Map<number, string | null>();
  const blipImage = (pib: number): string | null => {
    if (blipIds.has(pib)) return blipIds.get(pib)!;
    const b = art.blips[pib - 1];
    const bytes = b ? (b.inline ?? (b.foDelay + b.size <= wd.length ? wd.subarray(b.foDelay, b.foDelay + b.size) : null)) : null;
    const found: { mime: string; data: Uint8Array }[] = [];
    if (bytes) readArtRecords(bytes, 0, bytes.length, found);
    const id = found[0] ? registerImage(found[0].data, found[0].mime) : null;
    blipIds.set(pib, id);
    return id;
  };
  /** 锚点字符（0x08）所在 cp 的形状 → 文本框模型（组合拆开）；找不到绘图数据返回 null */
  const boxesAt = (c: number): TextBox[] | null => {
    const hdr = c >= hddStart0;
    const fspa = hdr ? spaHdr.get(c - hddStart0) : spaMain.get(c);
    const sp = fspa && (hdr ? art.header : art.main).get(fspa.spid);
    if (!fspa || !sp) return null;
    return shapeBoxes(sp, fspa, {
      image: blipImage,
      text: (spid) => {
        const range = (hdr ? txbxHdr : txbxMain).get(spid);
        if (!range) return null;
        const base = hdr ? txbxBase + ccpTxbx : txbxBase;
        // 每个文本框故事末尾多一个段落标记（故事的结束符），不是内容
        const to = Math.min(total, base + range[1]);
        const blocks = readStory(base + range[0], chars[to - 1] === '\r' ? to - 1 : to, false);
        if (blocks.some((x) => x.kind === 'table')) lossy.add('文本框中的表格（已按普通段落显示）');
        const paras = blocks.flatMap((x) => (x.kind === 'p' ? [x] : x.rows.flatMap((r) => r.cells.flatMap((cl) => cl.paragraphs))));
        return paras.length ? paras : null;
      },
      emptyPara: () => newParagraph(),
      lossy: (what) => lossy.add(what),
    });
  };

  // ---- 段落 / 表格 ----
  const unsupported = { footnote: false, shape: false, note: false };
  const buildPara = (c0: number, c1: number, info: ParaInfo, breakAt: number | null): Paragraph => {
    // c1 为终止符的 cp；文字在 [c0, c1)
    const runs: Run[] = [];
    const stack: { phase: 'code' | 'result'; code: string; skip: boolean }[] = [];
    const push = (text: string, rp: RunProps, extra: Partial<Run> = {}) => {
      const last = runs[runs.length - 1];
      if (!extra.field && !extra.image && !extra.box && last && !last.field && !last.image && !last.box && sameProps(last.props, rp)) last.text += text;
      else runs.push({ text, props: { ...rp }, ...extra });
    };
    for (let c = c0; c < c1; c++) {
      const ch = chars[c]!;
      if (ch === '\x13') {
        stack.push({ phase: 'code', code: '', skip: false });
        continue;
      }
      if (ch === '\x14' && stack.length) {
        const top = stack[stack.length - 1]!;
        top.phase = 'result';
        const word = /^\s*([A-Za-z]+)/.exec(top.code)?.[1]?.toUpperCase();
        if (word === 'PAGE' || word === 'NUMPAGES') {
          top.skip = true;
          push(FIELD_CHAR, charInfo(c, info.istd).rp, { field: word });
        }
        continue;
      }
      if (ch === '\x15' && stack.length) {
        stack.pop();
        continue;
      }
      if (stack.length) {
        const top = stack[stack.length - 1]!;
        if (top.phase === 'code') {
          top.code += ch;
          continue;
        }
        if (stack.some((s) => s.skip)) continue;
      }
      const ci = charInfo(c, info.istd);
      if (ci.spec) {
        if (ch === '\x01') {
          const pic = ci.picLoc !== null ? readPicture(data, ci.picLoc) : null;
          if (pic) push(FIELD_CHAR, ci.rp, { image: { id: pic.id, w: pic.w, h: pic.h } });
          else lossy.add('图片、公式或对象（格式暂不支持）');
          continue;
        }
        if (ch === '\x08') {
          const boxes = boxesAt(c);
          if (boxes) for (const box of boxes) push(FIELD_CHAR, ci.rp, { box });
          else unsupported.shape = true;
          continue;
        }
        if (ch === '\x02') {
          unsupported.footnote = true;
          continue;
        }
        if (ch === '\x05') {
          unsupported.note = true;
          continue;
        }
      }
      if (ci.symbol && ch === '\x28') {
        const f = fonts[ci.symbol.font] ?? 'Symbol';
        push(String.fromCharCode(ci.symbol.ch), { ...ci.rp, fontLatin: f, fontEA: f });
        continue;
      }
      switch (ch) {
        case '\x0b':
          push('\n', ci.rp);
          break;
        case '\x0c':
          push('\f', ci.rp);
          break;
        case '\x0e':
          push('\v', ci.rp);
          break;
        case '\x1e':
          push('-', ci.rp);
          break;
        case '\x1f':
          break;
        default:
          if (ch >= ' ' || ch === '\t') push(ch, ci.rp);
      }
    }
    const mark = { ...charInfo(Math.min(c1, total - 1), info.istd).rp };
    const para: Paragraph = { kind: 'p', props: info.props, runs, mark };
    void breakAt;
    return para;
  };

  interface Raw {
    para: Paragraph;
    info: ParaInfo;
    end: string;
    sect: number | null;
  }

  const readStory = (c0: number, c1: number, allowSections: boolean): Block[] => {
    const raws: Raw[] = [];
    let start = c0;
    for (let c = c0; c < c1; c++) {
      const ch = chars[c]!;
      const isSect = allowSections && ch === '\x0c' && sectEnds.has(c);
      if (ch !== '\r' && ch !== '\x07' && !isSect) continue;
      const info = paraInfo(c);
      const para = buildPara(start, c, info, null);
      raws.push({ para, info, end: ch, sect: isSect ? sectEnds.get(c)! : null });
      start = c + 1;
    }
    if (start < c1) {
      // 没有结束符的尾巴（损坏的文件）
      const info = paraInfo(Math.min(c1 - 1, total - 1));
      raws.push({ para: buildPara(start, c1, info, null), info, end: '\r', sect: null });
    }
    const blocks: Block[] = [];
    for (let i = 0; i < raws.length; i++) {
      const r = raws[i]!;
      if (r.info.inTable && r.info.itap <= 1) {
        // 取出到行结束为止的所有行
        const rows: { cells: Paragraph[][]; tdef: Sprm | null; info: ParaInfo }[] = [];
        let cur: Paragraph[][] = [];
        let cell: Paragraph[] = [];
        let j = i;
        let lastInfo = r.info;
        for (; j < raws.length && raws[j]!.info.inTable && raws[j]!.info.itap <= 1; j++) {
          const x = raws[j]!;
          if (x.info.ttp) {
            rows.push({ cells: cur, tdef: x.info.tdef ?? lastInfo.tdef, info: x.info });
            cur = [];
            cell = [];
            lastInfo = x.info;
            continue;
          }
          if (x.info.tdef) lastInfo = x.info;
          cell.push(x.para);
          if (x.end === '\x07') {
            cur.push(cell);
            cell = [];
          }
        }
        if (cur.length || cell.length) rows.push({ cells: cur.concat(cell.length ? [cell] : []), tdef: lastInfo.tdef, info: lastInfo });
        const table = buildTable(rows);
        if (table) blocks.push(table);
        i = j - 1;
        continue;
      }
      if (r.info.inTable) lossy.add('嵌套表格（已按普通段落显示）');
      const para = r.para;
      if (r.sect !== null) para.props.sectionBreak = breakType(sects[r.sect + 1]?.bkc ?? 2);
      blocks.push(para);
    }
    return blocks;
  };

  const breakType = (bkc: number): SectionBreak => (bkc === 0 ? 'continuous' : bkc === 3 ? 'evenPage' : bkc === 4 ? 'oddPage' : 'nextPage');

  const buildTable = (rows: { cells: Paragraph[][]; tdef: Sprm | null; info: ParaInfo }[]): Table | null => {
    if (!rows.length) return null;
    interface TC {
      first: boolean;
      merged: boolean;
      vmerge: boolean;
      vrestart: boolean;
      valign: number;
      /** 四条边（上、左、下、右）自己的边框；null = 没写，用表格的边框 */
      brc: (CellEdge | null)[];
    }
    const parsed = rows.map((r) => {
      const d = r.tdef?.data;
      const n = d ? d[0]! : r.cells.length;
      const edges: number[] = [];
      for (let k = 0; k <= n; k++) edges.push(d ? i16(d, 1 + 2 * k) : k * 100);
      const tcs: TC[] = [];
      const base = 1 + 2 * (n + 1);
      for (let k = 0; k < n; k++) {
        const o = base + 20 * k;
        const f = d && o + 20 <= d.length ? u16(d, o) : 0;
        const brc = [4, 8, 12, 16].map((b) => (d && o + b + 4 <= d.length ? brc80(d, o + b) : null));
        tcs.push({ first: !!(f & 1), merged: !!(f & 2), vmerge: !!(f & 0x20), vrestart: !!(f & 0x40), valign: (f >> 7) & 3, brc });
      }
      let gapHalf = 5.4;
      let jc: Table['align'] = 'left';
      for (const s of r.info.tableSprms) {
        if (s.op === 0x9602) gapHalf = i16(s.data, 0) / 20;
        else if (s.op === 0x5400) jc = s.data[0] === 1 ? 'center' : s.data[0] === 2 ? 'right' : 'left';
      }
      // 表格级边框（sprmTTableBorders：上、左、下、右、内部横线、内部竖线）和各单元格底纹
      const tb = r.info.tableSprms.find((x) => x.op === 0xd613)?.data;
      const tableBrc = tb && tb.length >= 48 ? [0, 1, 2, 3, 4, 5].map((i) => brc(tb, 8 * i)) : null;
      const shd = r.info.tableSprms.find((x) => x.op === 0xd612)?.data;
      const shd80 = r.info.tableSprms.find((x) => x.op === 0xd609)?.data;
      const shading = (k: number): string | undefined => {
        if (shd && 10 * (k + 1) <= shd.length) {
          const fore = shd.subarray(10 * k, 10 * k + 4);
          const back = shd.subarray(10 * k + 4, 10 * k + 8);
          const ipat = u16(shd, 10 * k + 8);
          const c = ipat === 1 ? fore : back;
          return c[3] ? undefined : [c[0], c[1], c[2]].map((v) => (v ?? 0).toString(16).padStart(2, '0')).join('').toUpperCase();
        }
        if (shd80 && 2 * (k + 1) <= shd80.length) {
          const v = u16(shd80, 2 * k);
          const ico = ipatIco(v);
          return ico || undefined;
        }
        return undefined;
      };
      return { edges, tcs, gapHalf, jc, row: r, tableBrc, shading };
    });
    // 网格列：所有行的列边界合并（容差 3 twips）
    const all = [...new Set(parsed.flatMap((p) => p.edges))].sort((a, b) => a - b);
    const grid: number[] = [];
    for (const e of all) if (!grid.length || e - grid[grid.length - 1]! > 3) grid.push(e);
    const snapEdge = (e: number) => grid.reduce((best, g) => (Math.abs(g - e) < Math.abs(best - e) ? g : best), grid[0]!);
    const gridOf = (e: number) => grid.indexOf(snapEdge(e));
    const tableRows: Row[] = parsed.map((p) => {
      const cells: Cell[] = [];
      let k = 0;
      while (k < p.tcs.length) {
        let e = k;
        if (p.tcs[k]!.first) while (e + 1 < p.tcs.length && p.tcs[e + 1]!.merged) e++;
        const g0 = gridOf(p.edges[k]!);
        const g1 = gridOf(p.edges[e + 1]!);
        const paras = p.row.cells[k] ?? [];
        const tc = p.tcs[k]!;
        const ri = parsed.indexOf(p);
        // 单元格自己的边框优先，没写的用表格边框：最外一圈用上 / 左 / 下 / 右，里面用内部横线 / 竖线
        const tbl = p.tableBrc;
        const pick = (own: CellEdge | null | undefined, outer: boolean, outerIdx: number, innerIdx: number): CellEdge | null => (own ?? (tbl ? (outer ? tbl[outerIdx] : tbl[innerIdx]) : null)) ?? null;
        const edges: CellEdges = {
          top: pick(tc.brc[0], ri === 0, 0, 4),
          left: pick(tc.brc[1], g0 === 0, 1, 5),
          bottom: pick(p.tcs[e]!.brc[2], ri === parsed.length - 1, 2, 4),
          right: pick(p.tcs[e]!.brc[3], g1 === grid.length - 1, 3, 5),
        };
        const shade = p.shading(k);
        cells.push({
          gridSpan: Math.max(1, g1 - g0),
          vMerge: tc.vmerge ? (tc.vrestart ? 'restart' : 'continue') : 'none',
          vAlign: tc.valign === 1 ? 'center' : tc.valign === 2 ? 'bottom' : 'top',
          edges,
          ...(shade ? { shading: shade } : {}),
          paragraphs: paras.length ? paras : [newParagraph()],
        });
        k = e + 1;
      }
      const h = p.row.info.rowHeight;
      return { height: h?.v ?? 0, heightRule: h?.rule ?? 'auto', header: p.row.info.header, cells };
    });
    const first = parsed[0]!;
    const table: Table = {
      kind: 'table',
      cols: grid.slice(1).map((g, i) => (g - grid[i]!) / 20),
      rows: tableRows,
      align: first.jc,
      // 第一条列边界就是表格左边框相对左页边距的位置（通常 -5.4 磅：边框伸进页边距，文字和正文左对齐）
      indent: Math.round((grid[0]! / 20) * 100) / 100,
      borders: tableRows.some((r) => r.cells.some((c) => c.edges && Object.values(c.edges).some(Boolean))),
      cellMargin: { left: first.gapHalf, right: first.gapHalf, top: 0, bottom: 0 },
    };
    return table;
  };

  // ---- 组装文档 ----
  const doc = newDocument();
  doc.defaults = { ...styleCharInfo(0).rp };
  const body = readStory(0, ccpText, true);
  doc.body = body.length ? body : [newParagraph()];
  // 表格后面必须有段落（Word 也一样）
  if (doc.body[doc.body.length - 1]!.kind === 'table') doc.body.push(newParagraph());

  const last = sects[sects.length - 1]!;
  doc.setup = { ...last.setup };
  if (last.cols) doc.setup.cols = last.cols;
  const [fcDop, lcbDop] = pair(31);
  doc.setup.evenAndOdd = fcDop + 1 < t.length && (t[fcDop]! & 1) === 1;
  // DopTypography（Dop97 偏移 90）的 iJustification（第 1–2 位）：0 不压缩，1 压缩标点，2 压缩标点和假名
  if (lcbDop > 91 && fcDop + 91 < t.length && ((t[fcDop + 90]! >> 1) & 3) !== 0) doc.setup.compressPunct = true;
  // 同一字节第 0 位 fKerningPunct：0 = 关闭标点挤压（同 docx 的 w:noPunctuationKerning）
  if (lcbDop > 91 && fcDop + 91 < t.length && (t[fcDop + 90]! & 1) === 0) doc.setup.noPunctKern = true;
  // Word 打开 .doc 一律是兼容模式 11；实测 Word 存出的 .doc 都带“另一套东亚断行规则”（useAltKinsokuLineBreakRules），
  // 它在 DOP 里的位还没有核实，先按开着处理
  doc.setup.compatMode = 11;
  doc.setup.altKinsoku = true;
  // DopBase 偏移 8 的 Copts60 第 13 位 fExpShRtn：0 = 以手动换行符结束的行不拉开（对应 docx 的 doNotExpandShiftReturn，
  // 用 Word 另存的样本核对过：该位为 0 的转出来带这个兼容选项）
  if (lcbDop > 10 && fcDop + 10 < t.length && ((t[fcDop + 8]! | (t[fcDop + 9]! << 8)) & (1 << 13)) === 0) doc.setup.noExpandShiftReturn = true;

  // 页眉页脚
  const [fcHdd, lcbHdd] = pair(11);
  const hddStart = ccpText + ccpFtn;
  const storyOf = (i: number): Paragraph[] | null => {
    if (!lcbHdd || 4 * (i + 2) > lcbHdd) return null;
    const a = u32(t, fcHdd + 4 * i);
    const b = u32(t, fcHdd + 4 * (i + 1));
    if (b <= a || a + hddStart >= total) return null;
    // 每个分区末尾多一个段落标记（分区的结束符），不是内容
    const to = hddStart + Math.min(b, ccpHdd);
    const endMark = String.fromCharCode(13);
    const blocks = readStory(hddStart + a, chars[to - 1] === endMark ? to - 1 : to, false);
    // 页眉页脚里的表格暂不能编辑：单元格里的段落按顺序排成普通段落
    if (blocks.some((x) => x.kind === 'table')) lossy.add('页眉页脚里的表格（已按普通段落显示）');
    const paras = blocks.flatMap((x) => (x.kind === 'p' ? [x] : x.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs))));
    return paras.length ? paras : [newHFParagraph()];
  };
  const hfs: SectionHF[] = [];
  const keys = ['headerEven', 'header', 'footerEven', 'footer', 'headerFirst', 'footerFirst'] as const;
  sects.forEach((s, si) => {
    const hf = { titlePg: s.setup.titlePg } as SectionHF;
    keys.forEach((key, k) => {
      const own = storyOf(6 + 6 * si + k);
      // 空 = 与上一节相同（第一节没有页眉就是空白）
      hf[key] = own ?? (si > 0 ? hfs[si - 1]![key] : [{ ...newHFParagraph(), hfNone: true }]);
    });
    hfs.push(hf);
  });
  const lastHF = hfs[hfs.length - 1]!;
  for (const key of keys) doc[key] = lastHF[key];
  doc.setup.titlePg = lastHF.titlePg;
  // 各节的设置记在结束它的分节段落上
  let si = 0;
  for (const b of doc.body) {
    if (b.kind !== 'p' || !b.props.sectionBreak) continue;
    const s = sects[si];
    if (s) {
      b.sectionHF = hfs[si];
      const g: PageGeom = { width: s.setup.width, height: s.setup.height, marginTop: s.setup.marginTop, marginBottom: s.setup.marginBottom, marginLeft: s.setup.marginLeft, marginRight: s.setup.marginRight, headerDistance: s.setup.headerDistance, footerDistance: s.setup.footerDistance };
      b.sectionPage = g;
      if (s.cols) b.sectionCols = s.cols;
    }
    si++;
  }

  if (ccpFtn > 0 || unsupported.footnote) lossy.add('脚注');
  if (ccpEdn > 0) lossy.add('尾注');
  if (ccpAtn > 0 || unsupported.note) lossy.add('批注');
  if (unsupported.shape) lossy.add('文本框和形状（找不到绘图数据）');
  doc.lists = lists;
  doc.lossy = [...lossy];
  return doc;
}

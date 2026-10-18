/**
 * 轻量 .xlsx 读写（fflate 解压 / 压缩 + 正则解析，不依赖 DOM）：
 * 读取每个工作表的单元格值（共享字符串、内联字符串、数字、布尔、公式缓存值），
 * 写出只含值的简单工作簿（可选表头 / 说明行样式与列宽），用于导入模版。
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

export type Cell = string | number | null;

export interface Sheet {
  name: string;
  /** 按行、列排列（空单元格为 null，行尾空单元格省略） */
  rows: Cell[][];
}

export interface SheetOut extends Sheet {
  /** 列宽（字符数） */
  widths?: number[];
  /**
   * 行样式：head 加粗浅蓝底，note 灰色小字；
   * 表格样式（参考 Excel 的“表格样式”）：th 深蓝底白色粗体表头、td 白底数据行、tdBand 浅蓝底数据行（与 td 交替）、tf 合计行（加粗、上方双线）
   */
  styles?: Record<number, RowStyle>;
  /** 金额列（从 0 起）：两位小数（0.00），存的仍是数字 */
  money?: number[];
  /** 居中的列（从 0 起）；th 表头行总是居中 */
  center?: number[];
  /** 筛选按钮的范围，如 `A1:L20`（表头行 + 数据行，不含合计行） */
  autoFilter?: string;
  /** 冻结首行（滚动时表头不动） */
  freezeHeader?: boolean;
}

export type RowStyle = 'head' | 'note' | 'th' | 'td' | 'tdBand' | 'tf';

const unescape = (s: string) =>
  s.replace(/&(lt|gt|quot|apos|amp|#\d+|#x[0-9a-fA-F]+);/g, (_m, e: string) => {
    if (e === 'lt') return '<';
    if (e === 'gt') return '>';
    if (e === 'quot') return '"';
    if (e === 'apos') return "'";
    if (e === 'amp') return '&';
    return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
  });

const escape = (s: string) =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // XML 1.0 不允许的控制字符
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');

const attr = (tag: string, name: string) => tag.match(new RegExp(`\\s${name}="([^"]*)"`))?.[1];

/** 富文本 / 普通文本节点里的文字（去掉拼音 rPh） */
function textOf(xml: string): string {
  const body = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  let out = '';
  for (const m of body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) out += m[1];
  return unescape(out);
}

/** 列字母 → 从 0 开始的列号（A → 0，AA → 26） */
export function colIndex(ref: string): number {
  const letters = ref.match(/^[A-Z]+/i)?.[0].toUpperCase() ?? 'A';
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function colName(i: number): string {
  let s = '';
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function resolve(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const p of target.split('/')) {
    if (p === '..') parts.pop();
    else if (p !== '.') parts.push(p);
  }
  return parts.join('/');
}

export function readXlsx(data: Uint8Array): Sheet[] {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(data);
  } catch {
    throw new Error('不是有效的 .xlsx 文件（请用 Excel 另存为 .xlsx 后再导入）');
  }
  const text = (p: string) => (files[p] ? strFromU8(files[p]) : '');
  const wb = text('xl/workbook.xml');
  if (!wb) throw new Error('不是有效的 .xlsx 文件');
  const rels = new Map<string, string>();
  for (const m of text('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], 'Id');
    const target = attr(m[0], 'Target');
    if (id && target) rels.set(id, resolve('xl/workbook.xml', target));
  }
  const shared = [...text('xl/sharedStrings.xml').matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => textOf(m[1]));

  const sheets: Sheet[] = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unescape(attr(m[0], 'name') ?? '');
    const path = rels.get(attr(m[0], 'r:id') ?? '');
    if (!path || !files[path]) continue;
    const rows: Cell[][] = [];
    let nextRow = 0;
    for (const r of text(path).matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const ri = Number(attr(r[1], 'r') ?? nextRow + 1) - 1;
      nextRow = ri + 1;
      const row: Cell[] = [];
      let nextCol = 0;
      for (const c of (r[2] ?? '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const ref = attr(c[1], 'r');
        const ci = ref ? colIndex(ref) : nextCol;
        nextCol = ci + 1;
        const t = attr(c[1], 't');
        const inner = c[2] ?? '';
        const v = inner.match(/<v>([\s\S]*?)<\/v>/)?.[1];
        let val: Cell = null;
        if (t === 'inlineStr') val = textOf(inner);
        else if (v === undefined) val = null;
        else if (t === 's') val = shared[Number(v)] ?? '';
        else if (t === 'str' || t === 'e') val = unescape(v);
        else if (t === 'b') val = v === '1' ? 'TRUE' : 'FALSE';
        else val = Number.isFinite(Number(v)) ? Number(v) : unescape(v);
        if (val === '') val = null;
        while (row.length < ci) row.push(null);
        row[ci] = val;
      }
      while (row.length && row[row.length - 1] === null) row.pop();
      rows[ri] = row;
    }
    for (let i = 0; i < rows.length; i++) rows[i] ??= [];
    sheets.push({ name, rows });
  }
  return sheets;
}

/** 行样式 → 字体 / 底色 / 边框（下标对应下面 STYLES 里的 fonts / fills / borders） */
const ROW_LOOK: Record<RowStyle | 'none', { font: number; fill: number; border: number }> = {
  none: { font: 0, fill: 0, border: 0 },
  head: { font: 1, fill: 2, border: 0 },
  note: { font: 2, fill: 0, border: 0 },
  th: { font: 3, fill: 3, border: 1 },
  td: { font: 0, fill: 0, border: 1 },
  tdBand: { font: 0, fill: 4, border: 1 },
  tf: { font: 1, fill: 4, border: 2 },
};
const ROW_ORDER = Object.keys(ROW_LOOK) as (RowStyle | 'none')[];

/** cellXfs 的下标：行样式 × 是否金额（numFmtId 2 = 0.00，Excel 内置）× 是否居中 */
export function cellStyleIndex(row: RowStyle | undefined, money: boolean, center = false): number {
  return (ROW_ORDER.indexOf(row ?? 'none') * 2 + (money ? 1 : 0)) * 2 + (center ? 1 : 0);
}

const CELL_XFS = ROW_ORDER.flatMap((r) =>
  [false, true].flatMap((money) =>
    [false, true].map((center) => {
      const { font, fill, border } = ROW_LOOK[r];
      const attrs = `numFmtId="${money ? 2 : 0}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${money ? ' applyNumberFormat="1"' : ''}${font ? ' applyFont="1"' : ''}${fill ? ' applyFill="1"' : ''}${border ? ' applyBorder="1"' : ''}`;
      return center ? `<xf ${attrs} applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>` : `<xf ${attrs}/>`;
    }),
  ),
);

const LINE = '<color rgb="FF8EA9DB"/>';
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="11"/><name val="等线"/></font><font><b/><sz val="11"/><name val="等线"/></font><font><sz val="10"/><color rgb="FF808080"/><name val="等线"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="等线"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin">${LINE}</left><right style="thin">${LINE}</right><top style="thin">${LINE}</top><bottom style="thin">${LINE}</bottom><diagonal/></border><border><left style="thin">${LINE}</left><right style="thin">${LINE}</right><top style="double"><color rgb="FF4472C4"/></top><bottom style="thin">${LINE}</bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${CELL_XFS.length}">${CELL_XFS.join('')}</cellXfs>
<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

/** 有底色 / 边框的行：空单元格也写出来，整行才连成一片 */
const FULL_ROW = new Set<RowStyle>(['head', 'th', 'td', 'tdBand', 'tf']);

function sheetXml(s: SheetOut): string {
  const money = new Set(s.money ?? []);
  const center = new Set(s.center ?? []);
  const cols = s.widths?.length
    ? `<cols>${s.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>`
    : '';
  const rows = s.rows
    .map((row, ri) => {
      const st = s.styles?.[ri];
      const width = st && FULL_ROW.has(st) ? Math.max(row.length, s.widths?.length ?? 0) : row.length;
      const cells = Array.from({ length: width }, (_, ci) => row[ci] ?? null)
        .map((v, ci) => {
          const idx = cellStyleIndex(st, money.has(ci) && typeof v === 'number', st === 'th' || center.has(ci));
          const sAttr = idx ? ` s="${idx}"` : '';
          const ref = `${colName(ci)}${ri + 1}`;
          if (v === null || v === '') return st && FULL_ROW.has(st) ? `<c r="${ref}"${sAttr}/>` : '';
          if (typeof v === 'number') return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
          return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${escape(v)}</t></is></c>`;
        })
        .join('');
      return `<row r="${ri + 1}">${cells}</row>`;
    })
    .join('');
  const views = s.freezeHeader
    ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    : '';
  const filter = s.autoFilter ? `<autoFilter ref="${s.autoFilter}"/>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${views}${cols}<sheetData>${rows}</sheetData>${filter}</worksheet>`;
}

export function writeXlsx(sheets: SheetOut[]): Uint8Array {
  const ct = sheets
    .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
    .join('');
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${ct}</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets
      .map((s, i) => `<sheet name="${escape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join('')}</sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
      .join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`),
    'xl/styles.xml': strToU8(STYLES),
  };
  sheets.forEach((s, i) => (files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml(s))));
  return zipSync(files);
}

/** 供表格查看器（`src/lib/sheet/workbook.ts`）复用的解析小工具 */
export { unescape as unescapeXml, attr as xmlAttr, textOf as xmlText, resolve as resolvePart };

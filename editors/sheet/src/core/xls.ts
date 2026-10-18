/** 后端读出的 .xls（`read_xls`）→ 表格模型。.xls 不带数字格式，数值按“常规”显示，日期按是否含时间套日期格式。 */
import type { XlsSheet } from '@wz/editor-shared';
import { formatGeneral, formatNumber } from './numfmt';
import { allCellStyles, deriveStyle, readStyleBook } from './styles';
import { MIN_ROW_HEIGHT, DEFAULT_COL_WIDTH, type SCell, type SheetData, type Workbook } from './workbook';

/** 日期序列号用的格式：纯日期、纯时间或日期 + 时间 */
export function dateFormat(serial: number): string {
  const frac = serial - Math.floor(serial);
  if (serial < 1) return 'h:mm:ss';
  if (Math.abs(frac) < 1e-9) return 'yyyy/m/d';
  return 'yyyy/m/d h:mm';
}

export function dateText(serial: number): string {
  return formatNumber(serial, dateFormat(serial));
}

export function fromXls(sheets: XlsSheet[]): Workbook {
  const styleBook = readStyleBook('');
  const dateStyles = new Map<string, number>();
  const styleFor = (code: string) => {
    let s = dateStyles.get(code);
    if (s === undefined) dateStyles.set(code, (s = deriveStyle(styleBook, 0, { numFmt: code })));
    return s;
  };
  const out: Workbook = {
    kind: 'xls',
    styleBook,
    styles: [],
    sheets: sheets.map((s): SheetData => {
      const [r0, c0] = s.start;
      const cells: (SCell | undefined)[][] = [];
      let rows = 0;
      let cols = 0;
      s.rows.forEach((row, i) => {
        row.forEach((v, j) => {
          let cell: SCell | undefined;
          if (v === null) return;
          if (typeof v === 'number') cell = { v, t: formatGeneral(v) };
          else if (typeof v === 'boolean') cell = { v, t: v ? 'TRUE' : 'FALSE' };
          else if (typeof v === 'string') cell = { v, t: v };
          else cell = { v: v.d, t: dateText(v.d), s: styleFor(dateFormat(v.d)) };
          const r = r0 + i;
          const c = c0 + j;
          (cells[r] ??= [])[c] = cell;
          rows = Math.max(rows, r + 1);
          cols = Math.max(cols, c + 1);
        });
      });
      const merges = s.merges
        .map(([a, b, c, d]) => ({ r0: a, c0: b, r1: c, c1: d }))
        .filter((m) => m.r1 > m.r0 || m.c1 > m.c0);
      for (const m of merges) {
        rows = Math.max(rows, m.r1 + 1);
        cols = Math.max(cols, m.c1 + 1);
      }
      return {
        name: s.name,
        cells,
        rows,
        cols,
        merges,
        colWidths: {},
        rowHeights: {},
        defaultColWidth: DEFAULT_COL_WIDTH + 24,
        defaultRowHeight: MIN_ROW_HEIGHT + 2,
        freeze: { rows: 0, cols: 0 },
        colDefs: [],
        rowDefs: {},
      };
    }),
  };
  out.styles = allCellStyles(styleBook);
  return out;
}

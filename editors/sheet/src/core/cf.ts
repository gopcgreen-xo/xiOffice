/**
 * 条件格式的显示（同 Excel 的规则）：单元格值比较、公式、包含文字、空 / 错误、重复 / 唯一、前 N 项、高于 / 低于平均、
 * 色阶、数据条。多条规则按优先级合并：每一项格式由优先级最高的满足规则决定；stopIfTrue 满足后不再看后面的规则。
 * 公式里的相对引用以规则第一块区域的左上角为准。图标集暂不显示。
 */
import { moveFormula } from './formula/refs';
import { compare, FErr, type Engine, type Scalar } from './formula/eval';
import { cellKey } from './grid';
import type { CellStyle } from './styles';
import type { CfRule, Range, SCell, Workbook } from './workbook';

export interface CfResult {
  style?: CellStyle;
  /** 数据条：长度（0–1）与颜色 */
  bar?: { pct: number; color: string };
}

const cellValue = (c: SCell | undefined): Scalar => (!c ? null : c.k === 'e' ? new FErr(String(c.v)) : c.v === '' ? null : c.v);

/** 规则覆盖的格子（裁到数据范围） */
function cellsOf(rule: CfRule, rows: number, cols: number): { r: number; c: number }[] {
  const out: { r: number; c: number }[] = [];
  for (const g of rule.ranges) for (let r = g.r0; r <= Math.min(g.r1, rows - 1); r++) for (let c = g.c0; c <= Math.min(g.c1, cols - 1); c++) out.push({ r, c });
  return out;
}

function hex(c: string): number[] {
  return [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16));
}
function mixColor(a: string, b: string, t: number): string {
  const x = hex(a);
  const y = hex(b);
  return `#${x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('')}`;
}

/** 色阶 / 数据条的阈值：cfvo（min / max / num / percent / percentile / formula） */
function threshold(v: { type: string; val?: string }, nums: number[], evalNum: (f: string) => number): number {
  const sorted = [...nums].sort((a, b) => a - b);
  const min = sorted[0] ?? 0;
  const max = sorted[sorted.length - 1] ?? 0;
  const n = Number(v.val);
  switch (v.type) {
    case 'min':
      return min;
    case 'max':
      return max;
    case 'percent':
      return min + ((max - min) * n) / 100;
    case 'percentile': {
      if (!sorted.length) return 0;
      const k = ((sorted.length - 1) * n) / 100;
      const lo = Math.floor(k);
      return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (k - lo);
    }
    case 'formula':
      return evalNum(v.val ?? '0');
    default:
      return Number.isFinite(n) ? n : evalNum(v.val ?? '0');
  }
}

export function evalCf(book: Workbook, si: number, engine: Engine | null): Map<number, CfResult> {
  const sheet = book.sheets[si];
  const out = new Map<number, CfResult>();
  const rules = sheet.cf ?? [];
  if (!rules.length) return out;
  const dxfs = book.styleBook?.dxfs ?? [];
  const stopped = new Set<number>();
  const get = (r: number, c: number) => cellValue(sheet.cells[r]?.[c]);
  for (const rule of rules) {
    const cells = cellsOf(rule, sheet.rows, sheet.cols);
    if (!cells.length) continue;
    const g0: Range = rule.ranges[0];
    const evalAt = (f: string, r: number, c: number): Scalar => {
      if (!engine) return null;
      return engine.evalFormula(moveFormula(f, r - g0.r0, c - g0.c0), si, r, c);
    };
    const values = cells.map(({ r, c }) => get(r, c));
    const nums = values.filter((v): v is number => typeof v === 'number');
    let test: ((v: Scalar, r: number, c: number, i: number) => boolean) | null = null;
    switch (rule.type) {
      case 'cellIs': {
        test = (v, r, c) => {
          if (v === null && rule.operator !== 'equal' && rule.operator !== 'notEqual') return false;
          const a = evalAt(rule.formulas[0] ?? '0', r, c);
          const b = rule.formulas[1] !== undefined ? evalAt(rule.formulas[1], r, c) : null;
          if (v instanceof FErr || a instanceof FErr) return false;
          const d = compare(v, a);
          switch (rule.operator) {
            case 'lessThan':
              return d < 0;
            case 'lessThanOrEqual':
              return d <= 0;
            case 'equal':
              return d === 0;
            case 'notEqual':
              return d !== 0;
            case 'greaterThanOrEqual':
              return d >= 0;
            case 'greaterThan':
              return d > 0;
            case 'between':
            case 'notBetween': {
              const lo = compare(a, b) <= 0 ? a : b;
              const hi = lo === a ? b : a;
              const inside = compare(v, lo) >= 0 && compare(v, hi) <= 0;
              return rule.operator === 'between' ? inside : !inside;
            }
          }
          return false;
        };
        break;
      }
      case 'expression':
        test = (_v, r, c) => {
          const x = evalAt(rule.formulas[0] ?? 'FALSE', r, c);
          return x === true || (typeof x === 'number' && x !== 0);
        };
        break;
      case 'containsText':
      case 'notContainsText':
      case 'beginsWith':
      case 'endsWith': {
        const t = (rule.text ?? '').toLowerCase();
        test = (_v, r, c) => {
          const s = (sheet.cells[r]?.[c]?.t ?? '').toLowerCase();
          return rule.type === 'containsText' ? s.includes(t) : rule.type === 'notContainsText' ? !s.includes(t) : rule.type === 'beginsWith' ? s.startsWith(t) : s.endsWith(t);
        };
        break;
      }
      case 'containsBlanks':
        test = (v) => v === null || (typeof v === 'string' && !v.trim());
        break;
      case 'notContainsBlanks':
        test = (v) => !(v === null || (typeof v === 'string' && !v.trim()));
        break;
      case 'containsErrors':
        test = (v) => v instanceof FErr;
        break;
      case 'notContainsErrors':
        test = (v) => !(v instanceof FErr);
        break;
      case 'duplicateValues':
      case 'uniqueValues': {
        const key = (v: Scalar) => (v === null ? null : typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${String(v)}`);
        const count = new Map<string, number>();
        for (const v of values) {
          const k = key(v);
          if (k) count.set(k, (count.get(k) ?? 0) + 1);
        }
        test = (v) => {
          const k = key(v);
          if (!k) return false;
          return rule.type === 'duplicateValues' ? count.get(k)! > 1 : count.get(k) === 1;
        };
        break;
      }
      case 'top10': {
        const sorted = [...nums].sort((a, b) => (rule.bottom ? a - b : b - a));
        const n = rule.percent ? Math.max(1, Math.floor((sorted.length * (rule.rank ?? 10)) / 100)) : (rule.rank ?? 10);
        const edge = sorted[Math.min(n, sorted.length) - 1];
        test = (v) => typeof v === 'number' && edge !== undefined && (rule.bottom ? v <= edge : v >= edge);
        break;
      }
      case 'aboveAverage': {
        const avg = nums.reduce((a, b) => a + b, 0) / (nums.length || 1);
        test = (v) => typeof v === 'number' && (rule.above === false ? v < avg : v > avg);
        break;
      }
      case 'colorScale': {
        const colors = rule.colors ?? [];
        const vo = rule.cfvo ?? [];
        if (colors.length < 2 || vo.length < 2) break;
        const num = (f: string) => {
          const x = evalAt(f, g0.r0, g0.c0);
          return typeof x === 'number' ? x : 0;
        };
        const stops = vo.map((v) => threshold(v, nums, num));
        cells.forEach(({ r, c }, i) => {
          const v = values[i];
          const k = cellKey(r, c);
          if (typeof v !== 'number' || stops.length < 2 || out.get(k)?.style?.fill || stopped.has(k)) return;
          let fill: string;
          if (colors.length >= 3 && stops.length >= 3) {
            fill = v <= stops[1] ? mixColor(colors[0], colors[1], clamp((v - stops[0]) / (stops[1] - stops[0] || 1))) : mixColor(colors[1], colors[2], clamp((v - stops[1]) / (stops[2] - stops[1] || 1)));
          } else fill = mixColor(colors[0], colors[1], clamp((v - stops[0]) / (stops[1] - stops[0] || 1)));
          const cur = out.get(k) ?? {};
          out.set(k, { ...cur, style: { ...cur.style, fill } });
        });
        continue;
      }
      case 'dataBar': {
        const vo = rule.cfvo ?? [];
        const num = (f: string) => {
          const x = evalAt(f, g0.r0, g0.c0);
          return typeof x === 'number' ? x : 0;
        };
        const lo = vo[0] ? threshold(vo[0], nums, num) : Math.min(...nums);
        const hi = vo[1] ? threshold(vo[1], nums, num) : Math.max(...nums);
        const color = rule.colors?.[0] ?? '#638ec6';
        cells.forEach(({ r, c }, i) => {
          const v = values[i];
          const k = cellKey(r, c);
          if (typeof v !== 'number' || stopped.has(k) || out.get(k)?.bar) return;
          // Excel 的数据条最短也有一点长度
          const pct = 0.1 + 0.9 * clamp((v - lo) / (hi - lo || 1));
          out.set(k, { ...out.get(k), bar: { pct, color } });
        });
        continue;
      }
    }
    if (!test) continue;
    const dxf = rule.dxf !== undefined ? dxfs[rule.dxf] : undefined;
    cells.forEach(({ r, c }, i) => {
      const k = cellKey(r, c);
      if (stopped.has(k)) return;
      let hit = false;
      try {
        hit = test!(values[i], r, c, i);
      } catch {
        hit = false;
      }
      if (!hit) return;
      if (rule.stopIfTrue) stopped.add(k);
      if (!dxf) return;
      // 优先级高的规则先处理：已经有的项不覆盖
      const cur = out.get(k) ?? {};
      out.set(k, { ...cur, style: { ...dxf, ...cur.style } });
    });
  }
  return out;
}

const clamp = (x: number) => Math.max(0, Math.min(1, Number.isFinite(x) ? x : 0));

/** 自动编号：编号格式化、预设列表、按文档顺序计算每段的编号文字。 */
import { type Doc, type NumFmt, type NumLevel, type NumList, type Paragraph } from './model';

const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
const TIANGAN = ['甲', '乙', '丙', '丁', '戊', '己', '庚', '辛', '壬', '癸'];
const CIRCLED = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳';

/** 中文计数：1→一，10→十，11→十一，20→二十，105→一百零五 */
export function chineseNumber(n: number): string {
  if (n <= 0) return String(n);
  if (n < 10) return CN_DIGITS[n];
  if (n < 20) return '十' + (n % 10 ? CN_DIGITS[n % 10] : '');
  if (n < 100) return CN_DIGITS[Math.floor(n / 10)] + '十' + (n % 10 ? CN_DIGITS[n % 10] : '');
  if (n < 1000) {
    const rest = n % 100;
    return CN_DIGITS[Math.floor(n / 100)] + '百' + (rest === 0 ? '' : rest < 10 ? '零' + CN_DIGITS[rest] : rest < 20 ? '一' + chineseNumber(rest) : chineseNumber(rest));
  }
  return String(n);
}

function roman(n: number): string {
  const map: [number, string][] = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let out = '';
  for (const [v, r] of map) while (n >= v) {
    out += r;
    n -= v;
  }
  return out;
}

function letter(n: number): string {
  // Word：a…z，然后 aa…zz
  const c = String.fromCharCode(97 + ((n - 1) % 26));
  return c.repeat(Math.floor((n - 1) / 26) + 1);
}

export function formatNumber(n: number, fmt: NumFmt): string {
  switch (fmt) {
    case 'chineseCounting':
      return chineseNumber(n);
    case 'ideographTraditional':
      return TIANGAN[(n - 1) % 10] ?? String(n);
    case 'upperLetter':
      return letter(n).toUpperCase();
    case 'lowerLetter':
      return letter(n);
    case 'upperRoman':
      return roman(n);
    case 'lowerRoman':
      return roman(n).toLowerCase();
    case 'decimalEnclosedCircle':
      return n >= 1 && n <= 20 ? CIRCLED[n - 1] : String(n);
    case 'bullet':
    case 'none':
      return '';
    default:
      return String(n);
  }
}

// ---------------- 预设 ----------------

const hanging = (lvl: number, fmt: NumFmt, text: string): NumLevel => ({
  fmt,
  text,
  start: 1,
  indentLeft: 21 * (lvl + 1),
  firstLine: -21,
  firstLineChars: null,
  suffix: 'tab',
});
/** 中文公文习惯：编号与正文同行，首行缩进 2 字符，转行顶格 */
const inline = (fmt: NumFmt, text: string): NumLevel => ({ fmt, text, start: 1, indentLeft: 0, firstLine: 0, firstLineChars: 2, suffix: 'nothing' });

function deepDecimal(from: number, levels: NumLevel[]): NumLevel[] {
  for (let l = from; l < 9; l++) levels.push(hanging(l, 'decimal', Array.from({ length: l + 1 }, (_, k) => `%${k + 1}`).join('.') + '.'));
  return levels;
}

export const NUM_PRESETS: { key: string; label: string; levels: () => NumLevel[] }[] = [
  { key: 'decimal', label: '1. 2. 3.', levels: () => deepDecimal(0, []) },
  {
    key: 'gongwen',
    label: '一、（一）1.（1）',
    levels: () => deepDecimal(4, [inline('chineseCounting', '%1、'), inline('chineseCounting', '（%2）'), inline('decimal', '%3.'), inline('decimal', '（%4）')]),
  },
  { key: 'chineseParen', label: '（一）（二）', levels: () => deepDecimal(1, [inline('chineseCounting', '（%1）')]) },
  { key: 'decimalParen', label: '（1）（2）', levels: () => deepDecimal(1, [inline('decimal', '（%1）')]) },
  { key: 'circle', label: '① ② ③', levels: () => deepDecimal(1, [hanging(0, 'decimalEnclosedCircle', '%1')]) },
  { key: 'letter', label: 'a) b) c)', levels: () => deepDecimal(1, [hanging(0, 'lowerLetter', '%1)')]) },
  {
    key: 'bullet',
    label: '● 项目符号',
    levels: () => ['●', '○', '■', '●', '○', '■', '●', '○', '■'].map((b, l) => hanging(l, 'bullet', b)),
  },
];

/** 找到或新建某个预设的列表（同一文档中同一预设连续编号）。 */
export function ensureList(doc: Doc, preset: string): NumList {
  const found = doc.lists.find((l) => l.preset === preset);
  if (found) return found;
  const def = NUM_PRESETS.find((p) => p.key === preset);
  if (!def) throw new Error(`未知编号样式：${preset}`);
  const list: NumList = { id: Math.max(0, ...doc.lists.map((l) => l.id)) + 1, preset, levels: def.levels() };
  doc.lists.push(list);
  return list;
}

// ---------------- 计算编号文字 ----------------

/** 按文档顺序（正文段落与表格中的段落）为每个编号段落算出编号文字。 */
export function computeLabels(doc: Doc): WeakMap<Paragraph, string> {
  const out = new WeakMap<Paragraph, string>();
  const counters = new Map<number, number[]>();
  const visit = (p: Paragraph) => {
    const num = p.props.num;
    if (!num) return;
    const list = doc.lists.find((l) => l.id === num.id);
    const lvl = list?.levels[num.level];
    if (!list || !lvl) return;
    let c = counters.get(list.id);
    if (!c) {
      c = list.levels.map((l) => l.start - 1);
      counters.set(list.id, c);
    }
    c[num.level] += 1;
    for (let j = num.level + 1; j < c.length; j++) c[j] = list.levels[j].start - 1;
    const text =
      lvl.fmt === 'bullet'
        ? lvl.text
        : lvl.text.replace(/%([1-9])/g, (_, k: string) => {
            const i = Number(k) - 1;
            const L = list.levels[i];
            return L ? formatNumber(Math.max(c![i], L.start), L.fmt) : '';
          });
    out.set(p, text);
  };
  for (const b of doc.body) {
    if (b.kind === 'p') visit(b);
    else for (const row of b.rows) for (const cell of row.cells) cell.paragraphs.forEach(visit);
  }
  return out;
}

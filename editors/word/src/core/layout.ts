/**
 * 排版引擎：段落断行、文档网格、对齐、行距、分页、表格、页眉页脚。
 * 规则参照 Word 中文版式的默认行为：
 * - 汉字之间可断行；西文单词不拆开（整行放不下时才强制拆）；
 * - 行首/行尾禁则；“，。”等标点允许溢出右边界（标点溢出边界）；
 * - 中文与西文字母/数字之间自动加 1/4 字宽间距；
 * - 指定行网格时，行高按网格行距取整（单倍行距 = 1 格，1.5 倍 = 1.5 格）；
 *   指定字符网格时，每个汉字占整数个字符格；
 * - 两端对齐的末行、以及分页符前的行不拉伸；分散对齐每行都拉伸；
 * - 孤行控制：段首一行不单独留在页底，段末一行不单独到下一页；
 * - 表格行不跨页拆分，标题行跨页重复。
 */
import {
  FIELD_CHAR,
  charPitch,
  columnBoxes,
  hfPath,
  hfStoryFor,
  newParagraph,
  isHeaderStory,
  sectionList,
  textWidth,
  type Doc,
  type EqDisplay,
  isAtom,
  type FieldKind,
  type HFStory,
  type PageSetup,
  type Paragraph,
  type Run,
  type RunProps,
  type SectionBreak,
  type SectionInfo,
  type Story,
  type RunImage,
  type TabStop,
  type Table,
  type BorderStyle,
  type CellEdge,
  type TextBox,
  textHeight,
} from './model';
import { fontEpoch } from './fontfile';
import { codeText, fieldCodeEpoch, fieldSpans, showsCode } from './fieldView';
import { MATH_FONT, layoutMath, type MathBox } from './math';
import { computeLabels } from './numbering';
import { pictureBounds, pictureKey } from './picture';
import { HANGING, NO_LINE_START, canBreakBetween, isCJK, isLatinAlnum } from './text';

/** 字体度量，均为相对字号的比例（1 表示等于字号）。 */
export interface Measurer {
  width(ch: string, family: string, bold: boolean, italic: boolean): number;
  /** 单倍行距下字体的上伸和下伸（em） */
  metrics(family: string): { ascent: number; descent: number };
  /** 测量结果变化的次数（字体加载完成后重测），排版缓存的键里带上 */
  epoch?: number;
}

/** 一行里参与算行高的文字：字符格式和实际用到的字体（缺省时中西文字体都算，空段落用段落标记的格式） */
interface LineFont {
  pr: RunProps;
  family?: string;
}

export interface Glyph {
  ch: string;
  /** Source paragraph offset, used to select atomic pictures by clicking them. */
  offset?: number;
  x: number;
  w: number;
  props: RunProps;
  family: string;
  cjk: boolean;
  /** 嵌入式图片（w / h 为排版后的显示尺寸，过宽时已按比例缩小）；底边落在基线上 */
  image?: RunImage;
  /** 制表符的前导符（目录里的“……”），画满整个宽度 */
  leader?: string;
  /** 公式（只读显示）：排版好的绘制指令，基线与行的基线对齐 */
  math?: MathBox;
  /** EQ 域（带圈字符等）：各部分及其宽度，在本字形宽度里叠画或依次画 */
  eq?: EqLayout;
}

export interface EqLayout extends EqDisplay {
  widths: number[];
}

export interface Line {
  /** 段内文本偏移区间 [start, end) */
  start: number;
  end: number;
  /** 每个偏移（start..end）的光标 x，相对段落区域左边 */
  xs: number[];
  glyphs: Glyph[];
  /** 可见内容右端（不含行尾空格） */
  right: number;
  height: number;
  /** 基线到行顶的距离 */
  baseline: number;
  /** 分页时算“放得下”用的高度（网格行：文字底部，行框可以伸进下边距），缺省 = height */
  fit?: number;
  /** wrap 自动换行；line 手动换行；page 分页符；end 段末 */
  brk: 'wrap' | 'line' | 'page' | 'column' | 'end';
  /** 锚定在这一行里的文本框 */
  anchors?: Anchor[];
  /** 这一行里引用的脚注（id），脚注排在这一行所在页的页底 */
  notes?: string[];
}

/** 行里的文本框锚点；areaW 为段落所在区域（栏 / 单元格）的宽度 */
export interface Anchor {
  box: TextBox;
  areaW: number;
}

/** 排好位置的文本框（页面坐标，磅） */
export interface FloatBox {
  x: number;
  y: number;
  box: TextBox;
  lines: PlacedLine[];
  /** 锚定在页眉页脚里的（Word 把它们画在正文下面一层，编辑正文时不能选中） */
  hf?: boolean;
}

export interface ParaLayout {
  lines: Line[];
  spaceBefore: number;
  spaceAfter: number;
}

export interface PlacedLine {
  story: Story;
  path: number[];
  line: Line;
  /** 段落区域左边（页面坐标，磅） */
  x: number;
  y: number;
  /** 段落最后一行：段后的分节符（用于显示编辑标记） */
  sectionBreak?: SectionBreak;
  /** 只读（页眉页脚里的表格）：点击定位不落在这里 */
  ro?: boolean;
}

export interface Seg {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  /** 线宽（磅），缺省 0.5 */
  w?: number;
  /** 颜色 '#RRGGBB'，缺省黑 */
  color?: string;
  /** 线型：缺省单实线 */
  style?: BorderStyle;
}

export interface PageLayout {
  index: number;
  /** 文字水印（所有页相同；文字下面一层画） */
  watermark?: Doc['watermark'];
  /** 本页所属的节（sectionList 下标）：页面开始处所在的节 */
  section: number;
  /** 本页的纸张与页边距（所属节的页面设置） */
  setup: PageSetup;
  /** 本页使用的页眉页脚分区，owner 为该节页眉页脚所在的分节段落（见 hfPath） */
  hf: { header: HFStory; footer: HFStory; owner: number | null };
  body: PlacedLine[];
  header: PlacedLine[];
  footer: PlacedLine[];
  borders: Seg[];
  /** 表格单元格底纹（先于文字画） */
  fills: { x: number; y: number; w: number; h: number; color: string }[];
  /** 表格单元格在本页的位置（选择单元格区域、拖动列线用）；g0..g1 为所占网格列 */
  cells?: { bi: number; r: number; c: number; g0: number; g1: number; x: number; y: number; w: number; h: number }[];
  /** 正文里的文本框 */
  floats: FloatBox[];
  /** 页底的脚注（引用在本页的脚注，按引用顺序）和脚注区的高度 */
  noteIds?: string[];
  noteArea?: number;
  notes?: PlacedLine[];
}

export interface DocLayout {
  pages: PageLayout[];
}

export interface Pos {
  story: Story;
  path: number[];
  offset: number;
}

interface Env {
  m: Measurer;
  doc: Doc;
  /** 是否使用行网格 / 字符网格 */
  lineGrid: boolean;
  charGrid: boolean;
  fieldText: (f: FieldKind) => string;
  /** 自动编号段落的编号文字 */
  labels?: WeakMap<Paragraph, string>;
  /** 脚注编号（按引用在正文里出现的顺序）；noteSelf = 正在排的脚注自己的编号 */
  notes?: Map<string, number>;
  noteSelf?: number;
}

interface Item {
  ch: string;
  off: number;
  len: number;
  w: number;
  gap: number;
  props: RunProps;
  family: string;
  cjk: boolean;
  kind: 'char' | 'tab' | 'break';
  image?: RunImage;
  anchor?: Anchor;
  leader?: string;
  /** 脚注引用（编号显示为上标） */
  note?: string;
  /** 公式 */
  math?: MathBox;
  /** 相邻标点挤压掉的左半边（开始类标点）：字形往左画这么多 */
  trimL?: number;
  eq?: EqLayout;
}

const LEADER_CHAR: Record<string, string> = { dot: '.', hyphen: '-', underscore: '_' };

/**
 * 制表符宽度：到下一个自定义制表位（左 / 居中 / 右对齐），没有时按默认制表位间隔。
 * 右对齐制表位让后面的文字（到下一个制表符或换行为止）右边对齐到制表位，居中则文字中点对齐。
 */
function tabWidth(items: Item[], j: number, pos: number, tabs: TabStop[] | undefined, def: number): number {
  const it = items[j];
  const stop = tabs?.find((t) => t.pos > pos + EPS);
  if (!stop) return Math.ceil((pos + EPS) / def) * def - pos;
  if (stop.leader !== 'none') it.leader = LEADER_CHAR[stop.leader];
  if (stop.align === 'left') return stop.pos - pos;
  let follow = 0;
  for (let k = j + 1; k < items.length && items[k].kind === 'char'; k++) follow += items[k].w + items[k].gap;
  return Math.max(0, stop.pos - pos - (stop.align === 'right' ? follow : follow / 2));
}

const EPS = 0.01;
/** 段前 / 段后“自动”（HTML 式间距）的磅值 */
const AUTO_SPACING = 14;

/**
 * 两端对齐的行稍微放不下时，Word 把空格和中西文自动间距压窄（最多压掉各自的 1/4）来多放一点；
 * 文档设了压缩标点（compressPunctuation）时全角标点还可以压掉一半。
 * 左对齐 / 居中 / 右对齐不压（Word 实测，scripts/word-measure.mjs：左对齐的“中：中”仍是三个整字宽）。
 */
const SHRINK = 0.25;
/** 可压缩的全角标点：开始的标点字形在右半边（压掉左边），其余在左半边（压掉右边） */
const PUNCT_OPEN = new Set([...'（［【《〈「『〔｛“‘']);
const PUNCT_CLOSE = new Set([...'，。、；：？！）］】》〉」』〕｝”’']);
function punctShrink(it: Item, compress: boolean): number {
  if (!compress || it.kind !== 'char' || !(PUNCT_OPEN.has(it.ch) || PUNCT_CLOSE.has(it.ch))) return 0;
  // 字体里本来就是半角宽的标点（等线的引号）不再压
  return it.w >= 0.8 * glyphSize(it.props) ? it.w / 2 : 0;
}
/** 空格和中西文间距可压的量（标点另算，见 punctShrink） */
function shrinkable(it: Item, first: boolean): number {
  if (it.kind !== 'char') return 0;
  return SHRINK * ((first ? 0 : it.gap) + (it.ch === ' ' ? it.w : 0));
}

/**
 * 行末的字放不下时，Word 靠压缩行内标点把它“挤进”本行还是“推出”到下一行：比较两种做法的变形程度。
 * Word 实测（宋体 10.5 / 12 磅，两端对齐、压缩标点，一行 20–49 字、K = 1–12 个可压标点，标点位置不影响）：
 * 挤进去允许的最大超出量 o = 字宽 × aK / (W / 字宽 + aK)，W 是这一行在该字之前的自然宽度，a = 33；
 * 同时不超过各标点半宽之和。等价于：超出量 / 可压总量 ≤ 66 × 推出后的空白 / W。
 * 该字后面紧跟行尾要悬挂的标点（。，等）时 Word 更倾向推出，a ≈ 21（近似：实测误差约 ±0.5 磅）。
 * 以上是兼容模式 15；兼容模式 11、12、14 不论几个标点，整行最多只压掉一个标点的半宽（Word 实测，与 useAltKinsoku 无关）。
 */
const SQUEEZE_A = 33;
const SQUEEZE_A_HANG = 21;
class PunctCap {
  sum = 0;
  max = 0;
  n = 0;
  constructor(private legacy = false) {}
  add(v: number) {
    if (v <= 0) return;
    this.sum += v;
    this.max = Math.max(this.max, v);
    this.n++;
  }
  /** 宽度 em 的字排在自然宽度 W 之后时，标点最多能压掉多少 */
  capFor(W: number, em: number, a = SQUEEZE_A): number {
    if (!this.n || em <= 0) return 0;
    if (this.legacy) return this.max;
    return Math.min(this.sum, (em * a * this.n) / (W / em + a * this.n));
  }
}

/**
 * 行尾放不下时怎么办（Word 实测，scripts 见 rules.md「标点压缩与悬挂」）：
 * - 兼容模式 15（Word 2013 及以后）：只有两端对齐 / 分散对齐的行能压缩标点、让行尾标点悬挂；
 * - 兼容模式 11–14：所有对齐方式都和两端对齐一样；
 * - 兼容模式 11–14 且兼容选项 useAltKinsokuLineBreakRules：行尾标点不悬挂（压缩标点时照样可以压到半宽放进本行）。
 */
function lineRules(align: string, setup: PageSetup): { squeeze: boolean; spaces: boolean; hang: boolean; legacy: boolean } {
  const legacy = (setup.compatMode ?? 15) < 15;
  const justified = align === 'both' || align === 'distribute';
  const squeeze = legacy || justified;
  // 空格和中西文间距的压窄只在两端对齐 / 分散对齐的行实测过（旧兼容模式的左对齐行 Word 不压它们）
  return { squeeze, spaces: justified, hang: squeeze && !(legacy && setup.altKinsoku), legacy };
}

function buildItems(p: Paragraph, env: Env, areaWidth: number): Item[] {
  const items: Item[] = [];
  let off = 0;
  // 显示域代码的域：开始标记显示 { 域代码 }，结果和结束标记不显示（还占着字符位置，宽度为 0）
  const codeAt = new Map<number, string>();
  const hidden = new Set<number>();
  if (p.runs.some((r) => r.fc))
    for (const f of fieldSpans(p)) {
      if (!showsCode(f.id) || hidden.has(f.begin)) continue;
      codeAt.set(f.begin, codeText(f.instr));
      for (let k = f.begin + 1; k <= f.end; k++) hidden.add(k);
    }
  for (let ri = 0; ri < p.runs.length; ri++) {
    const run = p.runs[ri];
    const pr = run.props;
    if (run.fc || hidden.has(ri)) {
      const code = codeAt.get(ri);
      if (code !== undefined) {
        let w = 0;
        for (const c of code) w += env.m.width(c, isCJK(c) ? pr.fontEA : pr.fontLatin, pr.bold, pr.italic) * glyphSize(pr);
        items.push({ ch: code, off, len: run.text.length, w, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'char' });
      } else if (run.fc || isAtom(run)) items.push({ ch: '', off, len: run.text.length, w: 0, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'char' });
      else {
        let k = off;
        for (const ch of run.text) {
          items.push({ ch: '', off: k, len: ch.length, w: 0, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'char' });
          k += ch.length;
        }
      }
      off += run.text.length;
      continue;
    }
    if (run.field) {
      const s = env.fieldText(run.field);
      let w = 0;
      for (const c of s) w += env.m.width(c, pr.fontLatin, pr.bold, pr.italic) * glyphSize(pr);
      items.push({ ch: s, off, len: 1, w, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'char' });
      off += run.text.length;
      continue;
    }
    if (run.box) {
      // 文本框的锚点：不占宽度，只记下它在哪一行
      items.push({ ch: '', off, len: run.text.length, w: 0, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'char', anchor: { box: run.box, areaW: areaWidth } });
      off += run.text.length;
      continue;
    }
    if (run.opaque?.note?.kind === 'foot') {
      // 脚注引用：显示编号（上标，字体按西文）
      const n = run.opaque.note.id === '@self' ? env.noteSelf : env.notes?.get(run.opaque.note.id);
      if (n !== undefined) {
        const s = String(n);
        const np = { ...pr, vertAlign: pr.vertAlign ?? ('superscript' as const) };
        let w = 0;
        for (const c of s) w += env.m.width(c, np.fontLatin, np.bold, np.italic) * glyphSize(np);
        items.push({ ch: s, off, len: run.text.length, w, gap: 0, props: np, family: np.fontLatin, cjk: false, kind: 'char', ...(run.opaque.note.id === '@self' ? {} : { note: run.opaque.note.id }) });
        off += run.text.length;
        continue;
      }
    }
    if (run.opaque?.kind === 'math') {
      // 公式：按 OMML 排版显示（原 XML 保存时原样写回）；单独成段的公式（m:oMathPara）居中
      const box = layoutMath(run.opaque.xml, pr.size, pr.fontEA, env.m);
      if (box) {
        items.push({ ch: FIELD_CHAR, off, len: run.text.length, w: box.w, gap: 0, props: pr, family: MATH_FONT, cjk: false, kind: 'char', image: { id: 'math', w: box.w, h: box.asc }, math: box });
        off += run.text.length;
        continue;
      }
    }
    if (run.opaque?.eq) {
      // EQ 域：各部分按自己的字号量宽度；重叠（\o）时宽度取最宽的一部分，否则相加
      const e = run.opaque.eq;
      const widths = e.parts.map((part) => part.reduce((a, seg) => a + [...seg.text].reduce((b, c) => b + env.m.width(c, isCJK(c) ? seg.props.fontEA : seg.props.fontLatin, seg.props.bold, seg.props.italic) * glyphSize(seg.props), 0), 0));
      const w = e.overstrike ? Math.max(0, ...widths) : widths.reduce((a, b) => a + b, 0);
      items.push({ ch: FIELD_CHAR, off, len: run.text.length, w, gap: 0, props: pr, family: pr.fontEA, cjk: true, kind: 'char', eq: { ...e, widths } });
      off += run.text.length;
      continue;
    }
    if (run.opaque) {
      // 原样对象：有大小的按图片排（有预览图就显示预览，没有画占位框），没有大小的（书签等）不占位置
      const o = run.opaque;
      if (o.w > 0) {
        const image = { id: o.image?.id ?? `opaque:${o.kind}`, w: o.w, h: o.h };
        items.push({ ch: FIELD_CHAR, off, len: run.text.length, w: o.w, gap: 0, props: pr, family: pr.fontLatin, cjk: true, kind: 'char', image });
      } else items.push({ ch: '', off, len: run.text.length, w: 0, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'char' });
      off += run.text.length;
      continue;
    }
    if (run.image) {
      // 比版心宽的图片按比例缩小到一行能放下
      const max = Math.max(10, areaWidth - p.props.indentLeft - p.props.indentRight);
      const k = Math.min(1, max / pictureBounds(run.image).w);
      const image = { ...run.image, w: run.image.w * k, h: run.image.h * k };
      items.push({ ch: FIELD_CHAR, off, len: run.text.length, w: pictureBounds(image).w, gap: 0, props: pr, family: pr.fontLatin, cjk: true, kind: 'char', image });
      off += run.text.length;
      continue;
    }
    for (const ch of run.text) {
      const len = ch.length;
      if (ch === '\t') {
        items.push({ ch, off, len, w: 0, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'tab' });
      } else if (ch === '\n' || ch === '\f' || ch === '\v' || ch === '\r') {
        items.push({ ch: ch === '\r' ? '\n' : ch, off, len, w: 0, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'break' });
      } else if (ch === FIELD_CHAR) {
        items.push({ ch: '', off, len, w: 0, gap: 0, props: pr, family: pr.fontLatin, cjk: false, kind: 'char' });
      } else {
        const cjk = isCJK(ch);
        const family = cjk ? pr.fontEA : pr.fontLatin;
        // 字符间距（w:spacing）加在每个字符后面；中文字符（全角）加两倍（Word 实测：12 磅宋体 / 微软雅黑，
        // 间距 −0.5 / +2 / +3 磅时汉字和全角逗号每字前进 −1 / +4 / +6，西文字母 −0.5 / +2 / +3）
        const w = gridWidth(env.m.width(ch, family, pr.bold, pr.italic) * glyphSize(pr), cjk, env, p.props.snapToGrid) + (pr.spacing ?? 0) * (cjk ? 2 : 1);
        items.push({ ch, off, len, w, gap: 0, props: pr, family, cjk, kind: 'char' });
      }
      off += len;
    }
  }
  // 中西文自动间距（Word 实测，与网格、字体、兼容模式无关）：中文与西文字母 / 数字之间加 1/4 字宽，两个方向一样；
  // 中文一侧是标点（：1、（A）等）或中间有空格时不加
  for (let i = 1; i < items.length; i++) {
    const a = items[i - 1];
    const b = items[i];
    const size = Math.min(a.props.size, b.props.size);
    if ((a.cjk && !isPunct(a.ch) && isLatinAlnum(b.ch)) || (isLatinAlnum(a.ch) && b.cjk && !isPunct(b.ch))) b.gap = size * 0.25;
  }
  // 相邻标点挤压（Word 实测，与压缩标点设置、兼容模式无关；文档设了 noPunctuationKerning 时不挤压）：
  // 结束类标点后面紧跟全角标点时它只占半宽（去掉右边的空白），两个开始类标点相连时后一个只占半宽（去掉左边的空白）
  for (let i = 1; i < items.length && !env.doc.setup.noPunctKern; i++) {
    const a = items[i - 1];
    const b = items[i];
    if (a.kind !== 'char' || b.kind !== 'char' || !fullPunct(a) || !fullPunct(b)) continue;
    if (PAIR_CLOSE.has(a.ch) && (PAIR_CLOSE.has(b.ch) || PUNCT_OPEN.has(b.ch))) a.w /= 2;
    else if (PUNCT_OPEN.has(a.ch) && PUNCT_OPEN.has(b.ch)) {
      b.trimL = b.w / 2;
      b.w /= 2;
    }
  }
  return items;
}

/** 标点、符号和空白（含全角空格；中西文自动间距不加在它们旁边，Word 实测） */
function isPunct(ch: string): boolean {
  return /^[\p{P}\p{S}\p{Z}]$/u.test(ch);
}

/** 参与相邻挤压的结束类标点（！？不参与，Word 实测） */
const PAIR_CLOSE = new Set([...'，。、；：）］】》〉」』〕｝”’']);

/** 全角宽的标点（字体里本来就是半角宽的不挤压） */
function fullPunct(it: Item): boolean {
  return (PUNCT_OPEN.has(it.ch) || PUNCT_CLOSE.has(it.ch)) && it.w >= 0.8 * glyphSize(it.props);
}

/** 行的自然高度：各字体上伸的最大值 + 下伸的最大值（Word 实测：TNR“10”与仿宋“钢板”同一行 = 1.117 + 0.216 em） */
function lineMetrics(fonts: LineFont[], env: Env): { natural: number; ascent: number } {
  let asc = 0;
  let desc = 0;
  for (const { pr, family } of fonts) {
    for (const f of family ? [family] : [pr.fontEA, pr.fontLatin]) {
      const m = env.m.metrics(f);
      asc = Math.max(asc, m.ascent * pr.size);
      desc = Math.max(desc, m.descent * pr.size);
    }
  }
  return { natural: asc + desc, ascent: asc };
}

/** 按行距规则计算行高与基线位置。imageH 为行内最高的图片（图片底边在基线上，把行撑高）。 */
function lineBox(p: Paragraph, props: LineFont[], env: Env, imageH = 0, extraDesc = 0): { height: number; baseline: number; fit?: number } {
  const m = lineMetrics(props, env);
  const ascent = Math.max(m.ascent, imageH);
  const natural = Math.max(m.natural - m.ascent, extraDesc) + ascent;
  const ls = p.props.line;
  const grid = env.lineGrid && p.props.snapToGrid;
  const pitch = env.doc.setup.grid.linePitch;
  const snapped = grid ? pitch * Math.max(1, Math.ceil(natural / pitch - EPS)) : natural;
  let height: number;
  if (ls.rule === 'exact') height = ls.value;
  else if (ls.rule === 'atLeast') height = Math.max(ls.value, snapped);
  // 多倍行距：有行网格时 = max(倍数 × 网格行距, 取整到网格的自然高度)（Word 实测：五号字 2.408 倍 = 37.6 磅，
  // 二号字 1.5 / 2 倍仍是两格 31.2 磅、3 倍 46.8 磅）；没有网格时 = 自然高度 × 倍数
  else height = grid ? Math.max(snapped, pitch * ls.value) : snapped * ls.value;
  // 文字在行内垂直居中。网格行在页底时 Word 只要求文字底部在版心内（实测：行距 29、三号字，
  // 页底剩 24 磅时西文行放得下、中文行放不下，正好是 (行距 + 文字自然高) / 2）
  const baseline = (height - natural) / 2 + ascent;
  return grid && ls.rule === 'auto' ? { height, baseline, fit: (height + natural) / 2 } : { height, baseline };
}

/** 每“行”的单位：网格行距或 1.2 倍默认字号。 */
function lineUnit(env: Env): number {
  return env.doc.setup.grid.type !== 'none' ? env.doc.setup.grid.linePitch : env.doc.defaults.size * 1.2;
}

const paraCache = new WeakMap<Paragraph, { key: string; result: ParaLayout }>();

const boxIds = new WeakMap<TextBox, number>();
let boxSeq = 0;
/** 文本框对象的编号（放进段落排版的缓存键：换了文本框对象就重排） */
function boxId(b: TextBox): number {
  let id = boxIds.get(b);
  if (!id) boxIds.set(b, (id = ++boxSeq));
  return id;
}

function propsKey(r: RunProps): string {
  return `${r.fontEA}|${r.fontLatin}|${r.size}|${r.bold ? 1 : 0}${r.italic ? 1 : 0}${r.underline ? 1 : 0}|${r.color}|${r.strike ? 1 : 0}${r.vertAlign ?? ''}${r.spacing ?? ''}${r.highlight ?? ''}${r.rev ? r.rev.kind + r.rev.author : ''}`;
}

/** 显示用的字号：上标 / 下标按 65% */
export function glyphSize(r: RunProps): number {
  return r.vertAlign ? r.size * 0.65 : r.size;
}

/** 带缓存的段落排版：内容、格式、宽度、网格或域值都没变时复用上次结果。 */
export function layoutParagraph(p: Paragraph, areaWidth: number, env: Env): ParaLayout {
  const g = env.doc.setup.grid;
  const hasField = p.runs.some((r) => r.field);
  const label = env.labels?.get(p) ?? '';
  const numLevel = p.props.num ? env.doc.lists.find((l) => l.id === p.props.num!.id)?.levels[p.props.num.level] : undefined;
  const key =
    `${env.m.epoch ?? 0}.${fontEpoch()}|${label}|${numLevel?.suffix ?? ''}|${numLevel?.label ? JSON.stringify(numLevel.label) : ''}|` +
    `${areaWidth}|${env.lineGrid ? 1 : 0}${env.charGrid ? 1 : 0}|${g.linePitch}|${g.charSpace}|${env.doc.defaults.size}|${env.doc.setup.defaultTab}|` +
    (hasField ? `${env.fieldText('PAGE')}/${env.fieldText('NUMPAGES')}|` : '') +
    (p.runs.some((r) => r.fc) ? `fc${fieldCodeEpoch()}|` : '') +
    JSON.stringify(p.props) +
    propsKey(p.mark) +
    p.runs.map((r) => `${r.text}\u0000${propsKey(r.props)}${r.field ?? ''}${r.fc ? `fc${r.fc.kind}${r.fc.id}${showsCode(r.fc.id) ? '#' : ''}${r.fc.instr ?? ''}` : ''}${r.image ? pictureKey(r.image) : ''}${r.box ? `box${boxId(r.box)}` : ''}${r.opaque ? `op${r.opaque.w}x${r.opaque.h}${r.opaque.image?.id ?? ''}${r.opaque.note ? `n${r.opaque.note.id === '@self' ? env.noteSelf : env.notes?.get(r.opaque.note.id)}` : ''}` : ''}`).join('\u0001');
  const hit = paraCache.get(p);
  if (hit && hit.key === key) return hit.result;
  const result = layoutParagraphUncached(p, areaWidth, env);
  paraCache.set(p, { key, result });
  return result;
}

function layoutParagraphUncached(p0: Paragraph, areaWidth: number, env: Env): ParaLayout {
  // 文本框架：水平按 xAlign 对齐（相对版心）；排成普通段落，放置时不占高度
  const fa = p0.props.frame?.xAlign;
  const p: Paragraph = fa ? { ...p0, props: { ...p0.props, align: fa === 'center' ? 'center' : fa === 'right' || fa === 'outside' ? 'right' : 'left' } } : p0;
  const items = buildItems(p, env, areaWidth);
  const pr = p.props;
  const firstSize = p.runs[0]?.props.size ?? p.mark.size;
  const charUnit = env.charGrid && pr.snapToGrid ? charPitch(env.doc) : firstSize;
  const firstIndent = pr.firstLineChars !== null ? pr.firstLineChars * charUnit : pr.firstLine;
  const tab = env.doc.setup.defaultTab;
  const lines: Line[] = [];
  const total = p.runs.reduce((n, r) => n + r.text.length, 0);

  // 自动编号：编号放在首行起点，正文从编号后（制表位 / 空格 / 紧接）开始
  const label = env.labels?.get(p) ?? '';
  const labelGlyphs: Glyph[] = [];
  let firstStart = pr.indentLeft + firstIndent;
  if (label) {
    const level = pr.num ? env.doc.lists.find((l) => l.id === pr.num!.id)?.levels[pr.num.level] : undefined;
    // 编号的格式：跟随正文，编号级别自己写了的（加粗、字号…）优先
    const lp = { ...(p.runs[0]?.props ?? p.mark), ...level?.label };
    let lx = firstStart;
    for (const ch of label) {
      const cjk = isCJK(ch);
      const family = cjk ? lp.fontEA : lp.fontLatin;
      const w = gridWidth(env.m.width(ch, family, lp.bold, lp.italic) * lp.size, cjk, env, pr.snapToGrid);
      labelGlyphs.push({ ch, x: lx, w, props: lp, family, cjk });
      lx += w;
    }
    const suffix = level?.suffix ?? 'tab';
    if (suffix === 'tab') firstStart = pr.indentLeft > lx + EPS ? pr.indentLeft : Math.ceil((lx + EPS) / tab) * tab;
    else if (suffix === 'space') firstStart = lx + env.m.width(' ', lp.fontLatin, lp.bold, lp.italic) * lp.size;
    else firstStart = lx;
  }

  let i = 0;
  let first = true;
  let lastBrk: Line['brk'] = 'end';
  for (;;) {
    if (i >= items.length) {
      if (lines.length === 0 || lastBrk === 'line' || lastBrk === 'page' || lastBrk === 'column') {
        const box = lineBox(p, [{ pr: p.mark }], env);
        const startX = first ? firstStart : pr.indentLeft;
        const x = alignOffset(pr.align, areaWidth - pr.indentRight - startX, 0) + startX;
        lines.push({ start: total, end: total, xs: [x], glyphs: [], right: x, height: box.height, baseline: box.baseline, ...(box.fit !== undefined ? { fit: box.fit } : {}), brk: 'end' });
      }
      break;
    }
    const startX = first ? firstStart : pr.indentLeft;
    const avail = areaWidth - pr.indentRight - startX;
    const widths: number[] = [];
    let x = 0;
    let cap = 0;
    const { squeeze, spaces, hang, legacy } = lineRules(pr.align, env.doc.setup);
    const punct = new PunctCap(legacy);
    /** 靠压缩挤进本行的单位到哪里为止（紧跟其后的悬挂标点照样悬挂） */
    let squeezedTo = -1;
    const compress = !!env.doc.setup.compressPunct;
    let j = i;
    let brk: Line['brk'] = 'end';
    for (; j < items.length; j++) {
      const it = items[j];
      if (it.kind === 'break') {
        widths.push(0);
        brk = it.ch === '\f' ? 'page' : it.ch === '\v' ? 'column' : 'line';
        j++;
        break;
      }
      const w = it.kind === 'tab' ? tabWidth(items, j, startX + x, pr.tabs, tab) : it.w + (j > i ? it.gap : 0);
      // 旧兼容模式：行内标点合计最多压一个半宽，行尾的结束标点自己的半宽另算（Word 实测）
      const own = squeeze ? punctShrink(it, compress) : 0;
      const legacyCap = punct.capFor(x, it.w) + (PUNCT_CLOSE.has(it.ch) ? own : 0);
      if (squeeze) {
        if (spaces) cap += shrinkable(it, j === i);
        punct.add(own);
      }
      if (x + w > avail + EPS + cap && j > i && it.ch !== ' ') {
        // 前面的字靠压缩放下时，行尾标点照样悬挂（Word 实测）
        const pcap = legacy ? legacyCap : punct.capFor(x, it.w);
        if (HANGING.has(it.ch) && hang && (x <= avail + EPS + cap + pcap || j === squeezedTo)) {
          // 标点溢出边界：行尾标点悬挂到边界外（哪些情况允许见 lineRules）；
          // 不允许时标点放不下就连同前一个字换到下一行
          widths.push(w);
          x += w;
          j++;
          brk = 'wrap';
          break;
        }
        let k = j;
        while (k > i && !canBreakBetween(items[k - 1].ch, items[k].ch)) k--;
        // 行末放不下的“单位”[k, m)（不可断开的一段：一个汉字连同它后面的行首禁则标点、一个西文单词或数字串如“100%”）：
        // 压缩前面的标点挤进本行，还是整个推到下一行（Word 按整个单位判断）
        if (k > i && squeeze && compress) {
          let m = j + 1;
          while (m < items.length && items[m].kind === 'char' && items[m].ch !== ' ' && !canBreakBetween(items[m - 1].ch, items[m].ch)) m++;
          // 单位末尾允许悬挂的标点悬挂到边界外，不算进单位（此时 Word 更倾向推出：a ≈ 21）
          const tailHangs = hang && m - 1 > j && HANGING.has(items[m - 1].ch);
          const end = tailHangs ? m - 1 : m;
          let before = 0;
          for (let q = 0; q < k - i; q++) before += widths[q];
          let unitW = 0;
          for (let q = k; q < end; q++) unitW += q < j ? widths[q - i] : q === j ? w : items[q].w + items[q].gap;
          const tail = items[end - 1];
          const tailOwn = tail && PUNCT_CLOSE.has(tail.ch) ? punctShrink(tail, compress) : 0;
          let sum = 0;
          let max = 0;
          for (let q = i; q < end; q++) {
            const v = punctShrink(items[q], compress);
            sum += v;
            // 旧兼容模式的“行内最多一个半宽”不含行尾那个标点（它自己的半宽另算）
            if (q < end - 1 || !tailOwn) max = Math.max(max, v);
          }
          const over = before + unitW - avail - cap;
          const room = avail - before;
          const fits =
            over <= EPS ||
            (legacy
              ? over <= max + tailOwn + EPS
              : over <= sum + EPS && over * before <= 2 * (tailHangs ? SQUEEZE_A_HANG : SQUEEZE_A) * sum * room + EPS);
          if (fits) {
            for (let q = j; q < end; q++) {
              const wq = q === j ? w : items[q].w + items[q].gap;
              if (q > j) punct.add(punctShrink(items[q], compress));
              widths.push(wq);
              x += wq;
            }
            squeezedTo = end;
            j = end - 1;
            continue;
          }
        }
        if (k === i) {
          // 找不到断点（超长单词）：强制断开，但仍避开行首禁则字符
          k = j;
          while (k > i + 1 && NO_LINE_START.has(items[k].ch)) k--;
        }
        widths.length = k - i;
        j = k;
        brk = 'wrap';
        break;
      }
      widths.push(w);
      x += w;
    }
    if (j >= items.length && brk === 'end' && widths.length === j - i) brk = 'end';
    lines.push(buildLine(p, items, i, j, widths, startX, avail, brk, env));
    lastBrk = brk;
    i = j;
    first = false;
    if (brk === 'end') break;
  }

  if (labelGlyphs.length && lines[0]) {
    // 居中/右对齐时编号随正文一起移动
    const shift = lines[0].xs[0] - firstStart;
    lines[0].glyphs.unshift(...labelGlyphs.map((g) => ({ ...g, x: g.x + (pr.align === 'center' || pr.align === 'right' ? shift : 0) })));
  }

  const unit = lineUnit(env);
  // 上下边框：线和它与文字的距离占段落的高度（边框画在段前段后之内）
  const bt = pr.borders?.top;
  const bb = pr.borders?.bottom;
  return {
    lines,
    // 段前 / 段后“自动”是 14 磅（Word 实测；与相邻段距一样取大）
    spaceBefore: (pr.spaceBeforeAuto ? AUTO_SPACING : pr.spaceBeforeLines !== null ? pr.spaceBeforeLines * unit : pr.spaceBefore) + (bt ? bt.w + bt.space : 0),
    spaceAfter: (pr.spaceAfterAuto ? AUTO_SPACING : pr.spaceAfterLines !== null ? pr.spaceAfterLines * unit : pr.spaceAfter) + (bb ? bb.w + bb.space : 0),
  };
}

function alignOffset(align: string, avail: number, width: number): number {
  const extra = avail - width;
  if (align === 'center') return extra / 2;
  if (align === 'right') return extra;
  return 0;
}

function buildLine(
  p: Paragraph,
  items: Item[],
  i: number,
  j: number,
  widths: number[],
  startX: number,
  avail: number,
  brk: Line['brk'],
  env: Env,
): Line {
  const seg = items.slice(i, j);
  // 可见内容末尾（去掉行尾空格与换行符）
  let last = seg.length - 1;
  while (last >= 0 && (seg[last].ch === ' ' || seg[last].kind === 'break')) last--;
  let contentW = 0;
  for (let k = 0; k <= last; k++) contentW += widths[k];

  // 对齐；只有一个单独成段的公式（m:oMathPara，Word 缺省居中）的行按居中
  const displayMath = seg.some((it) => it.math) && seg.every((it) => it.math || it.kind !== 'char' || it.ch === '' ) && p.runs.some((r) => r.opaque?.kind === 'math' && r.opaque.level === 'para' && /^<m:oMathPara[\s>]/.test(r.opaque.xml));
  const align = displayMath ? 'center' : p.props.align;
  let offset = 0;
  let perGap = 0;
  // 以手动换行符结束的行：文档有 doNotExpandShiftReturn 时不拉开（Word 按兼容选项），否则和自动换行的行一样撑满
  const justify = (align === 'both' && (brk === 'wrap' || (brk === 'line' && !env.doc.setup.noExpandShiftReturn))) || align === 'distribute';
  const gaps: boolean[] = seg.map((it, k) => k > 0 && k <= last && (it.cjk || seg[k - 1].cjk || seg[k - 1].ch === ' '));
  const nGaps = gaps.filter(Boolean).length;
  // 压窄（见 SHRINK）：按各自可压的量等比例压
  const cut: number[] = [];
  /** 开始标点压掉的左半边：字形往左移 */
  const shift: number[] = [];
  const compress = !!env.doc.setup.compressPunct;
  const rules = lineRules(align, env.doc.setup);
  if (rules.squeeze && contentW - avail > EPS) {
    // 行尾的结束标点先去掉自己右边的空白（Word：行尾逗号剩余 ≥ 半字宽时压成半宽留在本行）；
    // 允许悬挂的标点再放不下就悬挂到边界外，不去挤压行内其它字
    let lastAdj = 0;
    const tail = seg[last];
    if (tail && PUNCT_CLOSE.has(tail.ch)) {
      lastAdj = Math.min(contentW - avail, punctShrink(tail, compress));
      widths[last] -= lastAdj;
      contentW -= lastAdj;
    }
    const hangs = tail && rules.hang && HANGING.has(tail.ch);
    // 剩下的：先压行内标点，还不够再按比例压空格和中西文间距
    const upto = lastAdj > 0 || hangs ? last - 1 : last;
    const pc = new PunctCap(rules.legacy);
    let total = 0;
    for (let k = 0; k <= upto; k++) {
      if (rules.spaces) total += shrinkable(seg[k], k === 0);
      pc.add(punctShrink(seg[k], compress));
    }
    const need = Math.max(0, (hangs ? contentW - widths[last] : contentW) - avail);
    const fp = pc.sum > 0 ? Math.min(need, rules.legacy ? pc.max : pc.sum) / pc.sum : 0;
    const f = total > 0 ? Math.min(1, Math.max(0, need - pc.sum * fp) / total) : 0;
    for (let k = 0; k <= upto; k++) {
      const ps = fp * punctShrink(seg[k], compress);
      const s = f * shrinkable(seg[k], k === 0) + ps;
      cut[k] = f * SHRINK * (k > 0 && seg[k].kind === 'char' ? seg[k].gap : 0);
      if (PUNCT_OPEN.has(seg[k].ch)) shift[k] = ps;
      widths[k] -= s;
      contentW -= s;
    }
  }
  if (justify && avail - contentW > EPS && nGaps > 0) perGap = (avail - contentW) / nGaps;
  else if (align === 'distribute' && last === 0) offset = (avail - contentW) / 2;
  else if (align === 'center' || align === 'right') offset = alignOffset(align, avail, contentW);

  const start = items[i]?.off ?? 0;
  const end = j < items.length ? items[j].off : items.length ? items[items.length - 1].off + items[items.length - 1].len : 0;
  const xs = new Array<number>(end - start + 1).fill(0);
  const glyphs: Glyph[] = [];
  let x = startX + offset;
  let right = x;
  seg.forEach((it, k) => {
    if (gaps[k]) x += perGap;
    const gapBefore = k > 0 ? it.gap - (cut[k] ?? 0) : 0;
    const cx = x;
    for (let u = 0; u < it.len; u++) xs[it.off - start + u] = cx;
    if (it.kind === 'tab' && it.leader && k <= last) {
      glyphs.push({ ch: ' ', x: cx, w: widths[k] ?? 0, props: it.props, family: it.family, cjk: false, leader: it.leader });
    } else if (it.props.underline && k <= last && (it.ch === ' ' || it.kind === 'tab')) {
      // 带下划线的空格 / 制表符：没有字形，但下划线要画（行尾的空格不画，与 Word 一致）
      glyphs.push({ ch: ' ', x: cx, w: widths[k] ?? 0, props: it.props, family: it.family, cjk: false });
    } else if (it.kind === 'char' && it.ch && it.ch !== ' ') {
      glyphs.push({ ch: it.ch, offset: it.off, x: cx + (it.kind === 'char' ? gapBefore : 0) - (shift[k] ?? 0) - (it.trimL ?? 0), w: it.w, props: it.props, family: it.family, cjk: it.cjk, ...(it.image ? { image: it.image } : {}), ...(it.math ? { math: it.math } : {}), ...(it.eq ? { eq: it.eq } : {}) });
    }
    x += widths[k] ?? 0;
    if (k <= last) right = x;
  });
  xs[end - start] = x;
  // 空格不撑高行（Word 实测：行中、行尾的 20 磅空格都不改变五号字行的行距，换成 20 磅字母才变）
  const props: LineFont[] = seg.filter((it) => it.kind === 'char' && it.ch !== '' && it.ch !== ' ').map((it) => ({ pr: it.props, family: it.image ? undefined : it.family }));
  const imageH = seg.reduce((h, it) => Math.max(h, it.image ? pictureBounds(it.image).h : 0), 0);
  const mathDesc = seg.reduce((h, it) => Math.max(h, it.math?.desc ?? 0), 0);
  const box = lineBox(p, props.length ? props : [{ pr: p.mark }], env, imageH, mathDesc);
  const anchors = seg.flatMap((it) => (it.anchor ? [it.anchor] : []));
  const notes = seg.flatMap((it) => (it.note ? [it.note] : []));
  return { start, end, xs, glyphs, right, height: box.height, baseline: box.baseline, ...(box.fit !== undefined ? { fit: box.fit } : {}), brk, ...(anchors.length ? { anchors } : {}), ...(notes.length ? { notes } : {}) };
}

// ---------------- 文档排版 ----------------

interface RegionFrom {
  bi: number;
  line: number;
  lines?: Line[];
  width?: number;
}

/** 文字要绕开的区域（页面坐标，已加上环绕距离）；tb = 上下型，整行都不能放 */
interface Excl {
  x0: number;
  x1: number;
  y0: number;
  y1: number;
  tb: boolean;
}

/** 四周型环绕时，旁边剩下的宽度小于这个值就不放文字，直接排到浮动对象下面 */
const MIN_WRAP = 18;

interface FlowMark {
  excl: number;
  body: number;
  fills: number;
  borders: number;
  cells: number;
  y: number;
  col: number;
  hasContent: boolean;
  regionBottom: number;
  regionFrom: RegionFrom | null;
  bottom: number;
}

class Flow {
  pages: PageLayout[] = [];
  page!: PageLayout;
  y = 0;
  hasContent = false;
  /** 上一段的段后距：下一段的段前距只补足比它多的部分（Word 实测：两段之间取段后、段前中较大的一个） */
  lastAfter = 0;
  /** 当前页正文区域的上下边界 */
  top = 0;
  bottom = 0;
  /** 当前所在的节；新页属于它 */
  section = 0;
  /** 每节的首页（sectionFirst[i] 为第 i 页是否从一节开头开始） */
  sectionFirst: boolean[] = [];
  /** 分栏：当前栏、本页这一段分栏区域的顶部（连续分节后新节从页中间开始）、区域内容在 body 里的起点 */
  col = 0;
  regionTop = 0;
  regionStart = 0;
  regionFills = 0;
  regionBorders = 0;
  regionCells = 0;
  /** 区域里各栏排到的最低位置（画分隔线、平衡栏用） */
  regionBottom = 0;
  regionClosed = true;
  /** 本页分栏区域从哪里开始排（块下标、段内第几行和那时的行），平衡分栏时从这里重排；line < 0 为表格中间，不能重排 */
  regionFrom: RegionFrom | null = null;
  regionExcl = 0;
  /** 本页文字要绕开的区域（浮动图片、文本框、浮动表格） */
  excl: Excl[] = [];
  /** 排版时算出的浮动对象位置（排完正文后放置文本框 / 图片时用同一个位置） */
  floatPos = new Map<TextBox, { x: number; y: number }>();
  /** 本页页底的脚注区：引用到的脚注、版心底边（没有脚注时的 bottom）；noteHeight 算一组脚注占的高度 */
  noteIds: string[] = [];
  pageBottom = 0;
  noteHeight: (ids: string[]) => number = () => 0;
  /** 试排：排不下时只记 overflow，不换页 */
  trial = false;
  overflow = false;
  /**
   * 页眉 / 页脚比页边距高时正文让开（Word：正文顶 = max(上边距, 页眉距边界 + 页眉高)，底同理）；
   * 参数是新页的序号和它是不是一节的首页
   */
  constructor(
    public sections: SectionInfo[],
    public hfBounds: ((page: number, sectionStart: boolean, section: number) => { top: number; bottom: number }) | null = null,
  ) {
    this.newPage(true);
  }
  /** 当前正在排的内容所在节的页面设置（连续分节后可与本页的不同） */
  get setup(): PageSetup {
    return this.sections[Math.min(this.section, this.sections.length - 1)].setup;
  }
  get columns(): { x: number; w: number }[] {
    return columnBoxes(this.setup);
  }
  /** 当前栏的左边（页面坐标）和宽度 */
  get colX(): number {
    const cs = this.columns;
    return this.setup.marginLeft + cs[Math.min(this.col, cs.length - 1)].x;
  }
  get colW(): number {
    const cs = this.columns;
    return cs[Math.min(this.col, cs.length - 1)].w;
  }
  newPage(sectionStart = false) {
    if (this.trial) {
      this.overflow = true;
      return;
    }
    this.closeRegion();
    const setup = this.setup;
    this.page = { index: this.pages.length, section: this.section, setup, hf: { header: 'header', footer: 'footer', owner: null }, body: [], header: [], footer: [], borders: [], fills: [], floats: [] };
    this.pages.push(this.page);
    this.sectionFirst.push(sectionStart);
    this.top = setup.marginTop;
    this.bottom = setup.height - setup.marginBottom;
    const hb = this.hfBounds?.(this.pages.length - 1, sectionStart, this.section);
    if (hb) {
      this.top = Math.max(this.top, hb.top);
      this.bottom = Math.min(this.bottom, hb.bottom);
    }
    this.y = this.top;
    this.hasContent = false;
    this.lastAfter = 0;
    this.excl = [];
    this.noteIds = [];
    this.pageBottom = this.bottom;
    this.openRegion(this.top);
  }
  /** 本栏排满：换到下一栏，最后一栏换页 */
  nextColumn() {
    if (this.col + 1 < this.columns.length) {
      this.regionBottom = Math.max(this.regionBottom, this.y);
      this.col++;
      this.y = this.regionTop;
      this.hasContent = false;
    this.lastAfter = 0;
      return;
    }
    this.newPage();
  }
  /**
   * 当前栏里 [y, y + h] 这一行能放文字的范围（页面坐标）：没有浮动对象时是整栏；
   * 四周型取浮动对象旁边最宽的一段，太窄或遇到上下型时返回 skip = 应该移到的 y。
   */
  freeSpan(y: number, h: number): { x: number; w: number } | { skip: number } {
    const x0 = this.colX;
    const x1 = x0 + this.colW;
    const hits = this.excl.filter((e) => e.y0 < y + h - EPS && e.y1 > y + EPS && e.x0 < x1 - EPS && e.x1 > x0 + EPS);
    if (!hits.length) return { x: x0, w: this.colW };
    const tb = hits.filter((e) => e.tb);
    if (tb.length) return { skip: Math.max(...tb.map((e) => e.y1)) };
    let ivs: [number, number][] = [[x0, x1]];
    for (const e of hits) ivs = ivs.flatMap(([a, b]): [number, number][] => [[a, Math.min(b, e.x0)], [Math.max(a, e.x1), b]]).filter(([a, b]) => b - a > EPS);
    const best = ivs.sort((p, q) => q[1] - q[0] - (p[1] - p[0]))[0];
    if (!best || best[1] - best[0] < MIN_WRAP) return { skip: Math.min(...hits.map((e) => e.y1)) };
    return { x: best[0], w: best[1] - best[0] };
  }
  /** 浮动对象占的区域加入环绕（wrap 为 none 的不影响文字） */
  addExclusion(x: number, y: number, w: number, h: number, wrap: TextBox['wrap'], dist?: TextBox['dist']) {
    if (wrap === 'none') return;
    const d = dist ?? { t: 0, b: 0, l: 0, r: 0 };
    this.excl.push({ x0: x - d.l, x1: x + w + d.r, y0: y - d.t, y1: y + h + d.b, tb: wrap === 'topAndBottom' });
  }
  /** 加上一行引用的脚注后版心的底边（脚注区从页底往上占） */
  bottomWithNotes(ids: string[] | undefined): number {
    if (!ids?.length) return this.bottom;
    const all = [...this.noteIds, ...ids.filter((id) => !this.noteIds.includes(id))];
    return this.pageBottom - this.noteHeight(all);
  }
  addNotes(ids: string[] | undefined) {
    if (!ids?.length) return;
    for (const id of ids) if (!this.noteIds.includes(id)) this.noteIds.push(id);
    const area = this.noteHeight(this.noteIds);
    this.bottom = this.pageBottom - area;
    this.page.noteIds = [...this.noteIds];
    this.page.noteArea = area;
  }
  /** 下一栏的宽度与 width 不同（不等宽分栏、换页后回到第一栏）时需要按新宽度重排 */
  widthChanged(width: number): boolean {
    return Math.abs(this.colW - width) > EPS;
  }
  openRegion(top: number) {
    this.col = 0;
    this.regionTop = top;
    this.regionBottom = top;
    this.regionStart = this.page.body.length;
    this.regionFills = this.page.fills.length;
    this.regionBorders = this.page.borders.length;
    this.regionCells = this.page.cells?.length ?? 0;
    this.regionExcl = this.excl.length;
    this.regionClosed = false;
    this.regionFrom = null;
  }
  /** 结束本页的分栏区域：栏间分隔线画到最长一栏的底部 */
  closeRegion() {
    if (!this.page || this.regionClosed) return;
    this.regionClosed = true;
    const c = this.setup.cols;
    const bottom = Math.max(this.regionBottom, this.y);
    if (c?.sep && c.num > 1 && bottom > this.regionTop + EPS) {
      const cs = this.columns;
      const used = this.page.body.slice(this.regionStart);
      for (let i = 1; i < cs.length; i++) {
        // 只在后一栏有内容时画（Word 的分隔线只画在有文字的栏之间）
        const x0 = this.setup.marginLeft + cs[i].x;
        if (!used.some((l) => l.x >= x0 - EPS)) break;
        const x = this.setup.marginLeft + (cs[i - 1].x + cs[i - 1].w + cs[i].x) / 2;
        this.page.borders.push({ x1: x, y1: this.regionTop, x2: x, y2: bottom, w: 0.5 });
      }
    }
  }
  /** 本页分栏区域刚开始时的状态（平衡分栏时从这里试排） */
  regionMark(): FlowMark {
    return {
      excl: this.regionExcl,
      body: this.regionStart,
      fills: this.regionFills,
      borders: this.regionBorders,
      cells: this.regionCells,
      y: this.regionTop,
      col: 0,
      hasContent: this.regionTop > this.top + EPS,
      regionBottom: this.regionTop,
      regionFrom: this.regionFrom,
      bottom: this.bottom,
    };
  }
  restore(m: FlowMark) {
    const pg = this.page;
    pg.body.length = m.body;
    this.excl.length = m.excl;
    pg.fills.length = m.fills;
    pg.borders.length = m.borders;
    if (pg.cells) pg.cells.length = m.cells;
    this.y = m.y;
    this.col = m.col;
    this.hasContent = m.hasContent;
    this.regionBottom = m.regionBottom;
    this.regionFrom = m.regionFrom;
    this.bottom = m.bottom;
    this.overflow = false;
  }

}

/**
 * 字符网格下一个字符的宽度（Word 实测，OOXML 的 w:docGrid）：
 * - linesAndChars（“指定行和字符网格”）：每个字符（中西文都一样）在自然宽度上加 charSpace / 4096 磅，
 *   与字号无关——二号字不会占两格，小四号字也不会撑满一格；
 * - snapToChars（“文字对齐字符网格”）：汉字对齐到整数个网格（网格略窄于字号、不超过 25% 时仍占一格）。
 */
function gridWidth(w: number, cjk: boolean, env: Env, snap: boolean): number {
  if (!env.charGrid || !snap) return w;
  const g = env.doc.setup.grid;
  if (g.type === 'linesAndChars') return w + g.charSpace / 4096;
  const pitch = charPitch(env.doc);
  return cjk ? pitch * Math.max(1, Math.ceil(w / pitch - 0.25)) : w;
}

function makeEnv(doc: Doc, m: Measurer, grid: boolean, fieldText: Env['fieldText']): Env {
  const t = doc.setup.grid.type;
  return {
    m,
    doc,
    lineGrid: grid && t !== 'none',
    charGrid: grid && (t === 'linesAndChars' || t === 'snapToChars'),
    fieldText,
  };
}

/**
 * 一行能放文字的范围：按“行顶到文字底部”判断是否碰到浮动对象（Word 实测：网格行文字下面空出的部分
 * 伸进浮动表格上边几磅时仍排整行；行顶在表格下边以上、文字没碰到表格时也绕排）
 */
function textSpan(flow: Flow, y: number, l: Line): ReturnType<Flow['freeSpan']> {
  return flow.freeSpan(y, l.fit ?? l.height);
}

/**
 * 段落边框的线段（页面坐标）：top 为第一行顶、bottom 为最后一行底；左右从段落缩进处到右缩进处。
 * 线在文字外 space 磅处，与 Word 一样画在段落区域内。
 */
function paraBorderSegs(p: Paragraph, x: number, areaW: number, top: number, bottom: number): Seg[] {
  const b = p.props.borders;
  if (!b) return [];
  const x1 = x + p.props.indentLeft - (b.left ? b.left.space + b.left.w / 2 : 0);
  const x2 = x + areaW - p.props.indentRight + (b.right ? b.right.space + b.right.w / 2 : 0);
  const y1 = top - (b.top ? b.top.space + b.top.w / 2 : 0);
  const y2 = bottom + (b.bottom ? b.bottom.space + b.bottom.w / 2 : 0);
  const out: Seg[] = [];
  if (b.top) out.push({ x1, y1, x2, y2: y1, w: b.top.w });
  if (b.bottom) out.push({ x1, y1: y2, x2, y2, w: b.bottom.w });
  if (b.left) out.push({ x1, y1, x2: x1, y2, w: b.left.w });
  if (b.right) out.push({ x1: x2, y1, x2, y2, w: b.right.w });
  return out;
}

/**
 * 页眉页脚里的表格（只读）：从 (x, y) 起排，不分页；返回高度。行进 out（标为只读），框线和底纹进本页。
 */
function placeStaticTable(t: Table, x: number, y: number, areaW: number, env: Env, page: PageLayout, out: PlacedLine[], story: Story, path: number[]): number {
  const gridW = t.cols.reduce((a, b) => a + b, 0);
  const k = t.widthPct && gridW > 0 ? (t.widthPct * areaW) / gridW : 1;
  const tw = gridW * k;
  const left = x + (t.align === 'center' ? (areaW - tw) / 2 : t.align === 'right' ? areaW - tw : t.indent);
  const colX = [0];
  for (const w of t.cols) colX.push(colX[colX.length - 1] + w * k);
  const cm = t.cellMargin;
  const borderW = t.borders ? (t.borderSize ?? 0.5) : 0;
  const cellEnv = { ...env, lineGrid: false };
  let yy = y;
  for (const row of t.rows) {
    let g = 0;
    const boxes = row.cells.map((cell) => {
      const g0 = g;
      g += cell.gridSpan;
      const cx = left + colX[g0];
      const cx2 = left + colX[Math.min(g, t.cols.length)];
      let h = 0;
      const lines: { line: Line; dy: number }[] = [];
      for (const p of cell.paragraphs) {
        const pl = layoutParagraph(p, cx2 - cx - cm.left - cm.right, cellEnv);
        h += pl.spaceBefore;
        for (const line of pl.lines) {
          lines.push({ line, dy: h });
          h += line.height;
        }
        h += pl.spaceAfter;
      }
      return { cell, cx, cx2, h, lines };
    });
    const content = Math.max(0, ...boxes.map((b) => b.h)) + cm.top + cm.bottom + borderW;
    const rh = row.heightRule === 'exact' && row.height > 0 ? row.height : row.heightRule === 'atLeast' ? Math.max(row.height, content) : content;
    for (const b of boxes) {
      if (b.cell.shading) page.fills.push({ x: b.cx, y: yy, w: b.cx2 - b.cx, h: rh, color: `#${b.cell.shading}` });
      const free = rh - cm.top - cm.bottom - borderW - b.h;
      const dy = b.cell.vAlign === 'center' ? free / 2 : b.cell.vAlign === 'bottom' ? free : 0;
      for (const l of b.lines) out.push({ story, path, line: l.line, x: b.cx + cm.left, y: yy + borderW + cm.top + dy + l.dy, ro: true });
      const e = b.cell.edges;
      const col = (c: CellEdge) => (c.color && c.color !== 'auto' ? `#${c.color}` : undefined);
      if (e) {
        if (e.top) page.borders.push({ x1: b.cx, y1: yy, x2: b.cx2, y2: yy, w: e.top.w, color: col(e.top), ...(e.top.style && { style: e.top.style }) });
        if (e.bottom) page.borders.push({ x1: b.cx, y1: yy + rh, x2: b.cx2, y2: yy + rh, w: e.bottom.w, color: col(e.bottom), ...(e.bottom.style && { style: e.bottom.style }) });
        if (e.left) page.borders.push({ x1: b.cx, y1: yy, x2: b.cx, y2: yy + rh, w: e.left.w, color: col(e.left), ...(e.left.style && { style: e.left.style }) });
        if (e.right) page.borders.push({ x1: b.cx2, y1: yy, x2: b.cx2, y2: yy + rh, w: e.right.w, color: col(e.right), ...(e.right.style && { style: e.right.style }) });
      } else if (b.cell.borders ?? t.borders) {
        const w = t.borderSize ?? 0.5;
        page.borders.push({ x1: b.cx, y1: yy, x2: b.cx2, y2: yy, w }, { x1: b.cx, y1: yy + rh, x2: b.cx2, y2: yy + rh, w });
        page.borders.push({ x1: b.cx, y1: yy, x2: b.cx, y2: yy + rh, w }, { x1: b.cx2, y1: yy, x2: b.cx2, y2: yy + rh, w });
      }
    }
    yy += rh;
  }
  return yy - y;
}

/** 段落从 off 开始的部分（换到宽度不同的栏时重排用）：首行缩进、段前距、编号不再有 */
function sliceParagraph(p: Paragraph, off: number): Paragraph {
  const runs: Run[] = [];
  let pos = 0;
  for (const r of p.runs) {
    const end = pos + r.text.length;
    if (end > off) runs.push(pos >= off ? r : { ...r, text: r.text.slice(off - pos) });
    pos = end;
  }
  return { ...p, props: { ...p.props, firstLine: 0, firstLineChars: null, spaceBefore: 0, spaceBeforeLines: null }, runs };
}

function relayoutFrom(p: Paragraph, off: number, width: number, env: Env): Line[] {
  return layoutParagraph(sliceParagraph(p, off), width, env).lines.map((l) => ({ ...l, start: l.start + off, end: l.end + off }));
}

/**
 * 把段落的行放进版面（从第 from 行开始；from > 0 时为平衡分栏时从区域起点重排）。
 * 换到宽度不同的栏时，剩下的部分按新栏宽重新断行。
 */
function placeParagraph(pl: ParaLayout, p: Paragraph, path: number[], flow: Flow, nextFirstLine: number, env: Env, from = 0, width = flow.colW) {
  let lines = pl.lines;
  const pageHeight = flow.bottom - flow.top;
  const relayout = (k: number, w: number) => {
    width = w;
    lines = k === 0 ? layoutParagraph(p, w, env).lines : [...lines.slice(0, k), ...relayoutFrom(p, lines[k].start, w, env)];
  };
  const breakTo = (k: number, page: boolean) => {
    if (page) flow.newPage();
    else flow.nextColumn();
    if (flow.overflow) return false;
    if (k < lines.length && flow.widthChanged(width)) relayout(k, flow.colW);
    return true;
  };
  // 第 k 行放在当前 y 时能用的范围：绕开浮动对象（宽度变了就从这一行起重新断行），放不下就往下移 / 换栏
  let spanX = flow.colX;
  const placeSpan = (k: number): boolean => {
    for (let guard = 0; guard < 100; guard++) {
      const fs = flow.excl.length ? textSpan(flow, flow.y, lines[k]) : { x: flow.colX, w: flow.colW };
      if ('skip' in fs) {
        flow.y = Math.max(flow.y + EPS, fs.skip);
        if (flow.y + (lines[k].fit ?? lines[k].height) > flow.bottom + EPS && !breakTo(k, false)) return false;
        continue;
      }
      if (Math.abs(fs.w - width) > EPS) relayout(k, fs.w);
      spanX = fs.x;
      return true;
    }
    return true;
  };
  if (from === 0) {
    if (p.props.pageBreakBefore && flow.hasContent && !breakTo(0, true)) return;
    // 自动段前在页面（文档）开头不加（Word 实测）
    const before = p.props.spaceBeforeAuto && !flow.hasContent ? 0 : pl.spaceBefore;
    flow.y += Math.max(0, before - flow.lastAfter);
    const totalH = lines.reduce((s, l) => s + l.height, 0);
    if (p.props.keepNext && flow.hasContent && totalH + pl.spaceAfter + nextFirstLine <= pageHeight && flow.y + totalH + pl.spaceAfter + nextFirstLine > flow.bottom + EPS) {
      if (!breakTo(0, false)) return;
    } else if (p.props.keepLines && flow.hasContent && totalH <= pageHeight && flow.y + totalH > flow.bottom + EPS) {
      if (!breakTo(0, false)) return;
    }
  }
  if (from === 0) {
    // 段落里锚定的浮动对象：先定位，加入环绕，再排这一段的文字
    const paraTop = flow.y;
    let ly = paraTop;
    for (const l of lines) {
      for (const a of l.anchors ?? []) {
        const b = a.box;
        const pos = floatXY(b, flow.setup, flow.colX, flow.colW, paraTop, ly, l.height);
        flow.floatPos.set(b, pos);
        const bounds = b.image ? pictureBounds(b.image) : { w: b.w, h: b.h };
        flow.addExclusion(pos.x + (b.w - bounds.w) / 2, pos.y + (b.h - bounds.h) / 2, bounds.w, bounds.h, b.behind ? 'none' : b.wrap, b.dist);
      }
      ly += l.height;
    }
  }
  const frameY = p.props.frame ? flow.y - Math.max(0, (p.props.spaceBeforeAuto && !flow.hasContent ? 0 : pl.spaceBefore) - flow.lastAfter) : null;
  const firstY = flow.y;
  const firstPage = flow.page;
  let k = from;
  while (k < lines.length) {
    if (!placeSpan(k)) return;
    let fit = 0;
    let yy = flow.y;
    let pageBreak = false;
    let colBreak = false;
    let wrapStop = false;
    while (k + fit < lines.length && yy + (lines[k + fit].fit ?? lines[k + fit].height) <= flow.bottomWithNotes(lines[k + fit].notes) + EPS) {
      if (fit > 0 && flow.excl.length) {
        const fs = textSpan(flow, yy, lines[k + fit]);
        if ('skip' in fs || Math.abs(fs.x - spanX) > EPS || Math.abs(fs.w - width) > EPS) {
          wrapStop = true;
          break;
        }
      }
      yy += lines[k + fit].height;
      fit++;
      if (lines[k + fit - 1].brk === 'page') {
        pageBreak = true;
        break;
      }
      if (lines[k + fit - 1].brk === 'column') {
        colBreak = true;
        break;
      }
    }
    const rest = lines.length - k;
    if (fit < rest && !pageBreak && !colBreak && !wrapStop && p.props.widowControl && lines.length >= 2) {
      if (k === 0 && fit === 1) fit = 0; // 段首孤行
      else if (rest - fit === 1 && fit >= 2) fit -= 1; // 段末孤行
    }
    if (fit === 0 && !flow.hasContent) fit = 1; // 空页也放不下：强制放一行，避免死循环
    if (fit > 0) flow.regionFrom ??= { bi: path[0], line: k, lines, width };
    for (let q = 0; q < fit; q++) {
      const line = lines[k + q];
      flow.page.body.push({ story: 'body', path, line, x: spanX, y: flow.y });
      flow.addNotes(line.notes);
      flow.y += line.height;
      flow.hasContent = true;
    }
    k += fit;
    if (wrapStop && fit > 0) continue;
    if (pageBreak || colBreak || k < lines.length) {
      if (!breakTo(k, pageBreak)) return;
    }
  }
  if (p.props.borders && flow.page === firstPage && from === 0) flow.page.borders.push(...paraBorderSegs(p, flow.colX, width, firstY, flow.y));
  flow.y += pl.spaceAfter;
  flow.lastAfter = pl.spaceAfter;
  // 文本框架：段落浮在原处，后面的内容接着框架之前的位置排
  if (frameY !== null && flow.page === firstPage) {
    flow.y = frameY;
    flow.lastAfter = 0;
  }
}

/** 分节符所在段落排完后：换页或（连续分节）在本页接着排下一节 */
function endSection(flow: Flow, sb: SectionBreak, bi: number, balance: () => void) {
  const mark = () => {
    for (let i = flow.pages.length - 1; i >= 0; i--) {
      const l = [...flow.pages[i].body].reverse().find((x) => x.path.length === 1 && x.path[0] === bi);
      if (l) return void (l.sectionBreak = sb);
    }
  };
  // 连续分节但下一节纸张大小不同时，只能从新页开始
  const cur = flow.setup;
  const next = flow.sections[flow.section + 1]?.setup ?? cur;
  if (sb !== 'continuous' || next.width !== cur.width || next.height !== cur.height) {
    mark();
    flow.closeRegion();
    // 偶数页 / 奇数页分节：页码奇偶不符时空出一页（空白页仍属于上一节）
    const pageNo = flow.pages.length + 1;
    if ((sb === 'evenPage' && pageNo % 2 === 1) || (sb === 'oddPage' && pageNo % 2 === 0)) flow.newPage();
    flow.section++;
    flow.newPage(true);
    return;
  }
  // 连续分节：本页的页眉页脚仍由上一节决定，新节在本页中间开始，没有“本节首页”；
  // 分栏的一节先平衡各栏，新节从最长一栏的下面开始
  const multi = (cur.cols?.num ?? 1) > 1 || (next.cols?.num ?? 1) > 1;
  if ((cur.cols?.num ?? 1) > 1) balance();
  mark();
  const bottom = Math.max(flow.regionBottom, flow.y);
  flow.closeRegion();
  flow.section++;
  if (multi) {
    flow.y = bottom;
    flow.openRegion(bottom);
  }
}

/** 只含分节符的空段落（连续分节）：Word 里不占高度（实测：前后两节的行在网格上紧挨着） */
function isBareContinuousBreak(p: Paragraph): boolean {
  return p.props.sectionBreak === 'continuous' && p.runs.every((r) => !r.text);
}

interface CellBox {
  lines: { path: number[]; line: Line; dy: number }[];
  height: number;
}

function layoutCell(table: Table, bi: number, r: number, ci: number, width: number, env: Env): CellBox {
  const cell = table.rows[r].cells[ci];
  const out: CellBox = { lines: [], height: 0 };
  if (cell.vMerge === 'continue') return out;
  let y = 0;
  let lastAfter = 0;
  // Word 实测：表格单元格里的行只在有兼容选项 adjustLineHeightInTable 时对齐行网格，否则行高按字体算；
  // 字符网格的字距调整照常
  const cellEnv = env.lineGrid && !env.doc.setup.tableGrid ? { ...env, lineGrid: false } : env;
  cell.paragraphs.forEach((p, pi) => {
    const pl = layoutParagraph(p, width, cellEnv);
    // 段前、段后相邻时取较大的一个（同正文）；单元格第一段的自动段前不加
    y += Math.max(0, (pi === 0 && p.props.spaceBeforeAuto ? 0 : pl.spaceBefore) - lastAfter);
    for (const line of pl.lines) {
      out.lines.push({ path: [bi, r, ci, pi], line, dy: y });
      y += line.height;
    }
    y += pl.spaceAfter;
    lastAfter = pl.spaceAfter;
  });
  out.height = y;
  return out;
}

function placeTable(t: Table, bi: number, env: Env, flow: Flow) {
  flow.lastAfter = 0;
  const areaW = flow.colW;
  // 百分比宽度的表格：表宽 = 百分比 × 栏宽，各列按网格列宽的比例分
  const gridW = t.cols.reduce((a, b) => a + b, 0);
  const k = t.widthPct && gridW > 0 ? (t.widthPct * areaW) / gridW : 1;
  const tw = gridW * k;
  let left = flow.colX + (t.align === 'center' ? (areaW - tw) / 2 : t.align === 'right' ? areaW - tw : t.indent);
  let col = flow.col;
  const colX = [0];
  for (const w of t.cols) colX.push(colX[colX.length - 1] + w * k);
  const cm = t.cellMargin;

  // 每行各单元格所在的网格列
  const starts = t.rows.map((row) => {
    let g = 0;
    return row.cells.map((c) => {
      const st = g;
      g += c.gridSpan;
      return st;
    });
  });
  const isContinueAt = (r: number, g: number) => {
    const row = t.rows[r];
    if (!row) return false;
    const idx = starts[r].indexOf(g);
    return idx >= 0 && row.cells[idx].vMerge === 'continue';
  };

  const rowBoxes = t.rows.map((row, r) =>
    row.cells.map((c, ci) => {
      const g = starts[r][ci];
      const w = colX[Math.min(g + c.gridSpan, t.cols.length)] - colX[g];
      // x2 和右边单元格的 x 用同一个算式，两条公共边的坐标才完全相同（否则浮点误差会让它们差出一个像素，线就变粗了）
      return { x: left + colX[g], x2: left + colX[Math.min(g + c.gridSpan, t.cols.length)], w, box: layoutCell(t, bi, r, ci, w - cm.left - cm.right, env) };
    }),
  );
  const borderW = t.borders ? (t.borderSize ?? 0.5) : 0;
  const rowHeight = (r: number) => {
    const row = t.rows[r];
    // Word 实测：自动行高 = 上框线 + 单元格上下边距 + 内容（0.5 磅框线、三号字的行 = 20.8 + 0.5）
    const content = Math.max(0, ...rowBoxes[r].map((c) => c.box.height)) + cm.top + cm.bottom + borderW;
    if (row.heightRule === 'exact' && row.height > 0) return row.height;
    if (row.heightRule === 'atLeast') return Math.max(row.height, content);
    return content;
  };

  const placeRow = (r: number) => {
    const h = rowHeight(r);
    const y0 = flow.y;
    flow.regionFrom ??= { bi, line: r === 0 ? 0 : -1 };
    // 换到下一栏时整张表跟着平移
    if (flow.col !== col) {
      const dx = flow.colX - (left - (t.align === 'center' ? (areaW - tw) / 2 : t.align === 'right' ? areaW - tw : t.indent));
      for (const row of rowBoxes) for (const c of row) {
        c.x += dx;
        c.x2 += dx;
      }
      left += dx;
      col = flow.col;
    }
    rowBoxes[r].forEach((c, ci) => {
      const cell = t.rows[r].cells[ci];
      const free = h - cm.top - cm.bottom - borderW - c.box.height;
      const dy = cell.vAlign === 'center' ? free / 2 : cell.vAlign === 'bottom' ? free : 0;
      for (const l of c.box.lines) {
        flow.page.body.push({ story: 'body', path: l.path, line: l.line, x: c.x + cm.left, y: y0 + borderW + cm.top + dy + l.dy });
      }
      if (cell.shading) flow.page.fills.push({ x: c.x, y: y0, w: c.x2 - c.x, h, color: `#${cell.shading}` });
      (flow.page.cells ??= []).push({ bi, r, c: ci, g0: starts[r][ci], g1: starts[r][ci] + cell.gridSpan - 1, x: c.x, y: y0, w: c.x2 - c.x, h });
      if (cell.edges) {
        // 从文件读入的四条边（表格样式合成后的线宽、颜色）
        const g = starts[r][ci];
        const e = cell.edges;
        const col = (c: CellEdge) => (c.color && c.color !== 'auto' ? `#${c.color}` : undefined);
        if (e.top && cell.vMerge !== 'continue') flow.page.borders.push({ x1: c.x, y1: y0, x2: c.x2, y2: y0, w: e.top.w, color: col(e.top), ...(e.top.style && { style: e.top.style }) });
        if (e.bottom && !isContinueAt(r + 1, g)) flow.page.borders.push({ x1: c.x, y1: y0 + h, x2: c.x2, y2: y0 + h, w: e.bottom.w, color: col(e.bottom), ...(e.bottom.style && { style: e.bottom.style }) });
        if (e.left) flow.page.borders.push({ x1: c.x, y1: y0, x2: c.x, y2: y0 + h, w: e.left.w, color: col(e.left), ...(e.left.style && { style: e.left.style }) });
        if (e.right) flow.page.borders.push({ x1: c.x2, y1: y0, x2: c.x2, y2: y0 + h, w: e.right.w, color: col(e.right), ...(e.right.style && { style: e.right.style }) });
      } else if (cell.borders ?? t.borders) {
        const g = starts[r][ci];
        const x1 = c.x;
        const x2 = c.x2;
        const w = t.borderSize ?? 0.5;
        if (cell.vMerge !== 'continue') flow.page.borders.push({ x1, y1: y0, x2, y2: y0, w });
        if (!isContinueAt(r + 1, g)) flow.page.borders.push({ x1, y1: y0 + h, x2, y2: y0 + h, w });
        flow.page.borders.push({ x1, y1: y0, x2: x1, y2: y0 + h, w });
        flow.page.borders.push({ x1: x2, y1: y0, x2, y2: y0 + h, w });
      }
    });
    flow.y += h;
    flow.hasContent = true;
  };

  if (t.float) {
    // 浮动表格：放在 tblpPr 指定的位置，不占正文高度，后面的文字绕开它
    const f = t.float;
    const s = flow.setup;
    const H = t.rows.reduce((a, _, r) => a + rowHeight(r), 0);
    const [hx, hw] = f.hAnchor === 'page' ? [0, s.width] : f.hAnchor === 'margin' ? [s.marginLeft, textWidth(s)] : [flow.colX, flow.colW];
    const [vy, vh] = f.vAnchor === 'page' ? [0, s.height] : f.vAnchor === 'margin' ? [s.marginTop, textHeight(s)] : [flow.y, 0];
    const fx = f.xAlign ? floatCoord({ rel: '', offset: 0, align: f.xAlign }, hx, hw, tw) : hx + f.x;
    const fy = f.yAlign && f.yAlign !== 'inline' ? floatCoord({ rel: '', offset: 0, align: f.yAlign }, vy, vh, H) : vy + f.y;
    const dx = fx - left;
    for (const row of rowBoxes) for (const c of row) {
      c.x += dx;
      c.x2 += dx;
    }
    left += dx;
    const y0 = flow.y;
    const had = flow.hasContent;
    flow.y = fy;
    t.rows.forEach((_, r) => placeRow(r));
    flow.y = y0;
    flow.hasContent = had;
    flow.addExclusion(fx, fy, tw, H, 'square', { l: f.dist.l, r: f.dist.r, t: f.dist.t, b: f.dist.b });
    return;
  }
  const headerRows: number[] = [];
  for (let r = 0; r < t.rows.length && t.rows[r].header; r++) headerRows.push(r);
  t.rows.forEach((_, r) => {
    const h = rowHeight(r);
    if (flow.overflow) return;
    if (flow.y + h > flow.bottom + EPS && flow.hasContent) {
      flow.nextColumn();
      if (flow.overflow) return;
      if (r >= headerRows.length) for (const hr of headerRows) placeRow(hr);
    }
    placeRow(r);
  });
}

function firstLineHeight(doc: Doc, bi: number, areaW: number, env: Env): number {
  const b = doc.body[bi];
  if (!b) return 0;
  if (b.kind === 'table') return 20;
  const pl = layoutParagraph(b, areaW, env);
  return pl.spaceBefore + (pl.lines[0]?.height ?? 0);
}

/** 按对齐方式或偏移量，在参照范围 [start, start + size] 里定位长度为 len 的对象 */
function floatCoord(pos: TextBox['posH'], start: number, size: number, len: number): number {
  if (pos.align === 'center') return start + (size - len) / 2;
  if (pos.align === 'right' || pos.align === 'bottom' || pos.align === 'outside') return start + size - len;
  if (pos.align) return start;
  return start + pos.offset;
}

/** 浮动对象左上角：水平按 page / margin / 其它（所在栏），垂直按 page / margin / line（锚点所在行）/ 其它（段落顶部） */
function floatXY(b: TextBox, s: PageSetup, colX: number, colW: number, paraTop: number, lineY: number, lineH: number): { x: number; y: number } {
  const x = b.posH.rel === 'page' ? floatCoord(b.posH, 0, s.width, b.w) : b.posH.rel === 'margin' ? floatCoord(b.posH, s.marginLeft, textWidth(s), b.w) : floatCoord(b.posH, colX, colW, b.w);
  const y =
    b.posV.rel === 'page'
      ? floatCoord(b.posV, 0, s.height, b.h)
      : b.posV.rel === 'margin'
        ? floatCoord(b.posV, s.marginTop, textHeight(s), b.h)
        : b.posV.rel === 'line'
          ? floatCoord(b.posV, lineY, lineH, b.h)
          : floatCoord(b.posV, paraTop, 0, b.h);
  return { x, y };
}

/** 文本框的位置和里面的文字。at 为锚点所在的行。 */
function placeBox(a: Anchor, at: PlacedLine, page: PageLayout, env: Env, known?: { x: number; y: number }, lines: PlacedLine[] = page.body): FloatBox {
  const b = a.box;
  if (known) return placeBoxAt(b, known.x, known.y, env);
  const s = page.setup;
  // 水平参照：页面、页边距之间的版心，其它（栏 / 字符）按段落所在的栏或单元格
  const x =
    b.posH.rel === 'page'
      ? floatCoord(b.posH, 0, s.width, b.w)
      : b.posH.rel === 'margin'
        ? floatCoord(b.posH, s.marginLeft, textWidth(s), b.w)
        : floatCoord(b.posH, at.x, a.areaW, b.w);
  // 垂直参照：页面、版心、锚点所在的行，其它按段落顶部（本页上这个段落的第一行）
  const paraTop = lines.find((l) => l.story === at.story && samePath(l.path, at.path))?.y ?? at.y;
  const y =
    b.posV.rel === 'page'
      ? floatCoord(b.posV, 0, s.height, b.h)
      : b.posV.rel === 'margin'
        ? floatCoord(b.posV, s.marginTop, textHeight(s), b.h)
        : b.posV.rel === 'line'
          ? floatCoord(b.posV, at.y, at.line.height, b.h)
          : floatCoord(b.posV, paraTop, 0, b.h);
  return placeBoxAt(b, x, y, env);
}

function placeBoxAt(b: TextBox, x: number, y: number, env: Env): FloatBox {
  const innerW = Math.max(10, b.w - b.inset.l - b.inset.r);
  const inner: { line: Line; dy: number }[] = [];
  let h = 0;
  for (const p of b.paragraphs) {
    const lay = layoutParagraph(p, innerW, env);
    h += lay.spaceBefore;
    for (const line of lay.lines) {
      inner.push({ line, dy: h });
      h += line.height;
    }
    h += lay.spaceAfter;
  }
  const free = Math.max(0, b.h - b.inset.t - b.inset.b - h);
  const top = y + b.inset.t + (b.vAlign === 'center' ? free / 2 : b.vAlign === 'bottom' ? free : 0);
  return { x, y, box: b, lines: inner.map((l) => ({ story: 'body' as const, path: [], line: l.line, x: x + b.inset.l, y: top + l.dy })) };
}

export function layoutDocument(doc: Doc, m: Measurer): DocLayout {
  const sections = sectionList(doc);
  /** 原文件里没有这个页眉 / 页脚（占位段落还没写内容）：不占正文的位置 */
  const hfEmpty = (ps: Paragraph[]) => ps.length === 1 && !!ps[0].hfNone && !ps[0].runs.length;
  // 页眉页脚的高度（段落 + 表格；文本框架不占高度），用来让正文避开比页边距高的页眉页脚
  const hfEnv0 = makeEnv(doc, m, false, () => '1');
  const scratch = { fills: [], borders: [] } as unknown as PageLayout;
  const hfHeight = (ps: Paragraph[], w: number) =>
    ps.reduce((a, p) => {
      if (p.hfTable) return a + placeStaticTable(p.hfTable.table, 0, 0, w, hfEnv0, scratch, [], 'header', []);
      if (p.props.frame) return a;
      const pl = layoutParagraph(p, w, hfEnv0);
      return a + pl.spaceBefore + pl.spaceAfter + pl.lines.reduce((b, l) => b + l.height, 0);
    }, 0);
  const flow = new Flow(sections, (page, first, si) => {
    const sect = sections[Math.min(si, sections.length - 1)];
    const s = sect.setup;
    const w = textWidth(s);
    const hs = hfStoryFor(sect.hf.titlePg, s.evenAndOdd, first, page, 'header');
    const fs = hfStoryFor(sect.hf.titlePg, s.evenAndOdd, first, page, 'footer');
    // 没有页眉 / 页脚时不占位置，正文只受页边距限制（Word：下边距 12 磅、页脚距边界 50 磅、没有页脚的文档正文排到下边距；
    // 存在但只有一个空段落的页眉仍按“页眉距边界 + 空段落高”把正文往下推）
    const head = sect.hf[hs];
    const foot = sect.hf[fs];
    return {
      top: hfEmpty(head) ? 0 : s.headerDistance + hfHeight(head, w),
      bottom: hfEmpty(foot) ? s.height : s.height - s.footerDistance - hfHeight(foot, w),
    };
  });
  const labels = computeLabels(doc);
  const env = { ...makeEnv(doc, m, true, (f) => (f === 'PAGE' ? String(flow.pages.length) : '#')), labels };
  // 脚注：编号按引用在正文里出现的顺序；脚注区 = 分隔线所在的一行（正文样式的空段落）+ 各条脚注
  const noteNo = new Map<string, number>();
  const visit = (ps: Paragraph[]) => {
    for (const p of ps) for (const r of p.runs) if (r.opaque?.note?.kind === 'foot' && r.opaque.note.id !== '@self' && !noteNo.has(r.opaque.note.id)) noteNo.set(r.opaque.note.id, noteNo.size + 1);
  };
  for (const b of doc.body) {
    if (b.kind === 'p') visit([b]);
    else for (const row of b.rows) for (const c of row.cells) visit(c.paragraphs);
  }
  env.notes = noteNo;
  const notePara = (id: string, width: number) =>
    (doc.footnotes?.[id] ?? []).map((p) => layoutParagraph(p, width, { ...env, noteSelf: noteNo.get(id) }));
  const sepPara: Paragraph = { kind: 'p', props: { ...newParagraph('').props, spaceBefore: 0, spaceAfter: 0, spaceBeforeLines: null, spaceAfterLines: null, firstLine: 0, firstLineChars: null }, runs: [], mark: { ...doc.defaults } };
  const paraH = (pl: ParaLayout) => pl.spaceBefore + pl.spaceAfter + pl.lines.reduce((a, l) => a + l.height, 0);
  flow.noteHeight = (ids) => {
    if (!ids.length || !doc.footnotes) return 0;
    const w = textWidth(flow.setup);
    return paraH(layoutParagraph(sepPara, w, env)) + ids.reduce((a, id) => a + notePara(id, w).reduce((b, pl) => b + paraH(pl), 0), 0);
  };
  // 各节的版心宽度可以不同；分栏时按当前栏的宽度
  const place = (bi: number, from?: RegionFrom) => {
    const b = doc.body[bi];
    if (b.kind === 'table') return placeTable(b, bi, env, flow);
    const areaW = from?.width ?? flow.colW;
    let pl = layoutParagraph(b, areaW, env);
    if (isBareContinuousBreak(b)) pl = { spaceBefore: 0, spaceAfter: 0, lines: pl.lines.map((l) => ({ ...l, height: 0, fit: 0 })) };
    if (from?.lines) pl = { ...pl, lines: from.lines };
    placeParagraph(pl, b, [bi], flow, b.props.keepNext ? firstLineHeight(doc, bi + 1, areaW, env) : 0, env, from?.line ?? 0, areaW);
  };
  /**
   * 平衡本页这一节的各栏：从区域起点重排到 end，栏高从小到大二分查找能全部放下的最小值。
   * 区域从表格中间开始时不平衡。
   */
  const balance = (end: number) => {
    const from = flow.regionFrom;
    if (!from || from.line < 0 || flow.columns.length <= 1) return;
    const m = flow.regionMark();
    const run = (h: number) => {
      flow.restore(m);
      flow.bottom = flow.regionTop + h;
      flow.trial = true;
      for (let bi = from.bi; bi <= end && !flow.overflow; bi++) place(bi, bi === from.bi ? from : undefined);
      flow.trial = false;
      const ok = !flow.overflow;
      flow.bottom = m.bottom;
      return ok;
    };
    let lo = 0;
    let hi = m.bottom - flow.regionTop;
    if (!run(hi)) {
      flow.restore(m);
      for (let bi = from.bi; bi <= end; bi++) place(bi, bi === from.bi ? from : undefined);
      return;
    }
    while (hi - lo > 0.05) {
      const mid = (lo + hi) / 2;
      if (run(mid)) hi = mid;
      else lo = mid;
    }
    run(hi);
  };
  doc.body.forEach((b, bi) => {
    place(bi);
    if (b.kind === 'p' && b.props.sectionBreak) endSection(flow, b.props.sectionBreak, bi, () => balance(bi));
  });
  flow.closeRegion();

  // 脚注：排在页底，上面一条分隔线（版心宽度的三分之一，与 Word 默认的分隔线相同）
  for (const page of flow.pages) {
    if (!page.noteIds?.length) continue;
    const s = page.setup;
    const w = textWidth(s);
    let y = s.height - s.marginBottom - (page.noteArea ?? 0);
    const sep = layoutParagraph(sepPara, w, env);
    const sepH = paraH(sep);
    page.borders.push({ x1: s.marginLeft, y1: y + sepH / 2, x2: s.marginLeft + w / 3, y2: y + sepH / 2, w: 0.5 });
    y += sepH;
    page.notes = [];
    for (const id of page.noteIds) {
      for (const pl of notePara(id, w)) {
        y += pl.spaceBefore;
        for (const line of pl.lines) {
          page.notes.push({ story: 'body', path: [], line, x: s.marginLeft, y });
          y += line.height;
        }
        y += pl.spaceAfter;
      }
    }
  }

  // 文本框：正文排完后，按锚点所在的行 / 段落 / 页面定位
  const boxEnv = makeEnv(doc, m, false, () => '#');
  for (const page of flow.pages) {
    for (const pl of page.body) for (const a of pl.line.anchors ?? []) page.floats.push(placeBox(a, pl, page, boxEnv, flow.floatPos.get(a.box)));
  }

  // 页眉页脚（不使用文档网格），每页单独排版以显示页码
  const total = flow.pages.length;
  for (const page of flow.pages) {
    const s = page.setup;
    const areaW = textWidth(s);
    const hfEnv = makeEnv(doc, m, false, (f) => (f === 'PAGE' ? String(page.index + 1) : String(total)));
    // 各节的页眉页脚；首页不同 / 奇偶页不同：按页选用分区
    const sect = sections[Math.min(page.section, sections.length - 1)];
    const first = flow.sectionFirst[page.index];
    const hs = hfStoryFor(sect.hf.titlePg, s.evenAndOdd, first, page.index, 'header');
    const fs = hfStoryFor(sect.hf.titlePg, s.evenAndOdd, first, page.index, 'footer');
    page.hf = { header: hs, footer: fs, owner: sect.owner };
    // 页眉页脚的段落：边框画出来；文本框架（居中的页码等）不占高度
    const placeHF = (ps: Paragraph[], pls: ParaLayout[], y: number, out: PlacedLine[], story: HFStory) => {
      pls.forEach((pl, pi) => {
        const p = ps[pi];
        const y0 = y;
        if (p.hfTable) {
          y += placeStaticTable(p.hfTable.table, s.marginLeft, y, areaW, hfEnv, page, out, story, hfPath(sect.owner, pi));
          return;
        }
        y += pl.spaceBefore;
        const top = y;
        for (const line of pl.lines) {
          out.push({ story, path: hfPath(sect.owner, pi), line, x: s.marginLeft, y });
          y += line.height;
        }
        page.borders.push(...paraBorderSegs(p, s.marginLeft, areaW, top, y));
        y += pl.spaceAfter;
        if (p.props.frame) y = y0;
      });
    };
    const head = sect.hf[hs].map((p) => layoutParagraph(p, areaW, hfEnv));
    placeHF(sect.hf[hs], head, s.headerDistance, page.header, hs);
    const foot = sect.hf[fs].map((p) => layoutParagraph(p, areaW, hfEnv));
    // 表格的高度先试排一遍（不留下线和行）
    const scratch = { fills: [], borders: [] } as unknown as PageLayout;
    const fh = foot.reduce((a, pl, i) => {
      const p = sect.hf[fs][i];
      if (p.hfTable) return a + placeStaticTable(p.hfTable.table, 0, 0, areaW, hfEnv, scratch, [], fs, []);
      return p.props.frame ? a : a + pl.spaceBefore + pl.spaceAfter + pl.lines.reduce((b, l) => b + l.height, 0);
    }, 0);
    placeHF(sect.hf[fs], foot, s.height - s.footerDistance - fh, page.footer, fs);
    // 页眉页脚里的文本框 / 形状 / 浮动图片：同正文的定位规则，不影响正文排版
    for (const lines of [page.header, page.footer]) for (const pl of lines) for (const a of pl.line.anchors ?? []) page.floats.push({ ...placeBox(a, pl, page, boxEnv, undefined, lines), hf: true });
  }
  if (doc.watermark) for (const pg of flow.pages) pg.watermark = doc.watermark;
  return { pages: flow.pages };
}

// ---------------- 光标与点击定位 ----------------

/** 某页上属于指定故事的行。 */
export function pageStoryLines(pg: PageLayout, story: Story): PlacedLine[] {
  if (story === 'body') return pg.body;
  return (isHeaderStory(story) ? pg.header : pg.footer).filter((l) => l.story === story);
}

export function samePath(a: number[], b: number[]): boolean {
  return a.length === b.length && a.every((v, i) => v === b[i]);
}

export function comparePos(a: Pos, b: Pos): number {
  const n = Math.min(a.path.length, b.path.length);
  for (let i = 0; i < n; i++) if (a.path[i] !== b.path[i]) return a.path[i] - b.path[i];
  if (a.path.length !== b.path.length) return a.path.length - b.path.length;
  return a.offset - b.offset;
}

/** 页眉页脚的行是否属于路径 at 所在的那一节的分区 */
function sameHFContainer(pl: PlacedLine, at: number[]): boolean {
  return pl.path.length === at.length && (at.length < 2 || pl.path[0] === at[0]);
}

/**
 * 某个故事（正文/页眉/页脚）在页面上的所有行，正文跨所有页，页眉页脚只取指定页。
 * at 为页眉页脚中的某个路径，用来区分各节的同名分区。
 */
export function storyLines(L: DocLayout, story: Story, page: number, at?: number[]): { pl: PlacedLine; page: number }[] {
  if (story === 'body') return L.pages.flatMap((p) => p.body.map((pl) => ({ pl, page: p.index })));
  const linesOf = (p: PageLayout) => pageStoryLines(p, story).filter((pl) => !at || sameHFContainer(pl, at));
  // 页眉页脚分区可能出现在多页：优先取指定页，没有则取第一个用到它的页
  const pg = L.pages[Math.min(page, L.pages.length - 1)];
  const onPage = pg ? linesOf(pg) : [];
  if (onPage.length) return onPage.map((pl) => ({ pl, page: pg!.index }));
  const other = L.pages.find((p) => linesOf(p).length);
  return other ? linesOf(other).map((pl) => ({ pl, page: other.index })) : [];
}

export function findLine(L: DocLayout, pos: Pos, pageHint: number): { pl: PlacedLine; page: number } | null {
  const cands = storyLines(L, pos.story, pageHint, pos.path).filter(
    ({ pl }) => samePath(pl.path, pos.path) && pos.offset >= pl.line.start && pos.offset <= pl.line.end,
  );
  // 偏移正好在两行交界处时属于后一行（手动换行、自动换行之后）
  return cands.find(({ pl }) => pos.offset < pl.line.end) ?? cands[cands.length - 1] ?? null;
}

/** 表格一行在某页上的行尾标记：x 为这一行最右边的格线，y / h 为这一行在本页的范围，baseline / lineH 取行里第一行文字 */
export interface RowEnd {
  bi: number;
  r: number;
  x: number;
  y: number;
  h: number;
  baseline: number;
  lineH: number;
  /** 这一行第一行文字（行尾标记的颜色跟它的文字一致） */
  line?: Line;
}

/** 本页正文里各表格行的行尾标记位置（Word：行尾标记画在表格右边框外，与该行第一行文字对齐；回车在下面插入一行） */
export function rowEnds(page: PageLayout): RowEnd[] {
  const rows = new Map<string, RowEnd>();
  for (const c of page.cells ?? []) {
    const k = `${c.bi}:${c.r}`;
    const e = rows.get(k);
    if (!e) rows.set(k, { bi: c.bi, r: c.r, x: c.x + c.w, y: c.y, h: c.h, baseline: c.y + c.h, lineH: c.h });
    else {
      e.x = Math.max(e.x, c.x + c.w);
      e.h = Math.max(e.y + e.h, c.y + c.h) - Math.min(e.y, c.y);
      e.y = Math.min(e.y, c.y);
    }
  }
  for (const e of rows.values()) {
    // 这一行第一行文字（纵向最靠上的）；空行没有文字时按行底
    let best: PlacedLine | null = null;
    for (const l of page.body) if (l.path.length > 1 && l.path[0] === e.bi && l.path[1] === e.r && (!best || l.y < best.y)) best = l;
    if (best) {
      e.baseline = best.y + best.line.baseline;
      e.lineH = best.line.height;
      e.line = best.line;
    }
  }
  return [...rows.values()];
}

export function caretRect(L: DocLayout, pos: Pos, pageHint: number): { page: number; x: number; y: number; h: number } | null {
  const hit = findLine(L, pos, pageHint);
  if (!hit) return null;
  const { pl } = hit;
  return { page: hit.page, x: pl.x + pl.line.xs[pos.offset - pl.line.start], y: pl.y, h: pl.line.height };
}

function offsetAtX(pl: PlacedLine, x: number): number {
  const { line } = pl;
  let best = 0;
  let bestD = Infinity;
  // 行末若是换行/分页符，光标不能放在其后（那属于下一行）
  const maxK = line.brk === 'line' || line.brk === 'page' || line.brk === 'column' ? line.xs.length - 2 : line.xs.length - 1;
  for (let k = 0; k <= Math.max(0, maxK); k++) {
    const d = Math.abs(pl.x + line.xs[k] - x);
    if (d < bestD) {
      bestD = d;
      best = k;
    }
  }
  return line.start + best;
}

/** 页面坐标 (x, y) 正好落在这一页的某一行正文上（可读写的行）就返回它；正文行可以落在页边距里（浮动表格等） */
export function bodyLineAt(pg: PageLayout, x: number, y: number): PlacedLine | null {
  return pg.body.find((pl) => !pl.ro && y >= pl.y && y <= pl.y + pl.line.height && x >= pl.x + pl.line.xs[0] - 2 && x <= pl.x + Math.max(pl.line.right, pl.line.xs[0] + 4) + 2) ?? null;
}

export function hitTest(L: DocLayout, page: number, x: number, y: number, story: Story): Pos | null {
  const pg = L.pages[page];
  if (!pg) return null;
  let lines = pageStoryLines(pg, story).filter((l) => !l.ro);
  if (!lines.length) return null;
  // 分栏：先在点击位置所在的栏里找
  const cols = story === 'body' ? columnBoxes(pg.setup) : [];
  if (cols.length > 1) {
    const ml = pg.setup.marginLeft;
    const ci = cols.findIndex((c, i) => i === cols.length - 1 || x < ml + (c.x + c.w + cols[i + 1].x) / 2);
    const lo = ci === 0 ? -Infinity : ml + cols[ci].x - EPS;
    const hi = ci === cols.length - 1 ? Infinity : ml + cols[ci + 1].x - EPS;
    const inCol = lines.filter((pl) => pl.x >= lo && pl.x < hi);
    if (inCol.length) lines = inCol;
  }
  const vDist = (pl: PlacedLine) => (y < pl.y ? pl.y - y : y > pl.y + pl.line.height ? y - pl.y - pl.line.height : 0);
  const hDist = (pl: PlacedLine) => {
    const l = pl.x + pl.line.xs[0];
    const r = Math.max(pl.x + pl.line.right, l + 4);
    return x < l ? l - x : x > r ? x - r : 0;
  };
  let best = lines[0];
  let bestScore = Infinity;
  for (const pl of lines) {
    const score = vDist(pl) * 1000 + hDist(pl);
    if (score < bestScore) {
      bestScore = score;
      best = pl;
    }
  }
  return { story, path: best.path, offset: offsetAtX(best, x) };
}

/** 上下移动光标：找相邻视觉行中与 goalX 水平重叠的一行。 */
export function moveVertical(L: DocLayout, pos: Pos, dir: 1 | -1, pageHint: number, goalX: number): { pos: Pos; page: number } | null {
  const lines = storyLines(L, pos.story, pageHint, pos.path).sort((a, b) => a.page - b.page || a.pl.y - b.pl.y || a.pl.x - b.pl.x);
  const cur = findLine(L, pos, pageHint);
  if (!cur) return null;
  const idx = lines.findIndex((l) => l.pl === cur.pl);
  const overlaps = (pl: PlacedLine) => goalX >= pl.x + pl.line.xs[0] - 2 && goalX <= pl.x + Math.max(pl.line.right, pl.line.xs[0] + 4) + 40;
  let fallback: { pl: PlacedLine; page: number } | null = null;
  for (let i = idx + dir; i >= 0 && i < lines.length; i += dir) {
    const c = lines[i];
    if (c.page === cur.page && Math.abs(c.pl.y - cur.pl.y) < 0.5) continue;
    fallback ??= c;
    if (overlaps(c.pl)) return { pos: { story: pos.story, path: c.pl.path, offset: offsetAtX(c.pl, goalX) }, page: c.page };
    if (fallback && (c.page !== fallback.page || Math.abs(c.pl.y - fallback.pl.y) > 0.5)) break;
  }
  return fallback ? { pos: { story: pos.story, path: fallback.pl.path, offset: offsetAtX(fallback.pl, goalX) }, page: fallback.page } : null;
}

/** 选区高亮矩形（页面坐标）。 */
export function selectionRects(L: DocLayout, a: Pos, b: Pos, pageHint: number): { page: number; x: number; y: number; w: number; h: number }[] {
  const [s, e] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  const out: { page: number; x: number; y: number; w: number; h: number }[] = [];
  for (const { pl, page } of storyLines(L, s.story, pageHint, s.path)) {
    const ls: Pos = { story: s.story, path: pl.path, offset: pl.line.start };
    const le: Pos = { story: s.story, path: pl.path, offset: pl.line.end };
    if (comparePos(le, s) < 0 || comparePos(ls, e) > 0) continue;
    const from = comparePos(ls, s) >= 0 ? pl.line.start : s.offset;
    const to = comparePos(le, e) <= 0 ? pl.line.end : e.offset;
    const x1 = pl.x + pl.line.xs[Math.max(0, from - pl.line.start)];
    let x2 = pl.x + pl.line.xs[Math.min(pl.line.xs.length - 1, to - pl.line.start)];
    // 选中到段落末尾时多画一点，表示段落标记被选中
    if (comparePos(le, e) < 0 && (pl.line.brk === 'end' || pl.line.brk === 'line')) x2 += 4;
    if (x2 > x1) out.push({ page, x: x1, y: pl.y, w: x2 - x1, h: pl.line.height });
  }
  return out;
}

/**
 * Word 文档模型。长度单位统一为磅（pt）；1 pt = 20 twips，1 cm = 28.3465 pt。
 * 段落文本中的特殊字符：'\t' 制表符，'\n' 手动换行（Shift+Enter），'\f' 分页符，'\v' 分栏符，
 * FIELD_CHAR 为域（页码 / 总页数）和嵌入式图片的占位符，每个域 / 每张图片占一个字符。
 */
import type { DocProtection } from '../../vendor/docx-engine/src/types';

export const FIELD_CHAR = '￼';
export const PT_PER_CM = 72 / 2.54;

export type FieldKind = 'PAGE' | 'NUMPAGES';

export interface RunProps {
  /** 中文字体 */
  fontEA: string;
  /** 西文字体 */
  fontLatin: string;
  /** 字号（磅） */
  size: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** 'RRGGBB' 或 'auto' */
  color: string;
  /** 删除线 */
  strike?: boolean;
  /** 上标 / 下标（字号按 65% 显示） */
  vertAlign?: 'superscript' | 'subscript';
  /** 字符间距（w:spacing，磅；正数加宽、负数紧缩）：加在每个字符后面 */
  spacing?: number;
  /** 突出显示 / 文字底纹 'RRGGBB' */
  highlight?: string;
  /** 修订：这段文字是谁在什么时候插入 / 删除的（还没被接受或拒绝） */
  rev?: Revision;
  /**
   * 超链接：外部链接的关系 id（r:id）、新加的外部链接地址（url，保存时才生成关系）或文档内书签（anchor）。
   * 读入时 rid 对应的地址也记在 url 里（只用来显示和修改；有 rid 时保存仍写原来的 rid）
   */
  link?: { rid?: string; anchor?: string; url?: string };
  /** 字符样式（w:rStyle，原文件里的样式 id）；格式已按样式算进上面各项，保存时写回引用 */
  rStyle?: string;
}

/**
 * 原样对象：本编辑器不能编辑、但要在保存时原样写回的内容（公式、书签、形状、图表、OLE、脚注引用等）。
 * 占一个 FIELD_CHAR；w / h 是显示占位的大小（0 = 不占位置），image 是能显示的预览（如公式渲染成的图）。
 */
export interface Opaque {
  /** 原始 XML（段落级元素或 w:r 里面的一个子元素） */
  xml: string;
  /** para：段落的直接子元素（m:oMath、w:bookmarkStart…）；run：w:r 里的子元素（w:drawing、w:object…） */
  level: 'para' | 'run';
  kind: 'math' | 'bookmark' | 'drawing' | 'object' | 'note' | 'other';
  w: number;
  h: number;
  /** 公式的文字（OMML 里的 m:t），显示占位和导出纯文本用 */
  text?: string;
  /** 脚注 / 尾注的引用（kind = note）：id 为 footnotes.xml 里的编号；'@self' 是脚注正文开头的“本条编号” */
  note?: { kind: 'foot' | 'end'; id: string };
  image?: RunImage;
  /** 没有缓存结果的 EQ 域（公式域，如带圈字符 eq \o\ac(□,√)）的显示：overstrike 时各部分叠在一起（宽度取最宽的），否则依次排开 */
  eq?: EqDisplay;
  /** Word 的邮件合并域（MERGEFIELD）：字段名和 \b / \f 开关（有值时才加的前 / 后缀）；显示为 «字段名»（用 eq 显示），保存时原样写回 */
  merge?: { name: string; before?: string; after?: string; numFmt?: string; dateFmt?: string; caseFmt?: string };
  /** Word 的规则域 IF / NEXT / NEXTIF / SKIPIF：指令整串（嵌套的域写成花括号），填值时求值；显示为 «如果 …» */
  rule?: { full: string };
}

export interface EqDisplay {
  overstrike: boolean;
  align: 'left' | 'center' | 'right';
  /** 各部分；每部分是若干带格式的文字段 */
  parts: { text: string; props: RunProps }[][];
}

export interface Revision {
  kind: 'ins' | 'del';
  author: string;
  /** ISO 时间 */
  date: string;
}

/** 批注。回复（parent）没有自己的范围标记，跟随被回复的那条 */
export interface Comment {
  id: string;
  author: string;
  initials?: string;
  date: string;
  /** 批注文字，多段用换行连接 */
  text: string;
  parent?: string;
}

/** 嵌入式图片（随文字排版）：图片数据在 images.ts 的仓库里，这里只记 id 和显示尺寸（磅） */
export interface RunImage {
  id: string;
  w: number;
  h: number;
  /** DrawingML picture transform. Dimensions are before rotation, in points. */
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** Fractions removed from each edge of the original image; negative = padding. */
  crop?: { l: number; t: number; r: number; b: number };
  /** Image id associated with raw XML, so replacement can update its relationship. */
  rawId?: string;
  /** 从文件读入的图片：原始 w:drawing XML（保存时尺寸没变就原样写回，保留裁剪、效果、替代文字等） */
  raw?: string;
}

/** 浮动对象的位置：相对 rel 指的参照物，按 align 对齐或偏移 offset 磅 */
export interface FloatPos {
  /** 水平：column（段落所在的栏 / 单元格）、margin、page…；垂直：paragraph、line、margin、page… */
  rel: string;
  offset: number;
  /** left / center / right（水平）或 top / center / bottom（垂直）；设置时不用 offset */
  align?: string;
}

/**
 * 文本框（浮动在文字上方 / 下方，锚定在某个段落里）。只读显示：里面的文字暂时不能在编辑器里修改，保存时原样写回。
 */
export interface TextBox {
  w: number;
  h: number;
  posH: FloatPos;
  posV: FloatPos;
  /** 文字环绕方式：none 浮于文字上方 / 衬于文字下方；square / tight / through 四周型（按外框绕排）；topAndBottom 上下型 */
  wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  /** 环绕时与文字的距离（wp:anchor 的 distT / distB / distL / distR，磅），缺省 0 */
  dist?: { t: number; b: number; l: number; r: number };
  /** 衬于文字下方 */
  behind: boolean;
  /** 填充色 'RRGGBB'，null 为无填充 */
  fill: string | null;
  /** 边框线，null 为无边框 */
  line: { color: string; w: number } | null;
  /** 内边距（磅） */
  inset: { l: number; t: number; r: number; b: number };
  vAlign: 'top' | 'center' | 'bottom';
  paragraphs: Paragraph[];
  /** 浮动图片（公章、签名、Logo）：有它时这是一张图片，不画填充、边框和文字，可以拖动和缩放 */
  image?: RunImage;
  /**
   * 形状的外形（DrawingML 预设形状名：roundRect、ellipse、line…，见 core/shapes.ts 的 SHAPE_GEOMS）；
   * 没有 = 文本框（矩形）。形状可以没有文字（paragraphs 为空）。
   */
  geom?: string;
  /** 水平 / 垂直翻转（只翻转外形，文字不翻转） */
  flipH?: boolean;
  flipV?: boolean;
  /** 顺时针旋转角度（度），绕框的中心 */
  rot?: number;
  /** 从文件读入的：原始 XML（w:drawing 或 mc:AlternateContent）；保存时内容、位置、大小都没变就原样写回 */
  raw?: string;
}

export interface Run {
  text: string;
  props: RunProps;
  field?: FieldKind;
  /** 图片：text 为一个 FIELD_CHAR */
  image?: RunImage;
  /** 文本框的锚点：text 为一个 FIELD_CHAR，不占宽度 */
  box?: TextBox;
  /** 批注范围的开始 / 结束标记：text 为一个 FIELD_CHAR，不占宽度 */
  mark?: { kind: 'cs' | 'ce'; id: string };
  /** 原样对象（见 Opaque）：text 为一个 FIELD_CHAR */
  opaque?: Opaque;
  /** 域的开始 / 结束标记（见 FieldMark）：text 为一个 FIELD_CHAR，不占宽度；两者之间的文本段是域的显示结果 */
  fc?: FieldMark;
}

/**
 * 一般的域（REF、DATE、SEQ、FILENAME、= 公式……）：对应 OOXML 的 fldChar begin … separate … end。
 * 开始标记带域代码，结束标记只标位置；中间的文本段就是域结果（普通文字，可以换行、查找、编辑），“更新域”时整体换成新结果。
 * 只认开始和结束都在同一段里、结果里没有别的域的；其它的（跨段的目录域等）照旧只读显示文字。
 */
export interface FieldMark {
  kind: 'begin' | 'end';
  /** 同一个域的开始和结束标记共用；也用来记“显示域代码”（fields.ts） */
  id: string;
  /** 开始标记：域代码（前后空格去掉；代码里嵌套的域写成花括号 {}） */
  instr?: string;
  /** 从文件读入的原始 XML：开始标记 = begin 到 separate 的几个 w:r，结束标记 = end 的 w:r；域代码没改时原样写回 */
  raw?: string;
  /** 开始标记：域已锁定（w:fldLock），更新域时不改结果 */
  locked?: boolean;
  /** 开始标记：域代码里嵌套的域的结果（按左花括号出现的顺序）；按代码生成 XML 时写进嵌套域，Word 算外层域时要用 */
  inner?: string[];
}

/** 不可拆分的文本段（域 / 图片 / 文本框锚点 / 原样对象）：整体占一个字符，不能往里面输入文字 */
export const isAtom = (r: Run): boolean => !!(r.field || r.image || r.box || r.mark || r.opaque || r.fc);

export type Align = 'left' | 'center' | 'right' | 'both' | 'distribute';

/** auto：多倍行距（value 为倍数）；exact：固定值；atLeast：最小值（value 为磅）。 */
export interface LineSpacing {
  rule: 'auto' | 'exact' | 'atLeast';
  value: number;
}

export interface ParaProps {
  align: Align;
  indentLeft: number;
  indentRight: number;
  /** 首行缩进（磅），负数为悬挂缩进 */
  firstLine: number;
  /** 以字符为单位的首行缩进（Word 的 firstLineChars / 100），设置时优先 */
  firstLineChars: number | null;
  spaceBefore: number;
  spaceAfter: number;
  /** 以行为单位的段前/段后（Word 的 beforeLines / 100），设置时优先 */
  spaceBeforeLines: number | null;
  spaceAfterLines: number | null;
  /** 段前 / 段后“自动”（w:beforeAutospacing / w:afterAutospacing，HTML 式间距）：按 14 磅排，文档或页面开头的第一段不加段前 */
  spaceBeforeAuto?: boolean;
  spaceAfterAuto?: boolean;
  line: LineSpacing;
  /** 对齐到文档网格 */
  snapToGrid: boolean;
  keepNext: boolean;
  keepLines: boolean;
  widowControl: boolean;
  pageBreakBefore: boolean;
  /** 自动编号：引用 doc.lists 中的列表及级别（0 起） */
  num: { id: number; level: number } | null;
  /** 本段是一节的最后一段（段后是分节符）；该节的页眉页脚、纸张与页边距见 Paragraph.sectionHF / sectionPage */
  sectionBreak: SectionBreak | null;
  /** 自定义制表位（w:tabs），按位置升序；位置从段落区域左边（左页边距）算起 */
  tabs?: TabStop[];
  /** 大纲级别 0–8（标题 1–9；w:outlineLvl，来自样式或段落本身），正文没有 */
  outline?: number;
  /** 段落样式：只记本编辑器认识的标题（Heading1…）和目录（TOC1…），保存时写 pStyle */
  style?: string;
  /** 段落边框（w:pBdr）：线宽、与文字的距离（磅）、颜色 'RRGGBB' / 'auto' */
  borders?: { top?: ParaBorder; bottom?: ParaBorder; left?: ParaBorder; right?: ParaBorder };
  /** 文本框架（w:framePr）：段落浮在指定位置，不占正文的高度（页脚里居中的页码常用） */
  frame?: ParaFrame;
  /** 原文件里的段落样式 id（w:pStyle）；格式已按样式算进上面各项，字节级保留保存时写回引用、只写与样式不同的格式 */
  styleId?: string;
}

export interface ParaBorder {
  w: number;
  space: number;
  color: string;
}

/** 文本框架：水平相对 hAnchor（margin / page / text）按 xAlign 对齐或偏移 x，垂直相对 vAnchor 偏移 y */
export interface ParaFrame {
  hAnchor: string;
  vAnchor: string;
  xAlign?: string;
  x: number;
  y: number;
}

export interface TabStop {
  pos: number;
  align: 'left' | 'center' | 'right';
  leader: 'none' | 'dot' | 'hyphen' | 'underscore';
}

export type SectionBreak = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';

export const SECTION_BREAK_LABEL: Record<SectionBreak, string> = {
  nextPage: '分节符（下一页）',
  continuous: '分节符（连续）',
  evenPage: '分节符（偶数页）',
  oddPage: '分节符（奇数页）',
};

// ---------------- 自动编号 ----------------

export type NumFmt =
  | 'decimal'
  | 'chineseCounting'
  | 'ideographTraditional'
  | 'upperLetter'
  | 'lowerLetter'
  | 'upperRoman'
  | 'lowerRoman'
  | 'decimalEnclosedCircle'
  | 'bullet'
  | 'none';

export interface NumLevel {
  fmt: NumFmt;
  /** 编号文字模板，%1…%9 为各级编号，如 '%1、'、'（%2）'、'%1.%2.'；项目符号为符号本身 */
  text: string;
  start: number;
  /** 应用编号时写入段落的缩进 */
  indentLeft: number;
  /** 首行缩进（磅），负数为悬挂缩进 */
  firstLine: number;
  firstLineChars: number | null;
  /** 编号与正文之间：制表符 / 空格 / 无 */
  suffix: 'tab' | 'space' | 'nothing';
  /** 编号自己的字符格式（加粗、字号、字体…），没写的项跟随正文 */
  label?: Partial<RunProps>;
}

export interface NumList {
  id: number;
  /** 预设名（用于在同一文档中继续编号），来自 docx 的列表为 docx-<numId> */
  preset: string;
  levels: NumLevel[];
}

export interface Paragraph {
  kind: 'p';
  props: ParaProps;
  runs: Run[];
  /** 段落标记的字符格式：空段落的行高、新输入文字的默认格式 */
  mark: RunProps;
  /** 带分节符的段落：它所结束那一节的页眉页脚（缺省时与下一节相同） */
  sectionHF?: SectionHF;
  /** 带分节符的段落：它所结束那一节的纸张与页边距（缺省时与最后一节即 Doc.setup 相同） */
  sectionPage?: PageGeom;
  /** 带分节符的段落：它所结束那一节的分栏（缺省为一栏，不沿用最后一节） */
  sectionCols?: Columns;
  /** 来自原文件 document.xml 的第几个顶层元素（docxIndex）；字节级保留保存据此判断这块改没改过 */
  src?: number;
  /** 页眉页脚的段落：来自原文件的哪个部件（如 word/header1.xml） */
  hfPart?: string;
  /** 页眉页脚里的表格（只读显示）：这个“段落”只是表格的占位，排版时画表格，保存时写回原 XML */
  hfTable?: { table: Table; xml: string };
  /** 原文件里没有这个页眉 / 页脚，这是占位的空段落：没写内容前不占正文的位置（存在的空页眉会把正文往下推） */
  hfNone?: true;
}

export interface Cell {
  gridSpan: number;
  vMerge: 'none' | 'restart' | 'continue';
  vAlign: 'top' | 'center' | 'bottom';
  /** 单元格自己的边框设置（Word 的 tcBorders），未设置时沿用表格边框 */
  borders?: boolean;
  /** 底纹（w:shd 的 fill，'RRGGBB'） */
  shading?: string;
  /**
   * 在本编辑器里给单元格设过的框线（写成 w:tcBorders）：列出的边按它画、按它保存，null = 无框线；没列出的边沿用表格边框。
   * 设置时同时更新 edges（排版按 edges 画）。读入的文件不填它（原 tcBorders 原样保留）。
   */
  bdr?: Partial<CellEdges>;
  /**
   * 从文件读入的四条边（表格边框、表格样式的条件格式、单元格自己的 tcBorders 合成后）：null = 不画；
   * 有它时排版按它画线（线宽、颜色），没有时按 Table.borders / Cell.borders 画统一的黑线
   */
  edges?: CellEdges;
  paragraphs: Paragraph[];
}

export interface CellEdge {
  /** 线宽（磅）；双线为两条线加中间空隙的总宽 */
  w: number;
  /** 'RRGGBB'，auto 为黑 */
  color: string;
  /** 线型（w:val）：缺省为单实线 */
  style?: BorderStyle;
}

/** 本编辑器能画、能设置的框线线型（其它线型按单实线画） */
export type BorderStyle = 'single' | 'dashed' | 'dotted' | 'double';

export interface CellEdges {
  top: CellEdge | null;
  bottom: CellEdge | null;
  left: CellEdge | null;
  right: CellEdge | null;
}

export interface Row {
  height: number;
  heightRule: 'auto' | 'atLeast' | 'exact';
  /** 标题行：跨页时重复 */
  header: boolean;
  cells: Cell[];
}

export interface Table {
  kind: 'table';
  /** 网格列宽（磅） */
  cols: number[];
  rows: Row[];
  align: 'left' | 'center' | 'right';
  indent: number;
  borders: boolean;
  /** 边框粗细（磅），缺省 0.5 磅（Word 的 w:sz="4"） */
  borderSize?: number;
  cellMargin: { left: number; right: number; top: number; bottom: number };
  /** 表格宽度是版心（栏）宽度的百分比（w:tblW w:type="pct"，1 = 100%）：排版时各列按比例缩放 */
  widthPct?: number;
  /** 浮动表格（w:tblpPr）：不占正文位置，文字绕排 */
  float?: TableFloat;
  /** 见 Paragraph.src */
  src?: number;
}

/**
 * 浮动表格的位置（OOXML w:tblpPr）：水平相对 hAnchor（text = 所在栏、margin = 版心、page = 纸张），
 * 垂直相对 vAnchor（text = 下一段的顶部、margin、page）；有 xAlign / yAlign 时按对齐，否则偏移 x / y 磅。
 * dist 为与四周文字的距离。
 */
export interface TableFloat {
  hAnchor: 'text' | 'margin' | 'page';
  vAnchor: 'text' | 'margin' | 'page';
  x: number;
  y: number;
  xAlign?: string;
  yAlign?: string;
  dist: { l: number; r: number; t: number; b: number };
}

export type Block = Paragraph | Table;

export type GridType = 'none' | 'lines' | 'linesAndChars' | 'snapToChars';

export interface PageSetup {
  width: number;
  height: number;
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  headerDistance: number;
  footerDistance: number;
  grid: {
    type: GridType;
    /** 行距（磅） */
    linePitch: number;
    /** Word 的 charSpace：在默认字号基础上增加的字符间距，单位 1/4096 磅 */
    charSpace: number;
  };
  /** 默认制表位间隔（磅） */
  defaultTab: number;
  /** 首页不同的页眉页脚 */
  titlePg: boolean;
  /** 奇偶页不同的页眉页脚 */
  evenAndOdd: boolean;
  /** 分栏（w:cols），缺省为一栏 */
  cols?: Columns;
  /** 表格里的行也对齐行网格（settings.xml 兼容选项 w:adjustLineHeightInTable，中文 Word 默认模板带它）；全文统一 */
  tableGrid?: boolean;
  /** 压缩标点（settings.xml 的 w:characterSpacingControl = compressPunctuation，中文 Word 默认）：
   * 两端对齐的行放不下时全角标点可以压到半宽；缺省（doNotCompress）不压。全文统一 */
  compressPunct?: boolean;
  /** 兼容模式（settings.xml 的 compatSetting compatibilityMode：11 = Word 2003，12 = 2007，14 = 2010，15 = 2013 及以后）；
   * 缺省 15。.doc 文件是 11。影响标点压缩和悬挂（layout.ts 的 lineRules） */
  compatMode?: number;
  /** 兼容选项 w:doNotExpandShiftReturn（.doc 的 Copts60.fExpShRtn = 0）：以手动换行符结束的行不按两端对齐拉开；
   * 中文 Word 默认模板带它 */
  noExpandShiftReturn?: boolean;
  /** 关闭标点挤压（settings.xml 的 w:noPunctuationKerning；.doc 的 DopTypography.fKerningPunct = 0）：相邻标点不挤压 */
  noPunctKern?: boolean;
  /** 兼容选项 w:useAltKinsokuLineBreakRules（“使用另一套东亚断行规则”，Word 2003 文档带它）：兼容模式 15 以下时行尾标点不悬挂 */
  altKinsoku?: boolean;
}

/** 分栏：num 栏、栏间距 space（磅）、栏间分隔线；widths 为不等宽时每栏的宽度和它右边的间距 */
export interface Columns {
  num: number;
  space: number;
  sep: boolean;
  widths?: { w: number; space: number }[];
}

/** 各栏相对版心左边的位置和宽度（OOXML：等宽时栏宽 = (版心 − 间距 × (栏数 − 1)) / 栏数） */
export function columnBoxes(s: PageSetup): { x: number; w: number }[] {
  const c = s.cols;
  const tw = textWidth(s);
  if (!c || c.num <= 1) return [{ x: 0, w: tw }];
  if (c.widths && c.widths.length === c.num) {
    let x = 0;
    return c.widths.map((col) => {
      const box = { x, w: col.w };
      x += col.w + col.space;
      return box;
    });
  }
  const w = Math.max(10, (tw - c.space * (c.num - 1)) / c.num);
  return Array.from({ length: c.num }, (_, i) => ({ x: i * (w + c.space), w }));
}

/** 各节可以不同的页面设置：纸张与页边距（文档网格、制表位等全文统一） */
export type PageGeom = Pick<PageSetup, 'width' | 'height' | 'marginTop' | 'marginBottom' | 'marginLeft' | 'marginRight' | 'headerDistance' | 'footerDistance'>;

export const PAGE_GEOM_KEYS: (keyof PageGeom)[] = ['width', 'height', 'marginTop', 'marginBottom', 'marginLeft', 'marginRight', 'headerDistance', 'footerDistance'];

export function pageGeom(s: PageSetup | PageGeom): PageGeom {
  return Object.fromEntries(PAGE_GEOM_KEYS.map((k) => [k, s[k]])) as PageGeom;
}

export function sameGeom(a: PageGeom, b: PageGeom): boolean {
  return PAGE_GEOM_KEYS.every((k) => Math.abs(a[k] - b[k]) < 0.01);
}

/** 页眉页脚分区：默认（奇数页）、首页、偶数页 */
export type HFStory = 'header' | 'footer' | 'headerFirst' | 'footerFirst' | 'headerEven' | 'footerEven';
export type Story = 'body' | HFStory;

export const HF_STORIES: HFStory[] = ['header', 'footer', 'headerFirst', 'footerFirst', 'headerEven', 'footerEven'];

export const STORY_LABEL: Record<Story, string> = {
  body: '正文',
  header: '页眉',
  footer: '页脚',
  headerFirst: '首页页眉',
  footerFirst: '首页页脚',
  headerEven: '偶数页页眉',
  footerEven: '偶数页页脚',
};

export function isHeaderStory(s: Story): boolean {
  return s === 'header' || s === 'headerFirst' || s === 'headerEven';
}

/**
 * 一节的页眉页脚。“与上一节相同”的分区与上一节共用同一个数组对象（撤销快照 structuredClone 会保留这种共用），
 * 编辑时两节一起改变，保存时只写一份。
 */
export interface SectionHF extends Record<HFStory, Paragraph[]> {
  /** 本节首页不同 */
  titlePg: boolean;
}

/**
 * 文档的一节：index 为结束该节的分节段落下标，owner 为保存其页眉页脚的分节段落下标（该段没有 sectionHF 时为下一节的），
 * 最后一节（页眉页脚、页面设置在 Doc 上）两者都为 null。setup 为该节实际使用的页面设置。
 */
export interface SectionInfo {
  index: number | null;
  owner: number | null;
  hf: SectionHF;
  setup: PageSetup;
}

/** 最后一节（正文末尾 sectPr）的页眉页脚，数组即 Doc 上的数组。 */
export function lastSectionHF(doc: Doc): SectionHF {
  return {
    header: doc.header,
    footer: doc.footer,
    headerFirst: doc.headerFirst,
    footerFirst: doc.footerFirst,
    headerEven: doc.headerEven,
    footerEven: doc.footerEven,
    titlePg: doc.setup.titlePg,
  };
}

/** 各节按文档顺序（每个正文顶层分节段落结束一节，最后一节到文末）。没有 sectionHF 的分节段落沿用下一节的。 */
export function sectionList(doc: Doc): SectionInfo[] {
  const out: SectionInfo[] = [{ index: null, owner: null, hf: lastSectionHF(doc), setup: doc.setup }];
  for (let i = doc.body.length - 1; i >= 0; i--) {
    const b = doc.body[i];
    if (b.kind !== 'p' || !b.props.sectionBreak) continue;
    const next = out[0];
    out.unshift({
      index: i,
      owner: b.sectionHF ? i : next.owner,
      hf: b.sectionHF ?? next.hf,
      setup: b.sectionPage || b.sectionCols || doc.setup.cols ? { ...doc.setup, ...b.sectionPage, cols: b.sectionCols } : doc.setup,
    });
  }
  return out;
}

/** 页眉页脚段落的路径：最后一节为 [段落]，其它节为 [分节段落下标, 段落]。 */
export function hfPath(owner: number | null, index: number): number[] {
  return owner === null ? [index] : [owner, index];
}

/** 路径所在的页眉页脚段落数组。 */
export function hfList(doc: Doc, story: HFStory, path: number[]): Paragraph[] {
  if (path.length < 2) return doc[story];
  const hf = (doc.body[path[0]] as Paragraph).sectionHF;
  if (!hf) throw new Error('光标位置无效');
  return hf[story];
}

/** 某页实际使用的页眉或页脚分区：first 为该页是本节首页，pageIndex 从 0 起。 */
export function hfStoryFor(titlePg: boolean, evenAndOdd: boolean, first: boolean, pageIndex: number, kind: 'header' | 'footer'): HFStory {
  if (titlePg && first) return kind === 'header' ? 'headerFirst' : 'footerFirst';
  if (evenAndOdd && (pageIndex + 1) % 2 === 0) return kind === 'header' ? 'headerEven' : 'footerEven';
  return kind;
}

export interface Doc {
  setup: PageSetup;
  /** 文档默认字符格式（决定字符网格宽度） */
  defaults: RunProps;
  body: Block[];
  header: Paragraph[];
  footer: Paragraph[];
  headerFirst: Paragraph[];
  footerFirst: Paragraph[];
  headerEven: Paragraph[];
  footerEven: Paragraph[];
  /** 自动编号列表定义 */
  lists: NumList[];
  /** 打开时遇到但不支持、保存会丢失的内容说明 */
  lossy: string[];
  /** 批注（正文里的） */
  comments?: Comment[];
  /** 脚注正文（footnotes.xml，按 id），只读显示在引用所在页的页底 */
  footnotes?: Record<string, Paragraph[]>;
  /** 文字水印（Word 放在默认页眉里的艺术字，所有页都显示）；没有为 undefined */
  watermark?: Watermark;
  /** 文档保护（限制编辑）：settings.xml 的 w:documentProtection */
  protection?: DocProtection;
  /** 修订（审阅状态）：settings.xml 的 w:trackRevisions，打开文件时修订开关跟着它，保存时写回 */
  trackRevisions?: boolean;
  /** 填过数据（模板与批量生成）：脚注 / 尾注、页眉页脚表格的原始 XML 保存时按这份数据替换 {{字段}} */
  xmlFill?: { values: Record<string, string>; blank: boolean };
  /** 本次编辑中新建或修改的样式（保存时写进 styles.xml）；key 是样式 id */
  styleEdits?: Record<string, StyleEdit>;
  /** 本次编辑中删掉的文档原有样式（保存时从 styles.xml 去掉）；parent 是用到它的段落改用的样式 */
  styleDeletes?: { id: string; parent: string | null }[];
}

/** 文字水印：与 docx-engine 的 WatermarkSpec 一致 */
export interface TextWatermark {
  text: string;
  fontFamily?: string;
  /** 'RRGGBB'，缺省银色 */
  colorHex?: string;
  /** 0–1，缺省 0.5 */
  opacity?: number;
  /** 斜 45 度（缺省）；false 为水平 */
  diagonal?: boolean;
  bold?: boolean;
  italic?: boolean;
}

/** 图片水印（Word 的“图片水印”）：图片在 images.ts 的仓库里 */
export interface PictureWatermark {
  image: { id: string; /** 原始大小（像素） */ widthPx: number; heightPx: number };
  /** 缩放百分比（相对原始大小）；没有 = 缩到页面版心里 */
  scale?: number;
  /** 冲蚀（淡化），Word 默认开 */
  washout?: boolean;
}

export type Watermark = TextWatermark | PictureWatermark;

export const isPictureWatermark = (w: Watermark): w is PictureWatermark => 'image' in w;

/**
 * 样式的修改：只记本次新建 / 改过的项（没写的继承 basedOn 或原样式）。
 * 段落格式和字符格式用模型里的字段，保存时转成 styles.xml 的 w:pPr / w:rPr。
 */
export interface StyleEdit {
  id: string;
  type: 'paragraph' | 'character';
  /** 显示名；改已有样式时不写（保留原名） */
  name?: string;
  /** 新建的样式基于哪个样式 id；改已有样式时不写 */
  basedOn?: string;
  para?: Partial<Pick<ParaProps, 'align' | 'indentLeft' | 'indentRight' | 'firstLine' | 'spaceBefore' | 'spaceAfter' | 'line' | 'keepNext' | 'keepLines' | 'pageBreakBefore' | 'outline'>>;
  run?: Partial<Pick<RunProps, 'fontEA' | 'fontLatin' | 'size' | 'bold' | 'italic' | 'underline' | 'color' | 'strike'>>;
}

/** 中文字号 → 磅值 */
export const CN_SIZES: [string, number][] = [
  ['初号', 42],
  ['小初', 36],
  ['一号', 26],
  ['小一', 24],
  ['二号', 22],
  ['小二', 18],
  ['三号', 16],
  ['小三', 15],
  ['四号', 14],
  ['小四', 12],
  ['五号', 10.5],
  ['小五', 9],
  ['六号', 7.5],
  ['小六', 6.5],
  ['七号', 5.5],
  ['八号', 5],
];

export const PAPERS: { name: string; width: number; height: number }[] = [
  { name: 'A4', width: 595.3, height: 841.9 },
  { name: 'A3', width: 841.9, height: 1190.6 },
  { name: 'A5', width: 419.6, height: 595.3 },
  { name: 'B5 (JIS)', width: 515.9, height: 728.5 },
  { name: '16 开', width: 521.6, height: 737 },
  { name: 'Letter', width: 612, height: 792 },
];

export function defaultRunProps(): RunProps {
  return { fontEA: '宋体', fontLatin: 'Times New Roman', size: 10.5, bold: false, italic: false, underline: false, color: 'auto' };
}

export function defaultParaProps(): ParaProps {
  return {
    align: 'both',
    indentLeft: 0,
    indentRight: 0,
    firstLine: 0,
    firstLineChars: null,
    spaceBefore: 0,
    spaceAfter: 0,
    spaceBeforeLines: null,
    spaceAfterLines: null,
    line: { rule: 'auto', value: 1 },
    snapToGrid: true,
    keepNext: false,
    keepLines: false,
    // Word 中文模板的“正文”样式关闭了孤行控制
    widowControl: false,
    pageBreakBefore: false,
    num: null,
    sectionBreak: null,
  };
}

/** Word 中文默认页面：A4，上下 2.54 cm，左右 3.17 cm，每页 44 行（行网格 15.6 磅）。 */
export function defaultSetup(): PageSetup {
  return {
    width: 595.3,
    height: 841.9,
    marginTop: 72,
    marginBottom: 72,
    marginLeft: 90,
    marginRight: 90,
    headerDistance: 42.55,
    footerDistance: 49.6,
    grid: { type: 'lines', linePitch: 15.6, charSpace: 0 },
    defaultTab: 21,
    titlePg: false,
    evenAndOdd: false,
  };
}

export function newParagraph(text = '', props?: Partial<ParaProps>, run?: RunProps): Paragraph {
  const mark = run ?? defaultRunProps();
  return {
    kind: 'p',
    props: { ...defaultParaProps(), ...props },
    runs: text ? [{ text, props: { ...mark } }] : [],
    mark: { ...mark },
  };
}

/** 页眉页脚的默认空段落：居中、小五号 */
export function newHFParagraph(): Paragraph {
  return newParagraph('', { align: 'center' }, { ...defaultRunProps(), size: 9 });
}

export function newDocument(): Doc {
  return {
    setup: defaultSetup(),
    defaults: defaultRunProps(),
    body: [newParagraph()],
    header: [newHFParagraph()],
    footer: [newHFParagraph()],
    headerFirst: [newHFParagraph()],
    footerFirst: [newHFParagraph()],
    headerEven: [newHFParagraph()],
    footerEven: [newHFParagraph()],
    lists: [],
    lossy: [],
  };
}

export function newTable(cols: number, rows: number, width: number, run: RunProps): Table {
  const w = width / cols;
  return {
    kind: 'table',
    cols: Array.from({ length: cols }, () => w),
    rows: Array.from({ length: rows }, () => ({
      height: 0,
      heightRule: 'auto' as const,
      header: false,
      cells: Array.from({ length: cols }, () => ({
        gridSpan: 1,
        vMerge: 'none' as const,
        vAlign: 'top' as const,
        paragraphs: [newParagraph('', { align: 'left', firstLineChars: null }, run)],
      })),
    })),
    align: 'center',
    indent: 0,
    borders: true,
    cellMargin: { left: 5.4, right: 5.4, top: 0, bottom: 0 },
  };
}

export function textWidth(s: PageSetup): number {
  return s.width - s.marginLeft - s.marginRight;
}

export function textHeight(s: PageSetup): number {
  return s.height - s.marginTop - s.marginBottom;
}

/** 字符网格间距：默认字号加 charSpace。 */
export function charPitch(doc: Doc): number {
  return doc.defaults.size + doc.setup.grid.charSpace / 4096;
}

export function paraText(p: Paragraph): string {
  return p.runs.map((r) => r.text).join('');
}

export function sameProps(a: RunProps, b: RunProps): boolean {
  return (
    a.fontEA === b.fontEA &&
    a.fontLatin === b.fontLatin &&
    a.size === b.size &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.color === b.color &&
    !!a.strike === !!b.strike &&
    a.vertAlign === b.vertAlign &&
    (a.spacing ?? 0) === (b.spacing ?? 0) &&
    a.highlight === b.highlight &&
    a.rev?.kind === b.rev?.kind &&
    a.rev?.author === b.rev?.author &&
    a.rev?.date === b.rev?.date &&
    a.link?.rid === b.link?.rid &&
    a.link?.anchor === b.link?.anchor &&
    a.link?.url === b.link?.url &&
    a.rStyle === b.rStyle
  );
}

/** 正文（含表格）里用到的图片 id */
export function imageIds(doc: Doc): string[] {
  const out: string[] = [];
  const para = (p: Paragraph) =>
    p.runs.forEach((r) => {
      if (r.image) out.push(r.image.id);
      if (r.box?.image) out.push(r.box.image.id);
      r.box?.paragraphs.forEach(para);
    });
  const table = (t: Table) => t.rows.forEach((row) => row.cells.forEach((c) => c.paragraphs.forEach(para)));
  for (const b of doc.body) {
    if (b.kind === 'p') para(b);
    else table(b);
  }
  // 各节的页眉页脚（含页眉页脚里的表格）
  const seen = new Set<Paragraph[]>();
  for (const s of sectionList(doc)) {
    for (const story of HF_STORIES) {
      const ps = s.hf[story];
      if (seen.has(ps)) continue;
      seen.add(ps);
      for (const p of ps) {
        para(p);
        if (p.hfTable) table(p.hfTable.table);
      }
    }
  }
  if (doc.watermark && isPictureWatermark(doc.watermark)) out.push(doc.watermark.image.id);
  return [...new Set(out)];
}

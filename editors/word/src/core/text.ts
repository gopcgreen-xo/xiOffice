/** 中文排版的字符分类与断行规则（参照 Word 中文版式的默认设置）。 */

/** 行首禁则：不能出现在行首的字符（Word 默认“标准”禁则集合）。 */
export const NO_LINE_START = new Set(
  Array.from(
    '!%),.:;?]}¢°·’"”†‡›℃∶、。〃〆〕〗〞﹚﹜！＂％＇），．：；？］｝～–—…‧〉》」』】〙〛゛゜ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ・ー',
  ),
);

/** 行尾禁则：不能出现在行尾的字符。 */
export const NO_LINE_END = new Set(Array.from('$(£¥‘“〈《「『【〔〖〝﹙﹛＄（［｛￡￥'));

/** 允许“标点溢出边界”（悬挂在右边距外）的标点。 */
export const HANGING = new Set(Array.from('，。、．,.：；！？'));

/** 按中文字体排版的字符（汉字、全角符号、日韩文字，以及中文语境下的引号、省略号、破折号）。 */
export function isCJK(ch: string): boolean {
  const cp = ch.codePointAt(0)!;
  return (
    (cp >= 0x2e80 && cp <= 0x9fff) || // 部首、符号、假名、注音、CJK 扩展 A、基本汉字
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xac00 && cp <= 0xd7af) || // 韩文
    (cp >= 0xf900 && cp <= 0xfaff) || // 兼容汉字
    (cp >= 0xfe30 && cp <= 0xfe4f) || // 竖排兼容标点
    (cp >= 0xff00 && cp <= 0xffef) || // 全角字符
    (cp >= 0x20000 && cp <= 0x3134f) || // 扩展 B 及以后
    cp === 0x2018 ||
    cp === 0x2019 ||
    cp === 0x201c ||
    cp === 0x201d ||
    cp === 0x2026 ||
    cp === 0x2014 ||
    cp === 0x00b7
  );
}

/** 中西文之间自动加间距的西文字符：字母和数字。 */
export function isLatinAlnum(ch: string): boolean {
  return /^[A-Za-z0-9À-ɏ]$/.test(ch);
}

/** 在 prev 与 next 之间能否断行。 */
export function canBreakBetween(prev: string, next: string): boolean {
  if (NO_LINE_START.has(next) || NO_LINE_END.has(prev)) return false;
  if (prev === ' ' || prev === '\t') return true;
  if (next === ' ') return false; // 空格随前一行
  if (isCJK(prev) || isCJK(next)) return true;
  if (prev === '-' && /[A-Za-z]/.test(next)) return true;
  return false;
}

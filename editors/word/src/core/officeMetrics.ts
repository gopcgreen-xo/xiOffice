/**
 * 常用 Office 字体的度量（只是数字：上伸、下伸、外部行距，单位 em；ea = 中日韩字体），给没有装这些字体的地方用——
 * 安卓手机版（docs/ANDROID.md §9）：手机上没有宋体、黑体、Times New Roman、Calibri……，按这里的数字算行高，
 * 分页、行距就和电脑上一致（字形用手机自己的替代字体画）。
 *
 * 数字由 Windows 自带字体文件经本引擎的 faceInfo 算出（和桌面版读系统字体得到的一样）；不含任何字形或字体文件。
 * 登记顺序：先登记手机系统字体，再登记这张表（registerFonts 同名先到先得，真有这个字体时用真字体）。
 */
import { registerFonts, type FaceInfo } from './fontfile';

const ROWS: [string[], number, number, number, boolean][] = [
  [['SimSun', '宋体'], 0.8594, 0.1406, 0.1406, true],
  [['NSimSun', '新宋体'], 0.8594, 0.1406, 0.1406, true],
  [['SimHei', '黑体'], 0.8594, 0.1406, 0.1406, true],
  [['FangSong', '仿宋'], 0.8594, 0.1406, 0.1406, true],
  [['KaiTi', '楷体'], 0.8594, 0.1406, 0.1406, true],
  [['Microsoft YaHei', '微软雅黑'], 1.0581, 0.2617, 0, true],
  [['Microsoft YaHei UI'], 1.0156, 0.2544, 0, true],
  [['DengXian', '等线'], 0.8101, 0.2319, 0, true],
  [['Times New Roman'], 0.8911, 0.2163, 0.0425, false],
  [['Arial'], 0.9053, 0.2119, 0.0327, false],
  [['Calibri'], 0.9521, 0.2686, 0, false],
  [['Cambria'], 0.9502, 0.2222, 0, false],
  [['Cambria Math'], 3.1167, 2.4634, 0, false],
  [['Courier New'], 0.8325, 0.3003, 0, false],
  [['Verdana'], 1.0054, 0.21, 0, false],
  [['Tahoma'], 1.0005, 0.2065, 0, false],
  [['Georgia'], 0.917, 0.2192, 0, false],
  [['Segoe UI'], 1.0791, 0.251, 0, false],
  [['Consolas'], 0.9199, 0.251, 0, false],
  [['Symbol'], 1.0054, 0.2197, 0, false],
  [['Wingdings'], 0.8989, 0.2109, 0, false],
  [['STXihei', '华文细黑'], 0.871, 0.208, 0.298, true],
  [['STSong', '华文宋体'], 0.86, 0.26, 0.181, true],
  [['STKaiti', '华文楷体'], 0.86, 0.26, 0.181, true],
  [['STFangsong', '华文仿宋'], 0.86, 0.26, 0.181, true],
  [['FZShuTi', '方正舒体'], 0.9102, 0.1406, 0.0039, true],
  [['SimSun-ExtB'], 0.8594, 0.1406, 0, true],
];

export const OFFICE_FACES: FaceInfo[] = ROWS.map(([families, ascent, descent, lineGap, ea]) => ({ families, metrics: { ascent, descent, lineGap, ea } }));

/** 补登记常用 Office 字体的度量（已登记的同名字体不受影响） */
export function registerOfficeFallback(): void {
  registerFonts(OFFICE_FACES);
}

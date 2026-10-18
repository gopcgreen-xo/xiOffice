import { describe, expect, it } from 'vitest';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { writeXlsx } from '@wz/editor-shared/xlsx';
import { applyTint, DEFAULT_INDEXED, DEFAULT_THEME, readPalette, resolveBorder, resolveColor, resolveFill, type Palette } from './color';
import { readWorkbook, styleAt } from './workbook';
import { borderEdges, borderLine, cellBorder, cellColors } from './grid';

const PAL: Palette = { indexed: [...DEFAULT_INDEXED], theme: [...DEFAULT_THEME] };

/** 两个 #rrggbb / RRGGBB 每个通道相差不超过 2（Excel 色板的值本身有取整误差） */
function near(a: string, b: string) {
  const ch = (s: string) => [0, 2, 4].map((i) => parseInt(s.replace('#', '').slice(i, i + 2), 16));
  const [x, y] = [ch(a), ch(b)];
  expect(x.every((v, i) => Math.abs(v - y[i]) <= 2), `${a} ≈ ${b}`).toBe(true);
}

describe('color', () => {
  it('resolves rgb / indexed / theme / auto', () => {
    expect(resolveColor('<color rgb="FFFF0000"/>', PAL)).toBe('#ff0000');
    expect(resolveColor('<color rgb="00FF00"/>', PAL)).toBe('#00ff00');
    expect(resolveColor('<color indexed="10"/>', PAL)).toBe('#ff0000');
    expect(resolveColor('<color indexed="64"/>', PAL)).toBe('#000000');
    expect(resolveColor('<color indexed="65"/>', PAL)).toBe('#ffffff');
    // theme 序号：0 = lt1，1 = dk1，4 = accent1
    expect(resolveColor('<color theme="0"/>', PAL)).toBe('#ffffff');
    expect(resolveColor('<color theme="1"/>', PAL)).toBe('#000000');
    expect(resolveColor('<color theme="4"/>', PAL)).toBe('#4472c4');
    expect(resolveColor('<color auto="1"/>', PAL)).toBeUndefined();
    expect(resolveColor(undefined, PAL)).toBeUndefined();
  });

  it('applies tint in HLS luminance', () => {
    // Excel 色板上的值：白色深 15%、蓝色个性色 1 浅 80%、个性色 1 深 25%
    // tint 用 Excel 写进文件的值
    near(applyTint('FFFFFF', -0.1499984740745262), 'D9D9D9');
    near(applyTint('000000', 0.499984740745262), '7F7F7F');
    near(applyTint('4472C4', 0.7999816888943144), 'D9E1F2');
    near(applyTint('4472C4', -0.249977111117893), '305496');
    near(applyTint('ED7D31', 0.7999816888943144), 'FCE4D6');
    expect(resolveColor('<color theme="0" tint="-0.14999847407452621"/>', PAL)).toBe('#d9d9d9');
  });

  it('reads indexedColors and the theme colour scheme', () => {
    const pal = readPalette(
      '<styleSheet><colors><indexedColors><rgbColor rgb="FF112233"/><rgbColor rgb="FF445566"/></indexedColors></colors></styleSheet>',
      '<a:theme><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1>' +
        '<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="ABCDEF"/></a:accent1></a:clrScheme></a:themeElements></a:theme>',
    );
    expect(pal.indexed.slice(0, 3)).toEqual(['112233', '445566', 'FF0000']);
    expect(pal.theme.slice(0, 5)).toEqual(['FEFEFE', '111111', 'EEEEEE', '222222', 'ABCDEF']);
  });

  it('resolves fills: none, solid, patterns, gradients', () => {
    expect(resolveFill('<patternFill patternType="none"/>', PAL)).toBeUndefined();
    expect(resolveFill('<patternFill/>', PAL)).toBeUndefined();
    expect(resolveFill('<patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill>', PAL)).toBe('#ffff00');
    // solid 没写前景色：系统前景色（黑）
    expect(resolveFill('<patternFill patternType="solid"/>', PAL)).toBe('#000000');
    // 50% 灰图案：前景黑、背景白各一半
    expect(resolveFill('<patternFill patternType="mediumGray"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill>', PAL)).toBe('#808080');
    expect(resolveFill('<gradientFill><stop position="0"><color rgb="FF000000"/></stop><stop position="1"><color rgb="FFFFFFFF"/></stop></gradientFill>', PAL)).toBe('#808080');
  });

  it('resolves borders: styles, colours, start / end, none', () => {
    expect(resolveBorder('<left/><right/><top/><bottom/><diagonal/>', PAL)).toBeUndefined();
    expect(
      resolveBorder('<left style="thin"><color indexed="64"/></left><right style="medium"><color rgb="FFFF0000"/></right><top style="none"/><bottom style="double"/>', PAL),
    ).toEqual({ l: { w: 1, style: 'solid', color: '#000000' }, r: { w: 2, style: 'solid', color: '#ff0000' }, b: { w: 3, style: 'double' } });
    expect(resolveBorder('<start style="dashed"/><end style="hair"><color auto="1"/></end>', PAL)).toEqual({ l: { w: 1, style: 'dashed' }, r: { w: 1, style: 'dotted' } });
  });

  it('draws only the outer edges of a merged area', () => {
    const all = { w: 1, style: 'solid' as const };
    const st = { border: { l: all, r: all, t: all, b: all } };
    const m = { r0: 0, c0: 0, r1: 1, c1: 2 };
    expect(cellBorder(st, 0, 0, m)).toEqual({ l: all, t: all });
    expect(cellBorder(st, 0, 1, m)).toEqual({ t: all });
    expect(cellBorder(st, 1, 2, m)).toEqual({ r: all, b: all });
    expect(cellBorder({ border: { r: all } }, 0, 0, m)).toBeUndefined();
    expect(cellBorder(st, 5, 5)).toBe(st.border);
    expect(cellBorder({}, 0, 0)).toBeUndefined();
  });

  it('border line: auto / black follows the theme text colour unless the cell is filled', () => {
    expect(borderLine({ w: 1, style: 'solid' })).toBe('1px solid var(--text)');
    expect(borderLine({ w: 2, style: 'dashed', color: '#ff0000' })).toBe('2px dashed #ff0000');
    expect(borderLine({ w: 1, style: 'solid', color: '#000000' }, '#ffff00')).toBe('1px solid #000000');
  });

  it('a shared edge becomes one line segment; the thicker side wins', () => {
    const thin = { w: 1, style: 'solid' as const };
    const thick = { w: 3, style: 'solid' as const };
    // A1 右边线与 B1 左边线是同一条；A1 下边线与 A2 上边线是同一条
    const edges = borderEdges([
      { r: 0, c: 0, b: { r: thin, b: thin } },
      { r: 0, c: 1, b: { l: thin } },
      { r: 1, c: 0, b: { t: thick } },
    ]);
    expect(edges.map((e) => `${e.v ? 'v' : 'h'}${e.r},${e.c}:${e.w}`).sort()).toEqual(['h1,0:3', 'v0,1:1']);
  });

  it('colours a cell like Excel', () => {
    expect(cellColors(undefined)).toEqual({});
    // 无底色：黑字按自动色处理，其它颜色照原样
    expect(cellColors({ color: '#000000' })).toEqual({});
    expect(cellColors({ color: '#ff0000' })).toEqual({ color: '#ff0000' });
    // 有底色：盖住网格线，自动色文字为黑
    expect(cellColors({ fill: '#ffff00' })).toEqual({ background: '#ffff00', borderColor: '#ffff00', color: '#000000' });
    expect(cellColors({ fill: '#ffff00', color: '#0000ff' }, true).background).toContain('#ffff00');
  });
});

/** 用 writeXlsx 生成骨架，再替换样式、工作表，可加主题 */
function book(sheetXml: string, styles: string, theme?: string): Uint8Array {
  const files = unzipSync(writeXlsx([{ name: 'S', rows: [['x']] }]));
  files['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet>${sheetXml}</worksheet>`);
  files['xl/styles.xml'] = strToU8(styles);
  if (theme) {
    files['xl/theme/theme1.xml'] = strToU8(theme);
    const rels = new TextDecoder()
      .decode(files['xl/_rels/workbook.xml.rels'])
      .replace('</Relationships>', '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>');
    files['xl/_rels/workbook.xml.rels'] = strToU8(rels);
  }
  return zipSync(files);
}

describe('readWorkbook colours', () => {
  const STYLES = `<styleSheet>
<fonts count="2"><font><sz val="11"/><color theme="1"/></font><font><color rgb="FFFF0000"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor theme="4"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders>
<cellXfs count="4"><xf fontId="0" fillId="0"/><xf fontId="1" fillId="0"/><xf fontId="0" fillId="2" borderId="1"/><xf fontId="0" fillId="3"/></cellXfs></styleSheet>`;

  it('reads fill and font colours with the document theme', () => {
    const wb = readWorkbook(
      book(
        '<sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>红字</t></is></c><c r="B1" s="2" t="inlineStr"><is><t>底色</t></is></c></row></sheetData>',
        STYLES,
        '<a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="t"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="00B050"/></a:accent1></a:clrScheme></a:themeElements></a:theme>',
      ),
    );
    const s = wb.sheets[0];
    expect(wb.styles[styleAt(s, 0, 0)!]).toEqual({ color: '#ff0000' });
    // 文档主题的 accent1 覆盖 Office 默认的 4472C4
    expect(wb.styles[styleAt(s, 0, 1)!]).toMatchObject({ fill: '#00b050', color: '#000000' });
    // 有底色的格子照样有自己的边框
    const thin = { w: 1, style: 'solid', color: '#000000' };
    expect(wb.styles[styleAt(s, 0, 1)!].border).toEqual({ l: thin, r: thin, t: thin, b: thin });
  });

  it('empty cells, row and column formats carry fills (cell > row > column)', () => {
    const s = readWorkbook(
      book(
        '<cols><col min="2" max="3" width="9" style="3"/></cols>' +
          '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="D1" s="2"/></row>' +
          '<row r="2" s="2" customFormat="1"><c r="C2"/></row><row r="3" s="3"/></sheetData>',
        STYLES,
      ),
    ).sheets[0];
    expect(styleAt(s, 0, 3)).toBe(2); // 空格子自己的样式
    expect(styleAt(s, 0, 1)).toBe(3); // 没有 <c>：整列格式
    expect(styleAt(s, 1, 0)).toBe(2); // 整行格式优先于整列
    expect(styleAt(s, 1, 1)).toBe(2);
    expect(styleAt(s, 1, 2)).toBe(0); // 有 <c> 但没写 s：缺省样式，盖过整行 / 整列
    expect(styleAt(s, 2, 1)).toBe(3); // 行没有 customFormat：不算整行格式
    expect(styleAt(s, 0, 0)).toBeUndefined();
    expect(s.cells[0][3]).toBeUndefined(); // 空格子不算数据
  });
});

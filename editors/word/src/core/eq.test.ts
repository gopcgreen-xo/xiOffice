// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { parseEq } from './eq';
import { layoutDocument, type Measurer } from './layout';
import { newDocument, type Paragraph } from './model';
import { readDocx, writeDocx } from '../ooxml/docx';
import { isCJK } from './text';

const fake: Measurer = { width: (ch) => (isCJK(ch) || ch === '□' || ch === '√' ? 1 : 0.5), metrics: () => ({ ascent: 0.88, descent: 0.12 }) };
const d = newDocument();
const big = { ...d.defaults, size: 12 };
const small = { ...d.defaults, size: 8 };
const chars = (s: string, props = big) => [...s].map((ch) => ({ ch, props }));

describe('EQ 域', () => {
  it('\\o 重叠：各部分叠在一起，\\ac 居中；不是 EQ 域返回 null', () => {
    const e = parseEq([...chars('eq \\o\\ac(□,'), ...chars('√', small), ...chars(')')])!;
    expect(e.overstrike).toBe(true);
    expect(e.align).toBe('center');
    expect(e.parts.map((p) => p.map((s) => `${s.text}@${s.props.size}`))).toEqual([['□@12'], ['√@8']]);
    expect(parseEq(chars(' PAGE '))).toBeNull();
  });

  it('没有 \\o 时参数依次排开，嵌套开关不显示', () => {
    const e = parseEq(chars('EQ \\s\\up8(2)x'))!;
    expect(e.overstrike).toBe(false);
    expect(e.parts[0].map((s) => s.text).join('')).toBe('2');
  });

  // 带圈字符的域：begin / instrText（√ 是小一号的单独一段）/ end，没有 separate 和缓存结果
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const r = (inner: string, sz = 24) => `<w:r><w:rPr><w:sz w:val="${sz}"/></w:rPr>${inner}</w:r>`;
  const body =
    `<w:p>${r('<w:t>本合同由</w:t>')}${r('<w:fldChar w:fldCharType="begin"/>')}${r('<w:instrText>eq \\o\\ac(□,</w:instrText>')}` +
    `${r('<w:instrText>√</w:instrText>', 16)}${r('<w:instrText>)</w:instrText>')}${r('<w:fldChar w:fldCharType="end"/>')}${r('<w:t>双方</w:t>')}</w:p>`;
  const docx = () => {
    const base = unzipSync(writeDocx(newDocument()));
    return zipSync({ ...base, 'word/document.xml': strToU8(`<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`) });
  };

  it('没有显示结果的域读成原样对象（保存时域代码原样写回，不再丢失），EQ 宽度取最宽的部分', () => {
    const doc = readDocx(docx());
    const p = doc.body[0] as Paragraph;
    const o = p.runs.find((x) => x.opaque)!.opaque!;
    expect(o.level).toBe('para');
    expect(o.eq?.overstrike).toBe(true);
    const line = layoutDocument(doc, fake).pages[0].body[0].line;
    const g = line.glyphs.find((x) => x.eq)!;
    expect(g.w).toBeCloseTo(12); // □ 12 磅，√ 8 磅：取 12
    expect(line.glyphs.find((x) => x.ch === '双')!.x - g.x).toBeCloseTo(12);
    // 编辑这一段后整篇写出：域代码仍在
    p.runs.push({ text: '。', props: p.runs[0].props });
    const xml = strFromU8(unzipSync(writeDocx(doc))['word/document.xml']);
    expect(xml).toContain('<w:fldChar w:fldCharType="begin"/>');
    expect(xml).toContain('eq \\o\\ac(□,');
    expect(xml).toContain('√');
  });
});

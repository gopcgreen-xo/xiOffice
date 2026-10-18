// @vitest-environment jsdom
/**
 * .doc（Word 97-2003）读取测试。样本是 Microsoft Word 把 samples/ 里的 .docx 另存成的 .doc（scripts/word-save-doc.ps1），
 * 同一份内容读 .docx 和读 .doc 应该得到相同的文字、字号、表格结构和纸张设置；
 * 还没有支持的内容（脚注、批注、对象…）读成空白并写进 doc.lossy。
 */
import { readdirSync, readFileSync } from 'fs';
import { join } from 'path';
import * as CFB from 'cfb';
import { describe, expect, it } from 'vitest';
import { readDocx, writeDocx } from '../ooxml/docx';
import { FIELD_CHAR, PAGE_GEOM_KEYS, imageIds, paraText, type Block, type Doc } from '../core/model';
import { imageData } from '../core/images';
import { computeLabels } from '../core/numbering';
import { isDoc, papOf, readDoc } from './read';
import { parseGrpprl } from './sprm';

const SAMPLES = join(__dirname, '../../samples');
const DOCS = readdirSync(join(SAMPLES, 'doc')).filter((f) => f.endsWith('.doc'));
const bytes = (p: string) => new Uint8Array(readFileSync(p));

/** 文字（去掉占位符）、字号、粗体、对齐：Word 读出来应该一致的部分 */
function summary(d: Doc): string[] {
  const para = (p: Extract<Block, { kind: 'p' }>) =>
    `${p.props.align}${p.props.style ?? ''}|${p.runs
      .filter((r) => r.text !== FIELD_CHAR)
      .map((r) => `${r.props.bold ? '*' : ''}${r.props.size}:${r.text}`)
      .join('')}`;
  return d.body
    .map((b) => (b.kind === 'p' ? para(b) : `T${b.cols.map(Math.round).join(',')}:${b.rows.map((r) => r.cells.map((c) => `${c.gridSpan}${c.vMerge}${c.paragraphs.map(paraText).join('/')}`).join(';')).join('//')}`))
    // docx 里书签 / 公式 / 脚注引用只有占位符的段落，.doc 里读成空段落：两边都当作空
    .map((s) => s.replace(/\|$/, '|'));
}

describe.each(DOCS)('.doc 样本 %s', (name) => {
  const doc = readDoc(bytes(join(SAMPLES, 'doc', name)));
  const ref = readDocx(bytes(join(SAMPLES, name.replace(/\.doc$/, '.docx'))));

  it('是 OLE 复合文件', () => expect(isDoc(bytes(join(SAMPLES, 'doc', name)))).toBe(true));

  it('文字、字号、粗体、对齐、表格结构和 .docx 一致', () => {
    expect(summary(doc)).toEqual(summary(ref));
  });

  it('纸张、页边距、分栏、文档网格一致', () => {
    for (const k of PAGE_GEOM_KEYS) expect(doc.setup[k]).toBeCloseTo(ref.setup[k], 1);
    expect(doc.setup.grid.type).toBe(ref.setup.grid.type);
    expect(doc.setup.grid.linePitch).toBeCloseTo(ref.setup.grid.linePitch, 1);
    expect(doc.setup.grid.charSpace).toBe(ref.setup.grid.charSpace);
    expect(doc.setup.cols?.num).toBe(ref.setup.cols?.num);
  });

  it('“手动换行的行不拉开”兼容选项与 .docx 一致（DOP Copts60.fExpShRtn）', () => {
    expect(!!doc.setup.noExpandShiftReturn).toBe(!!ref.setup.noExpandShiftReturn);
  });

  it('页眉页脚文字一致', () => {
    const text = (ps: { runs: { text: string }[] }[]) => ps.map((p) => p.runs.map((r) => r.text).filter((t) => t !== FIELD_CHAR).join('')).join('/');
    for (const k of ['header', 'footer'] as const) {
      const table = ref[k].find((p) => p.hfTable)?.hfTable?.table;
      if (!table) expect(text(doc[k])).toBe(text(ref[k]));
      // 页眉页脚里的表格暂不能编辑：单元格里的段落按顺序排成普通段落，文字都在
      else for (const c of table.rows.flatMap((r) => r.cells)) for (const p of c.paragraphs) expect(text(doc[k])).toContain(paraText(p).split(FIELD_CHAR).join(''));
    }
  });

  it('转成 .docx 再读回，内容不变', () => {
    const again = readDocx(writeDocx(doc));
    expect(summary(again)).toEqual(summary(doc));
  });
});

describe('.doc 里的表格（9 列不等宽、整行合并、只画部分边框、底纹）', () => {
  const doc = readDoc(bytes(join(SAMPLES, 'doc', '复杂表格.doc')));
  const ref = readDocx(bytes(join(SAMPLES, '复杂表格.docx')));
  const tables = (d: Doc) => d.body.filter((b): b is Extract<Block, { kind: 'table' }> => b.kind === 'table');
  const edgeMark = (c: { edges?: object }) => (['top', 'left', 'bottom', 'right'] as const).map((k) => ((c.edges as Record<string, unknown> | undefined)?.[k] ? k[0] : '-')).join('');

  it('列宽和合并和 .docx 一致（表格定义太大时放在 Data 流里：sprmPHugePapx）', () => {
    expect(tables(doc).map((t) => t.cols.map(Math.round))).toEqual(tables(ref).map((t) => t.cols.map(Math.round)));
    expect(tables(doc)[0]!.cols).toHaveLength(9);
  });

  it('边框：表格级边框落到每个单元格，只画部分边框的表格逐边一致；底纹一致', () => {
    const [one, two] = tables(doc);
    const [refOne, refTwo] = tables(ref);
    // 第一张表：整张表有框线（.docx 里记在表格上，.doc 里记在表格定义上）→ 每个单元格四边都有
    expect(refOne!.borders).toBe(true);
    expect(one!.rows.flatMap((r) => r.cells).every((c) => edgeMark(c) === 'tlbr')).toBe(true);
    // 第二张表：只有右边一格画满、左边一格只有右线
    const marks = (t: typeof two) => t!.rows.flatMap((r) => r.cells.map(edgeMark));
    expect(marks(two)).toEqual(marks(refTwo));
    const shades = (t: typeof one) => t!.rows.flatMap((r) => r.cells.map((c) => c.shading ?? ''));
    expect(shades(one)).toEqual(shades(refOne));
    expect(shades(one).filter(Boolean)).toEqual(['CCCCCC']);
  });
});

describe('.doc 里的自动编号', () => {
  const doc = readDoc(bytes(join(SAMPLES, 'doc', '编号列表.doc')));
  const ref = readDocx(bytes(join(SAMPLES, '编号列表.docx')));
  const labels = (d: Doc) => {
    const l = computeLabels(d);
    return d.body.flatMap((b) => (b.kind === 'p' ? [`${l.get(b) ?? ''}|${paraText(b).split(FIELD_CHAR).join('')}`] : []));
  };

  it('每段的编号文字和 .docx 一致（阿拉伯、括号、中文数字、带圈数字、天干、字母、罗马、项目符号、多级、起始编号）', () => {
    expect(labels(doc)).toEqual(labels(ref));
    // 确实有编号可比（不是两边都空）
    expect(labels(doc).filter((x) => /^[^|]+\|/.test(x)).length).toBeGreaterThan(20);
  });

  it('列表级别的格式、文字模板、起始值、缩进和 .docx 一致', () => {
    for (const [i, l] of doc.lists.entries()) {
      const r = ref.lists[i]!;
      for (const k of [0, 1, 2]) {
        const a = l.levels[k]!;
        const b = r.levels[k]!;
        expect([a.fmt, a.text, a.start, a.suffix, a.indentLeft, a.firstLine]).toEqual([b.fmt, b.text, b.start, b.suffix, b.indentLeft, b.firstLine]);
      }
    }
  });

  it('不再提示“自动编号”不支持', () => {
    expect(doc.lossy.join()).not.toContain('自动编号');
  });
});

describe('.doc 里的图片', () => {
  it('内嵌的 png / jpeg 读成图片，尺寸取自文件', () => {
    const doc = readDoc(bytes(join(SAMPLES, 'doc', '页眉表格.doc')));
    const ref = readDocx(bytes(join(SAMPLES, '页眉表格.docx')));
    const ids = imageIds(doc);
    expect(ids.length).toBeGreaterThan(0);
    expect(ids.length).toBe(imageIds(ref).length);
    expect(doc.lossy.join()).not.toContain('图片');
    const img = doc.header.flatMap((p) => p.runs).find((r) => r.image)!.image!;
    expect(img.w).toBeGreaterThan(5);
    expect(img.h).toBeGreaterThan(5);
  });
});

describe('.doc 里的文本框、形状和浮动图片（Word 生成：scripts/word-make-shapes.ps1）', () => {
  const doc = readDoc(bytes(join(SAMPLES, 'doc', '文本框和形状.doc')));
  const ref = readDocx(bytes(join(SAMPLES, '文本框和形状.docx')));
  /** 正文里的绘图对象：文字只比内容，图片只比有没有（图片 id 每次登记都不同），读入的原始 XML 不比 */
  const boxes = (d: Doc) =>
    d.body.flatMap((b) =>
      b.kind === 'p'
        ? b.runs.filter((r) => r.box).map((r) => {
            const { raw, paragraphs, image, ...rest } = r.box!;
            void raw;
            return { ...rest, text: paragraphs.map(paraText), image: !!image };
          })
        : [],
    );

  it('不再提示“暂不支持”', () => expect(doc.lossy).toEqual([]));

  it('位置、大小、环绕、衬于文字下方、外形、填充、线条、文字边距、垂直对齐、文字都和 .docx 一致', () => {
    const a = boxes(doc);
    expect(a.length).toBe(10);
    expect(a).toEqual(boxes(ref));
    expect(a.map((b) => b.geom ?? (b.image ? 'picture' : 'textbox')).sort()).toEqual(
      ['diamond', 'ellipse', 'line', 'line', 'picture', 'rect', 'rect', 'roundRect', 'textbox', 'triangle'].sort(),
    );
  });

  it('浮动图片的数据在 WordDocument 流里（BSE.foDelay），和 .docx 里的图片字节相同', () => {
    const pic = (d: Doc) => d.body.flatMap((b) => (b.kind === 'p' ? b.runs : [])).find((r) => r.box?.image)!.box!.image!;
    expect(imageData(pic(doc).id)).toEqual(imageData(pic(ref).id));
  });

  it('转成 .docx 再读回，绘图对象不变', () => {
    expect(boxes(readDocx(writeDocx(doc)))).toEqual(boxes(doc));
  });
});

describe('.doc 里还不支持的内容', () => {
  it('脚注、批注在 lossy 里说明', () => {
    const doc = readDoc(bytes(join(SAMPLES, 'doc', 'Word生成的综合样本.doc')));
    expect(doc.lossy).toEqual(expect.arrayContaining(['脚注', '批注']));
  });
});

describe('.doc 文件检查', () => {
  const sample = bytes(join(SAMPLES, 'doc', '排版对照样例.doc'));
  const patched = (edit: (wd: Uint8Array) => void) => {
    const c = CFB.parse(sample);
    const e = CFB.find(c, 'WordDocument')!;
    const wd = Uint8Array.from(e.content as ArrayLike<number>);
    edit(wd);
    e.content = wd as never;
    return new Uint8Array(CFB.write(c, { type: 'array' }) as ArrayLike<number>);
  };

  it('不是 Word 文件', () => {
    expect(isDoc(new Uint8Array([80, 75, 3, 4, 0, 0, 0, 0, 0]))).toBe(false);
    expect(() => readDoc(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4, 5, 6, 7, 8]))).toThrow(/不是有效/);
  });

  it('加密的文档给出明确提示', () => {
    expect(() => readDoc(patched((wd) => { wd[11] = wd[11]! | 0x01; }))).toThrow(/加密/);
  });

  it('Word 95 及更早版本给出明确提示', () => {
    expect(() => readDoc(patched((wd) => { wd[2] = 0x65; wd[3] = 0; }))).toThrow(/Word 95/);
  });
});

describe('SPRM 解析', () => {
  it('按 spra 位决定操作数大小，TDefTable 是 cb + 1', () => {
    const buf = Uint8Array.from([
      0x35, 0x08, 0x01, // sprmCFBold：1 字节
      0x43, 0x4a, 0x20, 0x00, // sprmCHps：2 字节
      0x70, 0x68, 1, 2, 3, 0, // sprmCCv：4 字节
      0x0d, 0xc6, 0x03, 0xaa, 0xbb, 0xcc, // 变长：首字节是长度
      0x08, 0xd6, 0x03, 0x00, 0x01, 0x02, // sprmTDefTable：cb = 3 → 共 4 字节（含 cb）
    ]);
    const s = parseGrpprl(buf);
    expect(s.map((x) => x.op)).toEqual([0x0835, 0x4a43, 0x6870, 0xc60d, 0xd608]);
    expect(Array.from(s[3]!.data)).toEqual([0xaa, 0xbb, 0xcc]);
    expect(Array.from(s[4]!.data)).toEqual([0x01, 0x02]);
  });

  it('被截断的数据不越界', () => {
    expect(parseGrpprl(Uint8Array.from([0x43, 0x4a, 0x20]))).toEqual([]);
  });
});

describe('段落属性 SPRM（[MS-DOC] 2.6.2）', () => {
  const le = (op: number, ...data: number[]) => [op & 0xff, op >> 8, ...data];
  const pap = (...xs: number[][]) => papOf(Uint8Array.from(xs.flat()));

  it('旧版缩进：0x840E 是右缩进、0x840F 是左缩进（只有旧版 sprm 的 .doc 靠它们）', () => {
    const p = pap(le(0x840f, 0xa4, 0x01), le(0x840e, 0x14, 0x00)); // 左 420 twips、右 20 twips
    expect(p.indentLeft).toBe(21);
    expect(p.indentRight).toBe(1);
  });

  it('按行的段前 / 段后、首行缩进字符数；同时写了磅值时字符数优先', () => {
    const p = pap(le(0xa413, 0x9c, 0x00), le(0x4458, 0x32, 0x00), le(0x4459, 0x32, 0x00), le(0x4457, 0xc8, 0x00), le(0x8460, 0xe0, 0x01));
    expect(p.spaceBefore).toBe(7.8);
    expect(p.spaceBeforeLines).toBe(0.5);
    expect(p.spaceAfterLines).toBe(0.5);
    expect(p.firstLineChars).toBe(2);
    expect(p.firstLine).toBe(24);
    expect(pap(le(0x4458, 0, 0)).spaceBeforeLines).toBeNull();
  });

  it('段前 / 段后自动（0x245B / 0x245C），不影响对齐（0x2403 与 0x2461 共用分支）', () => {
    const p = pap(le(0x2403, 1), le(0x245b, 1), le(0x2461, 1));
    expect(p.spaceBeforeAuto).toBe(true);
    expect(p.spaceAfterAuto).toBeUndefined();
    expect(p.align).toBe('center');
    expect(pap(le(0x2403, 1)).spaceBeforeAuto).toBeUndefined();
  });

  it('图文框：sprmPPc 定锚点，sprmPDxaAbs 的 -4 是居中，sprmPDyaAbs 是磅值', () => {
    // pcVert = 2（段落），pcHorz = 1（页边距）
    const p = pap(le(0x261b, (2 << 4) | (1 << 6)), le(0x8418, 0xfc, 0xff), le(0x8419, 0x14, 0x00), le(0x2423, 2));
    expect(p.frame).toEqual({ hAnchor: 'margin', vAnchor: 'text', xAlign: 'center', x: 0, y: 1 });
    expect(pap(le(0x8418, 0x90, 0x01)).frame).toEqual({ hAnchor: 'text', vAnchor: 'margin', x: 20, y: 0 });
  });
});

describe('旧版兼容模式合同（Word 生成：scripts/word-make-compat.ps1，文字取自真实合同的几段条款）', () => {
  const name = '旧版兼容模式合同';
  const doc = readDoc(bytes(join(SAMPLES, 'doc', `${name}.doc`)));
  const ref = readDocx(bytes(join(SAMPLES, `${name}.docx`)));
  const paras = (d: Doc) => d.body.flatMap((b) => (b.kind === 'p' ? [b] : b.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs))));

  it('兼容模式 11、另一套东亚断行规则、压缩标点', () => {
    for (const d of [doc, ref]) expect([d.setup.compatMode, d.setup.altKinsoku, d.setup.compressPunct]).toEqual([11, true, true]);
  });

  it('段距（磅、行、自动）、首行缩进字符数、字符间距和 .docx 一致', () => {
    const key = (p: Extract<Block, { kind: 'p' }>) => {
      const x = p.props;
      return JSON.stringify([x.spaceBefore, x.spaceAfter, x.spaceBeforeLines, x.spaceAfterLines, !!x.spaceBeforeAuto, x.firstLineChars, x.line, p.runs.map((r) => r.props.spacing ?? 0).find((v) => v) ?? 0]);
    };
    expect(paras(doc).map(key)).toEqual(paras(ref).map(key));
    expect(paras(doc).some((p) => p.props.spaceBeforeAuto)).toBe(true);
    expect(paras(doc).some((p) => p.runs.some((r) => r.props.spacing === -0.5))).toBe(true);
  });

  it('页脚的图文框空段落读成图文框（不占页脚高度）', () => {
    const fr = (d: Doc) => d.footer.map((p) => p.props.frame && { h: p.props.frame.hAnchor, v: p.props.frame.vAnchor, x: p.props.frame.xAlign });
    expect(fr(doc)).toEqual(fr(ref));
    expect(fr(doc)[0]).toEqual({ h: 'margin', v: 'text', x: 'center' });
  });
});

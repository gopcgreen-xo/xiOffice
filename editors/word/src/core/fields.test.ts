import { describe, expect, it } from 'vitest';
import { FIELD_CHAR, newParagraph, type Doc, type Paragraph, type Run, type RunProps } from './model';
import { evalField, evalFormula, fieldIdsInRange, formatInstr, parseInstr, setFieldInstr, updateFields, withSwitch } from './fields';
import { fieldAt, fieldSpans, instrPieces, pairedMarkers } from './fieldView';

const props = (): RunProps => ({ ...newParagraph('').mark });
const text = (t: string): Run => ({ text: t, props: props() });
const begin = (id: string, instr: string, locked = false): Run => ({ text: FIELD_CHAR, props: props(), fc: { kind: 'begin', id, instr, ...(locked ? { locked } : {}) } });
const end = (id: string): Run => ({ text: FIELD_CHAR, props: props(), fc: { kind: 'end', id } });
const para = (...runs: Run[]): Paragraph => ({ ...newParagraph(''), runs });
const docOf = (...ps: Paragraph[]): Doc => {
  const d = { body: ps, header: [], footer: [], headerFirst: [], footerFirst: [], headerEven: [], footerEven: [], lists: [], lossy: [] } as unknown as Doc;
  d.setup = { ...({} as Doc['setup']) };
  return d;
};
const result = (p: Paragraph, id: string) => {
  const f = fieldSpans(p).find((x) => x.id === id)!;
  return p.runs.slice(f.begin + 1, f.end).map((r) => r.text).join('');
};

describe('域代码解析', () => {
  it('域名、参数、开关（带值的 \\@ \\# \\* \\r，不带值的 \\p \\h），引号和转义', () => {
    expect(parseInstr(' DATE \\@ "yyyy年M月d日" \\* MERGEFORMAT ')).toEqual({ name: 'DATE', args: [], sw: [{ key: '@', value: 'yyyy年M月d日' }, { key: '*', value: 'MERGEFORMAT' }] });
    expect(parseInstr('SEQ 表 \\r 3 \\h')).toEqual({ name: 'SEQ', args: ['表'], sw: [{ key: 'r', value: '3' }, { key: 'h' }] });
    expect(parseInstr('FILENAME \\p')).toEqual({ name: 'FILENAME', args: [], sw: [{ key: 'p' }] });
    expect(parseInstr('QUOTE "说 \\"你好\\""')).toEqual({ name: 'QUOTE', args: ['说 "你好"'], sw: [] });
    expect(parseInstr('= 2 * (3 + 4) \\# "0.00"')).toEqual({ name: '=', args: ['2 * (3 + 4)'], sw: [{ key: '#', value: '0.00' }] });
    expect(parseInstr('IF {MERGEFIELD a} = 1 "x" "y"').args[0]).toBe('{MERGEFIELD a}');
  });

  it('拼回域代码；改开关时 \\* MERGEFORMAT 保留，大小写类的 \\* 只留一个', () => {
    const i = parseInstr('DATE \\@ "yyyy/M/d" \\* MERGEFORMAT');
    expect(formatInstr(withSwitch(i, '@', 'H:mm'))).toBe('DATE \\* MERGEFORMAT \\@ "H:mm"');
    expect(formatInstr(withSwitch(withSwitch(i, '*', 'Upper'), '*', 'Lower'))).toBe('DATE \\@ "yyyy/M/d" \\* MERGEFORMAT \\* Lower');
    expect(formatInstr(withSwitch(i, '@', null))).toBe('DATE \\* MERGEFORMAT');
    expect(formatInstr(parseInstr('QUOTE "a b"'))).toBe('QUOTE "a b"');
  });

  it('嵌套的域按花括号拆开；花括号不配对时整串当文字', () => {
    expect(instrPieces(' = 1 + {QUOTE 2} ')).toEqual([{ t: 'text', v: ' = 1 + ' }, { t: 'begin' }, { t: 'text', v: 'QUOTE 2' }, { t: 'end' }, { t: 'text', v: ' ' }]);
    expect(instrPieces('QUOTE "}"')).toEqual([{ t: 'text', v: 'QUOTE "}"' }]);
  });
});

describe('计算域结果', () => {
  const now = new Date(2026, 9, 5, 9, 7, 3);

  it('日期 / 时间：默认格式与 \\@ 图片', () => {
    expect(evalField('DATE', { now })).toBe('2026/10/5');
    expect(evalField('DATE \\@ "yyyy年M月d日 dddd"', { now })).toBe('2026年10月5日 星期一');
    expect(evalField('TIME', { now })).toBe('9:07');
    expect(evalField('TIME \\@ "HH:mm:ss"', { now })).toBe('09:07:03');
  });

  it('文件名（\\p 带路径）、总页数、文字、大小写', () => {
    const env = { fileName: 'a.docx', filePath: 'D:\\x\\a.docx', numPages: 7 };
    expect(evalField('FILENAME', env)).toBe('a.docx');
    expect(evalField('FILENAME \\p \\* Upper', env)).toBe('D:\\X\\A.DOCX');
    expect(evalField('NUMPAGES', env)).toBe('7');
    expect(evalField('NUMPAGES', {})).toBeNull();
    expect(evalField('QUOTE "hello world" \\* Caps', {})).toBe('Hello World');
  });

  it('SEQ：按顺序加一，\\r 重设、\\c 重复上一个、\\h 不显示；不同标识分开数；\\* 数字样式', () => {
    const seq = new Map<string, number>();
    expect(evalField('SEQ 表', {}, seq)).toBe('1');
    expect(evalField('SEQ 图 \\* ROMAN', {}, seq)).toBe('I');
    expect(evalField('SEQ 表', {}, seq)).toBe('2');
    expect(evalField('SEQ 表 \\c', {}, seq)).toBe('2');
    expect(evalField('SEQ 表 \\h', {}, seq)).toBe('');
    expect(evalField('SEQ 表 \\r 10', {}, seq)).toBe('10');
    expect(evalField('SEQ 表 \\* CHINESENUM3', {}, seq)).toBe('十一');
    expect(evalField('SEQ 图 \\* alphabetic', {}, seq)).toBe('b');
  });

  it('公式：四则运算、乘方、百分号、函数、书签名；\\# 数字图片；不认识的返回 null', () => {
    expect(evalFormula('2 * (3 + 4) - 2 ^ 3')).toBe(6);
    expect(evalFormula('50%*4')).toBe(2);
    expect(evalFormula('SUM(1,2,3)+ROUND(2.345,2)')).toBeCloseTo(8.35);
    expect(evalFormula('单价*2', (n) => (n === '单价' ? '1,250.5' : null))).toBe(2501);
    expect(evalFormula('1 +')).toBeNull();
    expect(evalField('= 10 / 4', {})).toBe('2.5');
    expect(evalField('= 1234.5 \\# "#,##0.00"', {})).toBe('1,234.50');
    expect(evalField('= 1 / 0', {})).toBeNull();
  });

  it('嵌套的域先算；书签引用（REF 和只写书签名）；算不出来的返回 null', () => {
    const env = { bookmark: (n: string) => (n === '甲方' ? '某某公司' : null) };
    expect(evalField('= 1 + {QUOTE 2}', env)).toBe('3');
    expect(evalField('REF 甲方 \\h', env)).toBe('某某公司');
    expect(evalField('甲方', env)).toBe('某某公司');
    expect(evalField('REF 乙方', env)).toBeNull();
    expect(evalField('AUTHOR', env)).toBeNull();
    expect(evalField('TOC \\o "1-3"', env)).toBeNull();
  });
});

describe('段落里的域', () => {
  it('按出现顺序配对（同 Word），配不上的标记不算；光标所在的域取最里层', () => {
    const p = para(text('a'), begin('x', 'QUOTE 1'), text('1'), end('x'), end('y'), begin('z', 'QUOTE 2'), text('2'));
    const spans = fieldSpans(p);
    expect(spans.map((f) => [f.id, f.from, f.to])).toEqual([['x', 1, 4]]);
    expect(pairedMarkers(p.runs).size).toBe(2);
    expect(fieldAt(p, 1)).toBeNull();
    expect(fieldAt(p, 2)?.id).toBe('x');
    expect(fieldAt(p, 4)?.id).toBe('x');
    expect(fieldAt(p, 5)).toBeNull();
  });

  it('更新域：只换结果、用原结果的格式；SEQ 从头数（没选中的也算）；锁定的不改；算不出来的保留', () => {
    const bold = { ...props(), bold: true };
    const p1 = para(text('表'), begin('s1', 'SEQ 表'), { text: '9', props: bold }, end('s1'));
    const p2 = para(text('表'), begin('s2', 'SEQ 表'), text('9'), end('s2'), begin('l', 'QUOTE 新', true), text('旧'), end('l'));
    const p3 = para(begin('a', 'AUTHOR'), text('张三'), end('a'), begin('q', 'QUOTE ""'), text('x'), end('q'));
    const doc = docOf(p1, p2, p3);

    const only = updateFields(doc, {}, new Set(['s2']));
    expect(only).toEqual({ updated: 1, kept: [], locked: 0 });
    expect(result(p1, 's1')).toBe('9');
    expect(result(p2, 's2')).toBe('2');

    const all = updateFields(doc, {});
    expect(all).toEqual({ updated: 2, kept: ['AUTHOR'], locked: 1 });
    expect(result(p1, 's1')).toBe('1');
    expect(p1.runs[2].props.bold).toBe(true);
    expect(result(p2, 'l')).toBe('旧');
    expect(result(p3, 'a')).toBe('张三');
    // 结果为空时不留空文本段，标记还在
    expect(result(p3, 'q')).toBe('');
    expect(fieldSpans(p3)).toHaveLength(2);
    // 结果没变的不算更新
    expect(updateFields(doc, {}).updated).toBe(0);
  });

  it('改域代码：去掉原始 XML，保留锁定；选区碰到的域', () => {
    const b = begin('x', 'QUOTE 1', true);
    b.fc!.raw = '<w:r/>';
    const p = para(text('ab'), b, text('1'), end('x'), text('cd'));
    expect(setFieldInstr(p, 'x', '  QUOTE 2 ')).toBe(true);
    expect(p.runs[1].fc).toEqual({ kind: 'begin', id: 'x', instr: 'QUOTE 2', locked: true });
    expect(setFieldInstr(p, 'nope', 'X')).toBe(false);
    const doc = docOf(p);
    const at = (offset: number) => ({ story: 'body' as const, path: [0], offset });
    expect(fieldIdsInRange(doc, at(3), at(3))).toEqual(['x']);
    expect(fieldIdsInRange(doc, at(0), at(2))).toEqual([]);
    expect(fieldIdsInRange(doc, at(0), at(3))).toEqual(['x']);
  });
});

import { describe, expect, it } from 'vitest';
import { checkParaForm, formatLength, lengthToPt, lineAtFor, nudgeFontSize, paraFormOf, paraPatchOf, parseFontSize, parseLength, stepFontSize } from './format';
import { catalogFromDoc, clearFormatting } from './styles';
import { applyRunProps, changeIndent, getPara, insertRaw, splitParagraph } from './ops';
import { defaultParaProps, newDocument, type ParaProps } from './model';

describe('带单位的长度', () => {
  it('没写单位按缺省单位；中英文单位名都认', () => {
    expect(parseLength('2', 'char')).toEqual({ value: 2, unit: 'char' });
    expect(parseLength('0.5行', 'char')).toEqual({ value: 0.5, unit: 'line' });
    expect(parseLength(' 12 pt ', 'char')).toEqual({ value: 12, unit: 'pt' });
    expect(parseLength('1.5 厘米', 'pt')).toEqual({ value: 1.5, unit: 'cm' });
    expect(parseLength('abc', 'pt')).toBeNull();
    expect(parseLength('3 尺', 'pt')).toBeNull();
  });

  it('换算成磅：字符按字宽，厘米 / 毫米 / 英寸按 72 磅每英寸', () => {
    expect(lengthToPt({ value: 2, unit: 'char' }, 10.5)).toBe(21);
    expect(lengthToPt({ value: 2.54, unit: 'cm' }, 10.5)).toBeCloseTo(72);
    expect(lengthToPt({ value: 25.4, unit: 'mm' }, 10.5)).toBeCloseTo(72);
    expect(formatLength(1.999999, 'char')).toBe('2 字符');
  });
});

describe('字号', () => {
  it('中文字号名、数字、带“磅”都能输入；按 0.5 磅取整，超出 1–1638 无效', () => {
    expect(parseFontSize('小四')).toBe(12);
    expect(parseFontSize('13.5')).toBe(13.5);
    expect(parseFontSize('13.3 磅')).toBe(13.5);
    expect(parseFontSize('0')).toBeNull();
    expect(parseFontSize('2000')).toBeNull();
    expect(parseFontSize('大号')).toBeNull();
  });

  it('增大 / 减小字号按字号表走，72 磅以上每次 10 磅；Ctrl+] / [ 每次 1 磅', () => {
    expect(stepFontSize(10.5, 1)).toBe(11);
    expect(stepFontSize(12, 1)).toBe(14);
    expect(stepFontSize(13, 1)).toBe(14);
    expect(stepFontSize(13, -1)).toBe(12);
    expect(stepFontSize(72, 1)).toBe(80);
    expect(stepFontSize(80, -1)).toBe(72);
    expect(stepFontSize(5, -1)).toBe(4);
    expect(nudgeFontSize(10.5, 1)).toBe(12);
    expect(nudgeFontSize(1, -1)).toBe(1);
  });
});

describe('段落对话框的表单', () => {
  const props = (patch: Partial<ParaProps> = {}): ParaProps => ({ ...defaultParaProps(), ...patch });

  it('格式 → 表单：缩进按字符显示，段前段后按行或磅，行距分单倍 / 多倍 / 固定值', () => {
    const f = paraFormOf(props({ indentLeft: 21, firstLineChars: 2, spaceBeforeLines: 0.5, spaceAfter: 6, line: { rule: 'exact', value: 28 } }), 10.5);
    expect(f.left).toBe('2 字符');
    expect(f.special).toBe('first');
    expect(f.specialBy).toBe('2 字符');
    expect(f.before).toBe('0.5 行');
    expect(f.after).toBe('6 磅');
    expect(f.lineRule).toBe('exact');
    expect(f.lineAt).toBe('28 磅');
    expect(paraFormOf(props({ line: { rule: 'auto', value: 1.25 } }), 10.5).lineRule).toBe('multiple');
    expect(paraFormOf(props({ firstLine: -21 }), 10.5)).toMatchObject({ special: 'hanging', specialBy: '2 字符' });
  });

  it('只把改过的项写回：没动的缩进、间距不受换算误差影响', () => {
    const init = paraFormOf(props({ indentLeft: 17.3 }), 10.5);
    expect(paraPatchOf({ ...init }, init, 10.5)).toEqual({});
    const patch = paraPatchOf({ ...init, keepNext: true, after: '1 行' }, init, 10.5);
    expect(patch).toEqual({ keepNext: true, spaceAfterLines: 1, spaceAfter: 0, spaceAfterAuto: false });
  });

  it('悬挂缩进按字符记字符数（负数）；按厘米记磅值；行距固定值换成磅', () => {
    const init = paraFormOf(props(), 10.5);
    expect(paraPatchOf({ ...init, special: 'hanging', specialBy: '2' }, init, 10.5)).toEqual({ firstLine: -21, firstLineChars: -2 });
    expect(paraPatchOf({ ...init, special: 'first', specialBy: '1 厘米' }, init, 10.5)).toEqual({ firstLine: 28.35, firstLineChars: null });
    expect(paraPatchOf({ ...init, lineRule: 'exact', lineAt: '20 磅' }, init, 10.5).line).toEqual({ rule: 'exact', value: 20 });
    expect(paraPatchOf({ ...init, lineRule: 'multiple', lineAt: '1.25' }, init, 10.5).line).toEqual({ rule: 'auto', value: 1.25 });
    expect(paraPatchOf({ ...init, before: '自动' }, init, 10.5)).toEqual({ spaceBeforeAuto: true, spaceBeforeLines: null });
  });

  it('读不出来的值给出说明；换行距类型时设置值换成常用值', () => {
    const init = paraFormOf(props(), 10.5);
    expect(checkParaForm(init)).toBeNull();
    expect(checkParaForm({ ...init, left: '两字' })).toContain('左侧缩进');
    expect(checkParaForm({ ...init, before: '2 字符' })).toContain('段前');
    expect(checkParaForm({ ...init, lineRule: 'exact', lineAt: '0' })).toContain('行距');
    expect(checkParaForm({ ...init, lineRule: 'multiple', lineAt: '0' })).toContain('多倍行距');
    expect(lineAtFor('exact', init)).toBe('12 磅');
    expect(lineAtFor('multiple', init)).toBe('3');
  });
});

describe('增加 / 减少缩进量、清除格式', () => {
  it('左缩进移到下一个 / 上一个默认制表位，不小于 0', () => {
    const doc = newDocument();
    const at = { story: 'body' as const, path: [0], offset: 0 };
    const tab = doc.setup.defaultTab;
    expect(changeIndent(doc, at, at, 1)).toBe(true);
    expect(getPara(doc, at).props.indentLeft).toBe(tab);
    getPara(doc, at).props.indentLeft = tab * 1.5;
    changeIndent(doc, at, at, 1);
    expect(getPara(doc, at).props.indentLeft).toBe(tab * 2);
    getPara(doc, at).props.indentLeft = tab * 1.5;
    changeIndent(doc, at, at, -1);
    expect(getPara(doc, at).props.indentLeft).toBe(tab);
    changeIndent(doc, at, at, -1);
    expect(changeIndent(doc, at, at, -1)).toBe(false);
    expect(getPara(doc, at).props.indentLeft).toBe(0);
  });

  it('清除格式：文字回到段落样式的格式；整段清除时段落格式回到正文、去掉编号', () => {
    const doc = newDocument();
    const p0 = { story: 'body' as const, path: [0], offset: 0 };
    insertRaw(doc, p0, '甲乙丙');
    const end = { ...p0, offset: 3 };
    applyRunProps(doc, p0, end, { bold: true, size: 22, color: 'FF0000' });
    getPara(doc, p0).props = { ...getPara(doc, p0).props, align: 'center', indentLeft: 42, num: { id: 1, level: 0 } };
    const cat = catalogFromDoc(doc);
    clearFormatting(doc, cat, p0, { ...p0, offset: 2 }, false);
    const runs = getPara(doc, p0).runs;
    expect(runs[0].props).toMatchObject({ bold: false, size: doc.defaults.size, color: doc.defaults.color });
    expect(runs[runs.length - 1].props.bold).toBe(true);
    expect(getPara(doc, p0).props.align).toBe('center');
    clearFormatting(doc, cat, p0, end, true);
    expect(getPara(doc, p0).props).toMatchObject({ align: defaultParaProps().align, indentLeft: 0, num: null });
  });

  it('清除格式保留分节符', () => {
    const doc = newDocument();
    const p0 = { story: 'body' as const, path: [0], offset: 0 };
    insertRaw(doc, p0, '甲乙');
    splitParagraph(doc, { ...p0, offset: 1 });
    getPara(doc, p0).props.sectionBreak = 'nextPage';
    clearFormatting(doc, catalogFromDoc(doc), p0, { story: 'body', path: [1], offset: 1 }, true);
    expect(getPara(doc, p0).props.sectionBreak).toBe('nextPage');
  });
});

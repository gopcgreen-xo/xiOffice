/**
 * .doc 绘图对象（OfficeArt）的最小测试：手写记录，不依赖样本文件。
 * 和 Word 生成的样本（文本框、各种形状、浮动图片）的对照在 read.test.ts。
 */
import { describe, expect, it } from 'vitest';
import { newParagraph } from '../core/model';
import { parseOfficeArt, shapeBoxes, type Fspa, type ShapeEnv } from './art';

const u16 = (v: number) => [v & 0xff, (v >> 8) & 0xff];
const u32 = (v: number) => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
/** 一条 OfficeArt 记录：ver（低 4 位）+ inst、类型、长度、内容 */
const rec = (ver: number, inst: number, type: number, body: number[]) => [...u16(ver | (inst << 4)), ...u16(type), ...u32(body.length), ...body];
const container = (type: number, kids: number[][]) => rec(0xf, 0, type, kids.flat());
const fsp = (spt: number, spid: number, flags: number) => rec(2, spt, 0xf00a, [...u32(spid), ...u32(flags)]);
const opt = (type: number, props: [number, number][]) => rec(3, props.length, type, props.flatMap(([pid, v]) => [...u16(pid), ...u32(v)]));
const rect4 = (type: number, r: number[]) => rec(0, 0, type, r.flatMap(u32));
const sp = (...kids: number[][]) => container(0xf004, kids);

/** OfficeArtContent：空的 DggContainer + 正文绘图（dgglbl 0） */
function content(shapes: number[][]): Uint8Array {
  const dgg = container(0xf000, []);
  const patriarch = sp(rect4(0xf009, [0, 0, 0, 0]), fsp(0, 1024, 0x5));
  const dg = container(0xf002, [rec(0, 1, 0xf008, [...u32(0), ...u32(0)]), container(0xf003, [patriarch, ...shapes])]);
  return new Uint8Array([...dgg, 0, ...dg]);
}

const env = (lossy: string[] = []): ShapeEnv => ({ image: () => null, text: () => null, emptyPara: () => newParagraph(), lossy: (w) => lossy.push(w) });
const fspa = (spid: number, l: number, t: number, r: number, b: number, over: Partial<Fspa> = {}): Fspa => ({ spid, l, t, r, b, bx: 2, by: 2, wr: 3, below: false, ...over });

describe('.doc 绘图对象（OfficeArt）', () => {
  it('组合拆成成员：子形状的位置按组合坐标系（FSPGR）换算到锚点外框里', () => {
    const group = container(0xf003, [
      sp(rect4(0xf009, [0, 0, 1000, 1000]), fsp(0, 2000, 0x201)),
      sp(fsp(3, 2001, 0x202), opt(0xf00b, [[0x181, 0x0000ff]]), rect4(0xf00f, [0, 0, 500, 500])),
      // 第二个成员水平翻转
      sp(fsp(20, 2002, 0x242), rect4(0xf00f, [500, 500, 1000, 1000])),
    ]);
    const art = parseOfficeArt(content([group]));
    const g = art.main.get(2000)!;
    expect(g.group?.kids.map((k) => k.spid)).toEqual([2001, 2002]);
    // 锚点外框 2000 × 1000 twips = 100 × 50 磅
    const boxes = shapeBoxes(g, fspa(2000, 200, 400, 2200, 1400), env());
    expect(boxes.map((b) => [b.geom, b.posH.offset, b.posV.offset, b.w, b.h])).toEqual([
      ['ellipse', 10, 20, 50, 25],
      ['line', 60, 45, 50, 25],
    ]);
    expect(boxes[0]!.fill).toBe('FF0000');
    expect(boxes[1]!.flipH).toBe(true);
    expect(boxes[1]!.fill).toBeNull(); // 直线没有填充
  });

  it('布尔属性组：主属性表和第三属性表各写一部分位，按“有效位”合并（线条关闭不被第三属性表覆盖）', () => {
    const shape = sp(fsp(1, 3000, 0xa00), opt(0xf00b, [[0x1ff, 0x80000], [0x1bf, 0x100010]]), opt(0xf122, [[0x1ff, 0x400000]]));
    const s = parseOfficeArt(content([shape])).main.get(3000)!;
    const [b] = shapeBoxes(s, fspa(3000, 0, 0, 200, 200), env());
    expect(b!.line).toBeNull();
    expect(b!.fill).toBe('FFFFFF'); // 填充开、没写颜色 = 白色
  });

  it('缺省值：填充白色、线条黑色 0.75 磅、文字边距 7.2 / 3.6 磅、左右环绕距离 9 磅；衬于文字下方、环绕方式来自 FSPA', () => {
    const s = parseOfficeArt(content([sp(fsp(2, 3100, 0xa00))])).main.get(3100)!;
    const [b] = shapeBoxes(s, fspa(3100, 0, 0, 400, 200, { wr: 1, below: true, bx: 1, by: 0 }), env());
    expect(b).toMatchObject({
      geom: 'roundRect',
      fill: 'FFFFFF',
      line: { color: '000000', w: 0.75 },
      inset: { l: 7.2, t: 3.6, r: 7.2, b: 3.6 },
      dist: { l: 9, t: 0, r: 9, b: 0 },
      wrap: 'topAndBottom',
      behind: true,
      posH: { rel: 'page', offset: 0 },
      posV: { rel: 'margin', offset: 0 },
    });
  });

  it('旋转 45°–135° 的形状：FSPA 存的是转过 90° 的外框，宽高互换、中心不变', () => {
    const s = parseOfficeArt(content([sp(fsp(1, 3200, 0xa00), opt(0xf00b, [[0x0004, 90 * 65536]]))])).main.get(3200)!;
    const [b] = shapeBoxes(s, fspa(3200, 0, 0, 2000, 1000), env());
    expect([b!.posH.offset, b!.posV.offset, b!.w, b!.h, b!.rot]).toEqual([25, -25, 50, 100, 90]);
  });

  it('隐藏的形状不显示；自由曲线等不支持的形状记进“不支持的内容”', () => {
    const lossy: string[] = [];
    const art = parseOfficeArt(content([sp(fsp(1, 3300, 0xa00), opt(0xf122, [[0x3bf, 0x20002]])), sp(fsp(0, 3301, 0xa00))]));
    expect(shapeBoxes(art.main.get(3300)!, fspa(3300, 0, 0, 100, 100), env(lossy))).toEqual([]);
    expect(lossy).toEqual([]);
    expect(shapeBoxes(art.main.get(3301)!, fspa(3301, 0, 0, 100, 100), env(lossy))).toEqual([]);
    expect(lossy).toEqual(['自由曲线、艺术字等形状']);
  });

  it('文本框（202）没有文字时给一个空段落；形状有文字时带上文字', () => {
    const art = parseOfficeArt(content([sp(fsp(202, 3400, 0xa00)), sp(fsp(1, 3401, 0xa00))]));
    const text = new Map([[3401, [newParagraph('形状里的字')]]]);
    const e: ShapeEnv = { ...env(), text: (spid) => text.get(spid) ?? null };
    const [tb] = shapeBoxes(art.main.get(3400)!, fspa(3400, 0, 0, 100, 100), e);
    expect(tb!.geom).toBeUndefined();
    expect(tb!.paragraphs).toHaveLength(1);
    const [r] = shapeBoxes(art.main.get(3401)!, fspa(3401, 0, 0, 100, 100), e);
    expect(r!.geom).toBe('rect');
    expect(r!.paragraphs[0]!.runs[0]!.text).toBe('形状里的字');
  });
});

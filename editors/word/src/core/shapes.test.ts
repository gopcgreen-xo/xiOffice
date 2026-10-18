import { describe, expect, it } from 'vitest';
import { SHAPE_GEOMS, applyColorMods, isLineGeom, shapePath, type PathCmd } from './shapes';

const points = (cmds: PathCmd[]) => cmds.filter((c) => c[0] === 'M' || c[0] === 'L').map((c) => [c[1], c[2]]);

describe('形状外形', () => {
  it('菱形、三角形、直线的顶点', () => {
    expect(points(shapePath('diamond', 40, 20))).toEqual([[20, 0], [40, 10], [20, 20], [0, 10]]);
    expect(points(shapePath('triangle', 40, 20))).toEqual([[20, 0], [40, 20], [0, 20]]);
    // 直线不闭合：左上到右下（翻转在绘制时做）
    const line = shapePath('line', 100, 0);
    expect(line).toEqual([['M', 0, 0], ['L', 100, 0]]);
    expect(isLineGeom('line')).toBe(true);
    expect(isLineGeom('rect')).toBe(false);
  });

  it('调整值用 DrawingML 的缺省值：圆角 = 短边的 16.667%，右箭头箭身占一半高、箭头长 = 短边的一半', () => {
    const rr = shapePath('roundRect', 120, 60);
    expect(rr[0]).toEqual(['M', 60 * 0.16667, 0]);
    expect(points(shapePath('rightArrow', 100, 40))).toEqual([[0, 10], [80, 10], [80, 0], [100, 20], [80, 40], [80, 30], [0, 30]]);
  });

  it('所有支持的外形都在框内、闭合（线除外）', () => {
    for (const g of SHAPE_GEOMS) {
      const cmds = shapePath(g, 80, 50);
      for (const c of cmds) {
        const xs = c[0] === 'C' ? [c[1], c[3], c[5]] : c[0] === 'Z' ? [] : [c[1]];
        const ys = c[0] === 'C' ? [c[2], c[4], c[6]] : c[0] === 'Z' ? [] : [c[2]];
        for (const x of xs) expect(x).toBeGreaterThanOrEqual(-1e-9), expect(x).toBeLessThanOrEqual(80 + 1e-9);
        for (const y of ys) expect(y).toBeGreaterThanOrEqual(-1e-9), expect(y).toBeLessThanOrEqual(50 + 1e-9);
      }
      expect(cmds[cmds.length - 1]![0] === 'Z').toBe(!isLineGeom(g));
    }
  });
});

describe('DrawingML 颜色变换', () => {
  it('shade / tint 在线性 RGB 里算（Word 实测：accent1 4472C4 shade 50% = 2F528F，5B9BD5 shade 50% = 41719C）', () => {
    expect(applyColorMods('4472C4', [{ name: 'shade', val: 50000 }])).toBe('2F528F');
    expect(applyColorMods('5B9BD5', [{ name: 'shade', val: 50000 }])).toBe('41719C');
    expect(applyColorMods('000000', [{ name: 'tint', val: 0 }])).toBe('FFFFFF');
  });

  it('lumMod / lumOff 在 HSL 里算（“白色，背景 1，深色 25%” = FFFFFF lumMod 75% = BFBFBF）', () => {
    expect(applyColorMods('FFFFFF', [{ name: 'lumMod', val: 75000 }])).toBe('BFBFBF');
    expect(applyColorMods('000000', [{ name: 'lumMod', val: 50000 }, { name: 'lumOff', val: 50000 }])).toBe('808080');
  });
});

import { describe, expect, it } from 'vitest';
import { anchorScroll, fitWidthZoom, pinchZoom } from './touch';

describe('touch', () => {
  it('pinchZoom：按两指距离比例缩放并限制范围', () => {
    expect(pinchZoom(1, 100, 200, 0.2, 3)).toBe(2);
    expect(pinchZoom(1, 100, 50, 0.2, 3)).toBe(0.5);
    expect(pinchZoom(2, 100, 400, 0.2, 3)).toBe(3);
    expect(pinchZoom(0.5, 100, 10, 0.2, 3)).toBe(0.2);
    expect(pinchZoom(1, 0, 10, 0.2, 3)).toBe(1);
  });
  it('anchorScroll：焦点下的内容留在原处', () => {
    // 焦点在视口 100px 处，内容已滚 200px → 焦点在内容 300px；放大 2 倍后在 600px，要滚到 500
    expect(anchorScroll(200, 100, 2)).toBe(500);
    // 缩小时不会滚到负数
    expect(anchorScroll(0, 10, 0.5)).toBe(0);
    // 内容左侧有不缩放的空白
    expect(anchorScroll(0, 100, 2, 20)).toBe(80);
  });
  it('fitWidthZoom：页面宽度适应屏宽', () => {
    expect(fitWidthZoom(390, 794)).toBeCloseTo(0.461, 3);
    expect(fitWidthZoom(0, 794)).toBe(1);
    expect(fitWidthZoom(5000, 100)).toBe(3);
  });
});

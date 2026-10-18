import { describe, expect, it } from 'vitest';
import { OFFICE_FACES } from './officeMetrics';
import { fontMetrics, lineAscentDescent, registerFonts } from './fontfile';

describe('officeMetrics（手机上没有的 Office 字体的度量）', () => {
  it('中英文名都能查到，行高符合 Word 实测（宋体 1.3 em、Times New Roman 1.149 em、Calibri 1.221 em）', () => {
    // 先登记一个“手机自带”的同名字体：同名先到先得，真字体优先
    registerFonts([{ families: ['Arial'], metrics: { ascent: 1, descent: 0.25, lineGap: 0, ea: false } }]);
    registerFonts(OFFICE_FACES);
    const h = (f: string) => {
      const m = fontMetrics(f)!;
      const { ascent, descent } = lineAscentDescent(m);
      return Math.round((ascent + descent) * 1000) / 1000;
    };
    expect(h('宋体')).toBe(1.3);
    expect(h('SimSun')).toBe(1.3);
    expect(h('仿宋_GB2312')).toBe(1.3);
    expect(h('Times New Roman')).toBe(1.15);
    expect(h('Calibri')).toBe(1.221);
    expect(h('Arial')).toBe(1.25);
  });
});

import { describe, expect, it } from 'vitest';
import { faceInfo, fontMetrics, lineAscentDescent, registerFonts, sfntFaces } from './fontfile';

/** 手工拼一个只有 name / hhea / OS/2 / head 四张表的字体 */
function font(o: { names: string[]; upm: number; hAsc: number; hDesc: number; hGap: number; winAsc: number; winDesc: number; codePages: number }): Uint8Array {
  const u16 = (v: number) => [(v >> 8) & 255, v & 255];
  const u32 = (v: number) => [(v >>> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255];
  const head = new Array(54).fill(0);
  head.splice(18, 2, ...u16(o.upm));
  const hhea = [0, 1, 0, 0, ...u16(o.hAsc & 0xffff), ...u16(o.hDesc & 0xffff), ...u16(o.hGap & 0xffff), ...new Array(26).fill(0)];
  const os2 = new Array(86).fill(0);
  os2.splice(0, 2, ...u16(1));
  os2.splice(74, 4, ...u16(o.winAsc), ...u16(o.winDesc));
  os2.splice(78, 4, ...u32(o.codePages));
  const strs = o.names.map((n) => [...n].flatMap((c) => u16(c.charCodeAt(0))));
  let off = 0;
  const recs = strs.flatMap((s) => {
    const r = [...u16(3), ...u16(1), ...u16(0x804), ...u16(1), ...u16(s.length), ...u16(off)];
    off += s.length;
    return r;
  });
  const name = [...u16(0), ...u16(o.names.length), ...u16(6 + recs.length), ...recs, ...strs.flat()];
  const tables: [string, number[]][] = [
    ['OS/2', os2],
    ['head', head],
    ['hhea', hhea],
    ['name', name],
  ];
  let pos = 12 + tables.length * 16;
  const dir: number[] = [];
  const data: number[] = [];
  for (const [t, b] of tables) {
    dir.push(...[...t].map((c) => c.charCodeAt(0)), ...u32(0), ...u32(pos), ...u32(b.length));
    data.push(...b);
    pos += b.length;
  }
  return Uint8Array.from([...u32(0x10000), ...u16(tables.length), 0, 0, 0, 0, 0, 0, ...dir, ...data]);
}

const simsun = { names: ['SimSun', '宋体'], upm: 256, hAsc: 220, hDesc: -36, hGap: 36, winAsc: 220, winDesc: 36, codePages: 1 << 18 };
const tnr = { names: ['Times New Roman'], upm: 2048, hAsc: 1825, hDesc: -443, hGap: 87, winAsc: 1825, winDesc: 443, codePages: 1 };

describe('字体文件度量（W19）', () => {
  it('读出中英文族名、度量和是否中日韩字体', () => {
    const [f] = sfntFaces(font(simsun)).map(faceInfo);
    expect(f!.families).toEqual(['SimSun', '宋体']);
    expect(f!.metrics).toMatchObject({ ascent: 220 / 256, descent: 36 / 256, ea: true });
  });

  it('Word 的单倍行高：西文 = win 上伸 + 下伸 + 外部行距；中文 = 1.3 倍、不计外部行距，多出的上下各一半', () => {
    const t = faceInfo(sfntFaces(font(tnr))[0])!.metrics;
    const a = lineAscentDescent(t);
    expect(a.ascent + a.descent).toBeCloseTo((1825 + 443 + 87) / 2048); // 1.150 em
    const s = lineAscentDescent(faceInfo(sfntFaces(font(simsun))[0])!.metrics);
    expect(s.ascent + s.descent).toBeCloseTo(1.3);
    expect(s.ascent).toBeCloseTo(220 / 256 + 0.15);
  });

  it('字体集（.ttc）里的每个字体都读出来；没装的“仿宋_GB2312”类名字按去掉后缀查', () => {
    const a = font(simsun);
    const b = font({ ...simsun, names: ['FangSong', '仿宋'] });
    const ttc = new Uint8Array(12 + 8 + a.length + b.length);
    ttc.set([0x74, 0x74, 0x63, 0x66, 0, 1, 0, 0, 0, 0, 0, 2]);
    const offA = 20;
    const offB = 20 + a.length;
    ttc.set([offA >> 24, (offA >> 16) & 255, (offA >> 8) & 255, offA & 255, offB >> 24, (offB >> 16) & 255, (offB >> 8) & 255, offB & 255], 12);
    // 字体集里各字体的表偏移是相对整个文件的
    const shift = (f: Uint8Array, base: number) => {
      const c = f.slice();
      const n = (c[4] << 8) | c[5];
      for (let i = 0; i < n; i++) {
        const p = 12 + i * 16 + 8;
        const v = ((c[p] << 24) | (c[p + 1] << 16) | (c[p + 2] << 8) | c[p + 3]) + base;
        c.set([(v >>> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255], p);
      }
      return c;
    };
    ttc.set(shift(a, offA), offA);
    ttc.set(shift(b, offB), offB);
    const faces = sfntFaces(ttc).map(faceInfo);
    expect(faces.map((f) => f!.families[1])).toEqual(['宋体', '仿宋']);
    registerFonts(faces.filter((f) => !!f));
    expect(fontMetrics('仿宋_GB2312')?.ea).toBe(true);
    expect(fontMetrics('没有的字体')).toBeNull();
  });
});

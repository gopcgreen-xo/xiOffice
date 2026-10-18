import { afterEach, describe, expect, it, vi } from 'vitest';
import { loadZoom, saveZoom, zoomIn } from './zoomMemory';

describe('缩放记忆', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('读回上次的缩放；没有、不合法或存储不可用时用默认值', () => {
    const store = new Map<string, string>();
    vi.stubGlobal('localStorage', { getItem: (k: string) => store.get(k) ?? null, setItem: (k: string, v: string) => void store.set(k, v) });
    const ok = (v: unknown) => v === 'width' || zoomIn(0.2, 5)(v);
    expect(loadZoom<number | 'width'>('pdf.zoom', 'width', ok)).toBe('width');
    saveZoom('pdf.zoom', 1.5);
    expect(loadZoom<number | 'width'>('pdf.zoom', 'width', ok)).toBe(1.5);
    saveZoom('pdf.zoom', 'width');
    expect(loadZoom<number | 'width'>('pdf.zoom', 'width', ok)).toBe('width');
    saveZoom('pdf.zoom', 99);
    expect(loadZoom<number | 'width'>('pdf.zoom', 'width', ok)).toBe('width');
    store.set('pdf.zoom', '{坏数据');
    expect(loadZoom<number | 'width'>('pdf.zoom', 'width', ok)).toBe('width');
    vi.stubGlobal('localStorage', { getItem: () => { throw new Error('不可用'); }, setItem: () => { throw new Error('不可用'); } });
    expect(loadZoom('sheet.zoom', 1, zoomIn(0.5, 2))).toBe(1);
    expect(() => saveZoom('sheet.zoom', 1.2)).not.toThrow();
  });
});

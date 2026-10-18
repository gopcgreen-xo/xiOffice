import { afterEach, describe, expect, it, vi } from 'vitest';
import { clearDraft, disableDrafts, loadDraft, saveDraft } from './recovery';

describe('草稿', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('关掉草稿后不读、不写、不删（不碰共用的 IndexedDB）', async () => {
    const open = vi.fn(() => {
      throw new Error('不该打开数据库');
    });
    vi.stubGlobal('indexedDB', { open });
    disableDrafts();
    await saveDraft('D:/a.docx', new Uint8Array([1]));
    expect(await loadDraft('D:/a.docx')).toBeNull();
    await clearDraft('D:/a.docx');
    expect(open).not.toHaveBeenCalled();
  });
});

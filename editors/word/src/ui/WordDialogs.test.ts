// @vitest-environment jsdom
/** 对话框能渲染、能提交（不测样式）：链接、书签、水印、样式、模板与批量生成 */
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { act, createElement as h, type ReactElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { setEditorHost, type EditorHost } from '@wz/editor-shared';
import { BookmarkDialog, LinkDialog, MergeDialog, ProtectDialog, StylesDialog, WatermarkDialog } from './WordDialogs';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const toasts: string[] = [];
const fakeHost = (extra: Partial<EditorHost> = {}) =>
  setEditorHost({ toast: (m: string) => void toasts.push(m), guard: async (p: Promise<unknown>) => p, listDir: async () => [], writeBytes: async () => {}, ...extra } as unknown as EditorHost);
beforeAll(() => {
  fakeHost();
  // jsdom 的 File 没有 arrayBuffer（WebView 里有）
  if (!File.prototype.arrayBuffer)
    File.prototype.arrayBuffer = function (this: File) {
      return new Promise<ArrayBuffer>((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result as ArrayBuffer);
        r.onerror = () => reject(r.error);
        r.readAsArrayBuffer(this);
      });
    };
});

let root: Root | null = null;
let box: HTMLDivElement | null = null;
const mount = (el: ReactElement) => {
  box = document.createElement('div');
  document.body.appendChild(box);
  root = createRoot(box);
  act(() => root!.render(el));
  return box;
};
afterEach(() => {
  act(() => root?.unmount());
  box?.remove();
  root = null;
  toasts.length = 0;
});
const setValue = (input: HTMLInputElement | HTMLSelectElement, v: string) => {
  const proto = input instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
  Object.getOwnPropertyDescriptor(proto, 'value')!.set!.call(input, v);
  act(() => void input.dispatchEvent(new Event(input instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true })));
};
const click = (el: Element | null | undefined) => act(() => void (el as HTMLElement).click());
const button = (c: HTMLElement, text: string) => [...c.querySelectorAll('button')].find((b) => b.textContent?.trim() === text);
const noop = () => {};

describe('对话框', () => {
  it('超链接：填地址确定，补全 https；地址为空提示；已有链接时有“取消链接”', () => {
    const onOk = vi.fn();
    const onRemove = vi.fn();
    const c = mount(h(LinkDialog, { text: '合同', bookmarks: ['第一章'], onOk, onClose: noop }));
    click(button(c, '确定'));
    expect(onOk).not.toHaveBeenCalled();
    expect(toasts.at(-1)).toContain('地址');
    setValue(c.querySelector('input[placeholder^="https"]') as HTMLInputElement, 'example.com');
    click(button(c, '确定'));
    expect(onOk).toHaveBeenCalledWith({ text: '合同', link: { url: 'https://example.com' } });
    act(() => root!.render(h(LinkDialog, { text: '合同', link: { url: 'https://a.cn' }, bookmarks: [], onOk, onRemove, onClose: noop })));
    click(button(c, '取消链接'));
    expect(onRemove).toHaveBeenCalled();
  });

  it('超链接：链到书签；原文件里的链接（有 rid）不改地址时保留 rid', () => {
    const onOk = vi.fn();
    const c = mount(h(LinkDialog, { text: '见', link: { rid: 'rId9', url: 'https://old.cn' }, bookmarks: ['第一章', '付款'], onOk, onClose: noop }));
    click(button(c, '确定'));
    expect(onOk).toHaveBeenLastCalledWith({ text: '见', link: { rid: 'rId9', url: 'https://old.cn' } });
    setValue(c.querySelectorAll('select')[0] as HTMLSelectElement, 'anchor');
    setValue(c.querySelectorAll('select')[1] as HTMLSelectElement, '付款');
    click(button(c, '确定'));
    expect(onOk).toHaveBeenLastCalledWith({ text: '见', link: { anchor: '付款' } });
  });

  it('书签：名字不合法时不能添加；合法时回调；列表里能跳转 / 删除', () => {
    const onAdd = vi.fn();
    const onGo = vi.fn();
    const onDelete = vi.fn();
    const pos = { story: 'body' as const, path: [0], offset: 0 };
    const c = mount(h(BookmarkDialog, { bookmarks: [{ name: '付款', start: pos, end: pos, text: '付款条款' }], canAdd: true, suggest: '', onAdd, onGo, onDelete, onClose: noop }));
    const input = c.querySelector('input') as HTMLInputElement;
    setValue(input, '1abc');
    expect((button(c, '添加') as HTMLButtonElement).disabled).toBe(true);
    expect(c.textContent).toContain('以字母或汉字开头');
    setValue(input, '交货期');
    click(button(c, '添加'));
    expect(onAdd).toHaveBeenCalledWith('交货期');
    setValue(input, '付款');
    expect(button(c, '移到这里')).toBeTruthy();
    click([...c.querySelectorAll('button')].find((b) => b.textContent?.startsWith('付款')));
    expect(onGo).toHaveBeenCalledWith('付款');
    click(button(c, '删除'));
    expect(onDelete).toHaveBeenCalledWith('付款');
  });

  it('水印：确定带出各项；已有水印时可删除', () => {
    const onOk = vi.fn();
    const c = mount(h(WatermarkDialog, { value: { text: '草稿', colorHex: 'FF0000', opacity: 0.3, diagonal: false }, onOk, onClose: noop }));
    click(button(c, '确定'));
    expect(onOk).toHaveBeenLastCalledWith({ text: '草稿', colorHex: 'FF0000', opacity: 0.3, diagonal: false, fontFamily: '等线' });
    click(button(c, '删除水印'));
    expect(onOk).toHaveBeenLastCalledWith(null);
  });

  it('样式：默认只列常用样式，选中后套用；新建要名字；字符样式没选文字时不能新建', () => {
    const onApply = vi.fn();
    const onCreate = vi.fn(() => null);
    const styles = [
      { id: 'a', name: 'Normal', type: 'paragraph' as const, basedOn: null, quick: true, custom: false },
      { id: '1', name: 'heading 1', type: 'paragraph' as const, basedOn: 'a', quick: true, custom: false },
      { id: 'x', name: 'Obscure', type: 'paragraph' as const, basedOn: 'a', quick: false, custom: false },
      { id: 'Red', name: 'Strong Red', type: 'character' as const, basedOn: null, quick: true, custom: false },
    ];
    const c = mount(h(StylesDialog, { styles, currentPara: 'a', hasSelection: false, onApply, onCreate, onRedefine: noop, onRename: () => null, onDelete: () => null, onClose: noop }));
    expect(c.textContent).toContain('正文');
    expect(c.textContent).toContain('标题 1');
    expect(c.textContent).not.toContain('Obscure');
    click([...c.querySelectorAll('div')].find((d) => d.textContent === '标题 1'));
    click(button(c, '套用'));
    expect(onApply).toHaveBeenCalledWith('1', 'paragraph');
    setValue(c.querySelector('input[placeholder^="样式名"]') as HTMLInputElement, '公文正文');
    click(button(c, '新建'));
    expect(onCreate).toHaveBeenCalledWith('公文正文', 'paragraph');
    click(button(c, '字符样式'));
    setValue(c.querySelector('input[placeholder^="样式名"]') as HTMLInputElement, '强调');
    expect((button(c, '新建') as HTMLButtonElement).disabled).toBe(true);
  });

  it('模板与批量生成：读 csv，逐条生成，文件名按字段，不覆盖已有文件', async () => {
    const written: string[] = [];
    fakeHost({ listDir: async () => [{ name: '合同_甲.docx', path: '', size: 0, modified: null }], writeBytes: async (p: string) => void written.push(p) });
    const generate = vi.fn(async (recs: Record<string, string>[], start: number) => ({ bytes: new TextEncoder().encode(recs[start]['名称']), consumed: 1 }));
    const generateAll = vi.fn(async (recs: Record<string, string>[]) => ({ bytes: new Uint8Array([1]), count: recs.length, skipped: 0 }));
    const onClose = vi.fn();
    const c = mount(h(MergeDialog, { fields: [{ key: '名称', count: 1 }], tabPath: 'D:\\合同\\模板.docx', generate, generateAll, onInsertField: noop, onInsertNext: noop, onClose }));
    const csv = new File(['名称,金额\n甲,100\n乙,200\n'], 'data.csv', { type: 'text/csv' });
    const input = c.querySelector('input[type=file]') as HTMLInputElement;
    Object.defineProperty(input, 'files', { value: [csv], configurable: true });
    await act(async () => {
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await new Promise((r) => setTimeout(r, 50));
    });
    expect(c.textContent).toContain('2 条数据');
    const nameInput = [...c.querySelectorAll('input.field')].find((i) => (i as HTMLInputElement).value.includes('{#}')) as HTMLInputElement;
    setValue(nameInput, '合同_{{名称}}');
    await act(async () => {
      click(button(c, '生成 2 份文档'));
      await new Promise((r) => setTimeout(r, 50));
    });
    expect(generate).toHaveBeenCalledTimes(2);
    expect(written).toEqual(['D:\\合同\\合同_甲 (2).docx', 'D:\\合同\\合同_乙.docx']);
    expect(onClose).toHaveBeenCalled();
  });

  it('模板与批量生成：合并成一个文档只写一个文件', async () => {
    const written: string[] = [];
    fakeHost({ writeBytes: async (p: string) => void written.push(p) });
    const generate = vi.fn();
    const generateAll = vi.fn(async (recs: Record<string, string>[]) => ({ bytes: new Uint8Array([1]), count: recs.length, skipped: 0 }));
    const c = mount(h(MergeDialog, { fields: [{ key: '名称', count: 1 }], tabPath: 'D:\\合同\\模板.docx', generate, generateAll, onInsertField: noop, onInsertNext: noop, onClose: noop }));
    const input = c.querySelector('input[type=file]') as HTMLInputElement;
    Object.defineProperty(input, 'files', { value: [new File(['名称\n甲\n乙\n丙\n'], 'd.csv')], configurable: true });
    await act(async () => {
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await new Promise((r) => setTimeout(r, 50));
    });
    click(c.querySelectorAll('input[type=radio]')[1]);
    await act(async () => {
      click(button(c, '合并 3 条数据'));
      await new Promise((r) => setTimeout(r, 50));
    });
    expect(generate).not.toHaveBeenCalled();
    expect(generateAll).toHaveBeenCalledTimes(1);
    expect(written).toEqual(['D:\\合同\\模板_1.docx']);
  });

  it('限制编辑：两次密码不一致不能启动；启动带出方式和密码；已保护时要密码才能停止；水印可切到图片', async () => {
    const onStart = vi.fn(async () => null);
    const onStop = vi.fn(async (pw: string) => (pw === 'ok' ? null : '密码不对'));
    const c = mount(h(ProtectDialog, { hasSelection: true, permCount: 1, onStart, onStop, onAddPerm: noop, onClearPerms: noop, onClose: noop }));
    const [p1, p2] = [...c.querySelectorAll('input[type=password]')] as HTMLInputElement[];
    setValue(p1, 'abc');
    setValue(p2, 'abd');
    click(button(c, '启动保护'));
    expect(onStart).not.toHaveBeenCalled();
    expect(c.textContent).toContain('两次输入的密码不一样');
    setValue(p2, 'abc');
    setValue(c.querySelector('select') as HTMLSelectElement, 'comments');
    await act(async () => {
      click(button(c, '启动保护'));
      await new Promise((r) => setTimeout(r, 10));
    });
    expect(onStart).toHaveBeenCalledWith('comments', 'abc');
    act(() => root!.render(h(ProtectDialog, { protection: { edit: 'readOnly', enforced: true, hash: 'x' }, hasSelection: false, permCount: 2, onStart, onStop, onAddPerm: noop, onClearPerms: noop, onClose: noop })));
    expect(c.textContent).toContain('只读');
    expect(c.textContent).toContain('2 处允许编辑的区域');
    setValue(c.querySelector('input[type=password]') as HTMLInputElement, 'bad');
    await act(async () => {
      click(button(c, '停止保护'));
      await new Promise((r) => setTimeout(r, 10));
    });
    expect(c.textContent).toContain('密码不对');
  });

  it('水印：图片水印没选图片不能确定；切换后能回到文字水印', () => {
    const onOk = vi.fn();
    const c = mount(h(WatermarkDialog, { onOk, onClose: noop }));
    click(button(c, '图片水印'));
    expect((button(c, '确定') as HTMLButtonElement).disabled).toBe(true);
    click(button(c, '文字水印'));
    click(button(c, '确定'));
    expect(onOk).toHaveBeenCalledWith(expect.objectContaining({ text: '机密' }));
  });
});

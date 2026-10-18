/** 壳子的状态：打开的标签页、弹窗、提示。很小，不引入状态库 */
import { useSyncExternalStore } from 'react';

export type Kind = 'word' | 'sheet';

export interface Tab {
  id: string;
  kind: Kind;
  title: string;
  /** 内存文件系统里的路径；新建、还没保存的文件没有 */
  path?: string;
  dirty?: boolean;
}

/** 输入框 / 确认框（代替 window.prompt / confirm） */
export type Dialog =
  | { type: 'ask'; title: string; value: string; resolve: (v: string | null) => void }
  | { type: 'confirm'; message: string; ok: string; resolve: (v: boolean) => void };

export interface State {
  tabs: Tab[];
  active: string | null;
  dialogs: Dialog[];
  toast: { id: number; msg: string } | null;
  userName: string;
}

let state: State = { tabs: [], active: null, dialogs: [], toast: null, userName: readUserName() };
const subs = new Set<() => void>();

export const getState = () => state;

export function setState(patch: Partial<State> | ((s: State) => Partial<State>)) {
  state = { ...state, ...(typeof patch === 'function' ? patch(state) : patch) };
  subs.forEach((f) => f());
}

export function subscribe(cb: () => void): () => void {
  subs.add(cb);
  return () => void subs.delete(cb);
}

export function useStore<T>(pick: (s: State) => T): T {
  return useSyncExternalStore(subscribe, () => pick(state));
}

let seq = 0;
export const newId = () => `t${Date.now().toString(36)}${(seq++).toString(36)}`;

export function openTab(t: Omit<Tab, 'id'>) {
  // 同一个文件已经打开：切过去
  const same = t.path && state.tabs.find((x) => x.path === t.path);
  if (same) return setState({ active: same.id });
  const tab = { ...t, id: newId() };
  setState((s) => ({ tabs: [...s.tabs, tab], active: tab.id }));
}

export function patchTab(id: string, patch: Partial<Tab>) {
  setState((s) => ({ tabs: s.tabs.map((t) => (t.id === id ? { ...t, ...patch } : t)) }));
}

export function removeTab(id: string) {
  setState((s) => {
    const i = s.tabs.findIndex((t) => t.id === id);
    const tabs = s.tabs.filter((t) => t.id !== id);
    const active = s.active === id ? (tabs[Math.min(i, tabs.length - 1)]?.id ?? null) : s.active;
    return { tabs, active };
  });
}

let toastSeq = 0;
let toastTimer: ReturnType<typeof setTimeout> | undefined;
export function showToast(msg: string, ms = 3000) {
  const id = ++toastSeq;
  setState({ toast: { id, msg } });
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => getState().toast?.id === id && setState({ toast: null }), ms);
}

export function pushDialog(d: Dialog) {
  setState((s) => ({ dialogs: [...s.dialogs, d] }));
}

export function popDialog(d: Dialog) {
  setState((s) => ({ dialogs: s.dialogs.filter((x) => x !== d) }));
}

function readUserName(): string {
  try {
    return localStorage.getItem('xioffice.user') || '用户';
  } catch {
    return '用户';
  }
}

export function setUserName(name: string) {
  try {
    localStorage.setItem('xioffice.user', name);
  } catch {
    // 隐私模式下存不了，只在本次会话里生效
  }
  setState({ userName: name });
}

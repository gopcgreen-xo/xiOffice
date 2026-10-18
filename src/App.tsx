import { Suspense, lazy, useEffect, useRef, useState, type ComponentType, type DragEvent } from 'react';
import type { EditorTab } from '@wz/editor-shared';
import { DIR, EDITORS, chooseAndOpen, closeTab, newDocument, openFiles } from './host';
import { desktop } from './desktop';
import { openTab, setState, useStore, type Dialog, type Kind } from './store';
import sampleLayout from '../editors/word/samples/排版对照样例.docx?url';
import sampleWord from '../editors/word/samples/Word生成的综合样本.docx?url';
import sampleSheet from '../editors/sheet/samples/报价单样例.xlsx?url';

const ISSUES = 'https://github.com/gopcgreen-xo/xiOffice/issues';

// 编辑器第一次打开时才加载
const LAZY: Record<Kind, ComponentType<{ tab: EditorTab }>> = {
  word: lazy(() => import('../editors/word/src').then((m) => ({ default: m.Editor }))),
  sheet: lazy(() => import('../editors/sheet/src').then((m) => ({ default: m.Editor }))),
};

const SAMPLES = [
  { name: '排版对照样例.docx', url: sampleLayout },
  { name: 'Word生成的综合样本.docx', url: sampleWord },
  { name: '报价单样例.xlsx', url: sampleSheet },
];

async function openSample(name: string, url: string) {
  const data = await (await fetch(url)).arrayBuffer();
  await openFiles([new File([data], name)]);
}


export function App() {
  const tabs = useStore((s) => s.tabs);
  const active = useStore((s) => s.active);
  const [drag, setDrag] = useState(false);

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDrag(false);
    void openFiles(e.dataTransfer.files);
  };

  return (
    <main
      className="center xo-app"
      onDragOver={(e) => {
        if (![...e.dataTransfer.types].includes('Files')) return;
        e.preventDefault();
        setDrag(true);
      }}
      onDragLeave={(e) => e.currentTarget === e.target && setDrag(false)}
      onDrop={onDrop}
    >
      <div className="tabbar" role="tablist" aria-label="已打开的文件">
        <button className="xo-brand" title="首页" onClick={() => setState({ active: null })}>
          xiOffice
        </button>
        <div className="tab-strip">
          {tabs.map((t) => (
            <div key={t.id} className={`tab${t.id === active ? ' on' : ''}`} role="tab" aria-selected={t.id === active}>
              <button onClick={() => setState({ active: t.id })} title={t.path ?? t.title}>
                <span className={`xo-badge ${t.kind}`}>{t.kind === 'word' ? 'W' : 'X'}</span>
                <span className="ellipsis">{t.title}</span>
                {t.dirty && <span className="dirty" aria-label="未保存" />}
              </button>
              <button className="icon-btn tab-close" aria-label="关闭标签" title="关闭" onClick={() => void closeTab(t.id)}>
                ×
              </button>
            </div>
          ))}
        </div>
        <div className="xo-actions">
          <button className="btn" onClick={() => newDocument('word')}>新建文档</button>
          <button className="btn" onClick={() => newDocument('sheet')}>新建表格</button>
          <button className="btn primary" onClick={() => void chooseAndOpen()}>
            打开文件
          </button>
        </div>
      </div>
      <div className="content">
        {!active && <Home />}
        {tabs.map((t) => {
          const Editor = LAZY[t.kind];
          return (
            <section key={t.id} className="pane" hidden={t.id !== active}>
              <Suspense fallback={<div className="xo-loading">正在加载编辑器…</div>}>
                <Editor tab={t} />
              </Suspense>
            </section>
          );
        })}
      </div>
      {drag && <div className="xo-drop">松开鼠标打开文件</div>}
      <Dialogs />
      <Toast />
    </main>
  );
}

function Home() {
  return (
    <div className="xo-home">
      <h1>xiOffice</h1>
      <p className="xo-sub">编辑 Word（.docx / .doc）和 Excel（.xlsx / .csv），文件不离开本机。</p>
      <div className="xo-cards">
        {EDITORS.map((e) => (
          <button key={e.kind} className="xo-card" onClick={() => newDocument(e.kind)}>
            <span className={`xo-badge big ${e.kind}`}>{e.kind === 'word' ? 'W' : 'X'}</span>
            <strong>{e.create.name}</strong>
            <span className="xo-sub">支持 {e.exts.map((x) => `.${x}`).join(' ')}</span>
          </button>
        ))}
        <button className="xo-card" onClick={() => void chooseAndOpen()}>
          <span className="xo-badge big open">↑</span>
          <strong>打开本机文件</strong>
          <span className="xo-sub">也可以直接把文件拖进窗口</span>
        </button>
      </div>
      <h2>示例文件</h2>
      <div className="xo-samples">
        {SAMPLES.map((s) => (
          <button key={s.name} className="btn" onClick={() => void openSample(s.name, s.url)}>
            {s.name}
          </button>
        ))}
      </div>
      <p className="xo-note">
        {desktop ? (
          <>保存（Ctrl+S）直接写回原文件；新建的文件和示例第一次保存时选择保存位置。</>
        ) : (
          <>
            浏览器版：保存（Ctrl+S）会把文件下载到本机；文件只保存在这个页面的内存里（虚拟文件夹 <code>{DIR}</code>），刷新页面后需要重新打开。
          </>
        )}
        <br />
        遇到打不开、排版不对、保存后 Office 报错等问题，欢迎 <a href={ISSUES} target="_blank" rel="noreferrer">提交 Issue</a>，最好附上能复现问题的文件（请先去掉敏感内容）。
      </p>
    </div>
  );
}

function Dialogs() {
  const d = useStore((s) => s.dialogs[s.dialogs.length - 1]);
  if (!d) return null;
  return (
    <div className="xo-backdrop">
      <DialogBox key={d.type === 'ask' ? d.title : d.message} d={d} />
    </div>
  );
}

function DialogBox({ d }: { d: Dialog }) {
  const [value, setValue] = useState(d.type === 'ask' ? d.value : '');
  const input = useRef<HTMLInputElement>(null);
  const okBtn = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    if (d.type === 'ask') {
      input.current?.focus();
      // 文件名只选中扩展名前的部分
      const dot = d.value.lastIndexOf('.');
      input.current?.setSelectionRange(0, dot > 0 ? dot : d.value.length);
    } else okBtn.current?.focus();
  }, [d]);
  const cancel = () => (d.type === 'ask' ? d.resolve(null) : d.resolve(false));
  const ok = () => (d.type === 'ask' ? d.resolve(value.trim() || null) : d.resolve(true));
  return (
    <form
      className="xo-dialog"
      role="dialog"
      aria-modal="true"
      onSubmit={(e) => (e.preventDefault(), ok())}
      onKeyDown={(e) => e.key === 'Escape' && (e.stopPropagation(), cancel())}
    >
      {d.type === 'ask' ? (
        <>
          <label>{d.title}</label>
          <input ref={input} value={value} onChange={(e) => setValue(e.target.value)} />
        </>
      ) : (
        <p>{d.message}</p>
      )}
      <div className="xo-dialog-buttons">
        <button type="button" className="btn" onClick={cancel}>
          取消
        </button>
        <button ref={okBtn} type="submit" className="btn primary">
          {d.type === 'confirm' ? d.ok : '确定'}
        </button>
      </div>
    </form>
  );
}

function Toast() {
  const t = useStore((s) => s.toast);
  return t ? <div className="toast" role="status">{t.msg}</div> : null;
}

// 地址栏 ?new=word / ?new=sheet：直接新建；?sample=<示例文件名>：打开示例（方便复现问题时贴链接）
const params = new URLSearchParams(location.search);
const q = params.get('new');
if (q === 'word' || q === 'sheet') openTab({ kind: q, title: EDITORS.find((e) => e.kind === q)!.create.title });
const sample = SAMPLES.find((s) => s.name === params.get('sample'));
if (sample) void openSample(sample.name, sample.url);

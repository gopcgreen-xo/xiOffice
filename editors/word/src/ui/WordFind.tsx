/** Word 编辑器的查找和替换（Ctrl+F / Ctrl+H）：浮在文档右上角，查找范围是正文、表格和页眉页脚。 */
import { useEffect, useRef, useState } from 'react';
import { findAll, findNext, type FindOpts, type Match } from '../core/find';
import type { Doc } from '../core/model';
import type { Pos } from '../core/layout';
import { Icon } from '@wz/editor-shared/Icon';

interface Props {
  doc: Doc;
  /** 当前光标（从它后面开始找） */
  at: Pos;
  /** 当前选中的是不是某个匹配（替换按钮先替换它） */
  current: Match | null;
  /** 从工具条输入框带进来的查找内容 */
  initialQuery?: string;
  mode: 'find' | 'replace';
  version: number;
  onMode: (m: 'find' | 'replace') => void;
  onSelect: (m: Match) => void;
  onReplace: (m: Match, repl: string) => void;
  onReplaceAll: (query: string, repl: string, o: FindOpts) => number;
  onClose: () => void;
}

export function WordFind(p: Props) {
  const [query, setQuery] = useState(p.initialQuery ?? '');
  const [repl, setRepl] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [msg, setMsg] = useState('');
  const input = useRef<HTMLInputElement>(null);
  const opts: FindOpts = { matchCase, whole };

  useEffect(() => {
    input.current?.focus();
    input.current?.select();
  }, [p.mode]);
  useEffect(() => setMsg(''), [query, matchCase, whole]);

  const next = (back = false) => {
    const all = findAll(p.doc, query, opts);
    if (!all.length) return setMsg('找不到要查找的内容');
    const m = findNext(all, p.at, back)!;
    setMsg(`第 ${all.indexOf(m) + 1} / ${all.length} 处`);
    p.onSelect(m);
  };
  const replaceOne = () => {
    const all = findAll(p.doc, query, opts);
    const cur = p.current && all.find((m) => m.story === p.current!.story && m.offset === p.current!.offset && m.path.join() === p.current!.path.join());
    if (!cur) return next();
    p.onReplace(cur, repl);
    // 替换后跳到下一处（同 Word）
    next();
  };
  const replaceEvery = () => {
    const n = p.onReplaceAll(query, repl, opts);
    setMsg(n ? `已替换 ${n} 处（可以撤销）` : '找不到要替换的内容');
  };

  return (
    <div className="ed-find word-find" role="dialog" aria-label="查找和替换" onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
      <div className="ed-find-head">
        <button className={p.mode === 'find' ? 'on' : ''} onClick={() => p.onMode('find')}>
          查找
        </button>
        <button className={p.mode === 'replace' ? 'on' : ''} onClick={() => p.onMode('replace')}>
          替换
        </button>
        <span className="grow" />
        <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
          <Icon name="close" size={12} />
        </button>
      </div>
      <div className="ed-find-row">
        <span className="sub">查找</span>
        <input
          ref={input}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              next(e.shiftKey);
            }
          }}
        />
      </div>
      {p.mode === 'replace' && (
        <div className="ed-find-row">
          <span className="sub">替换为</span>
          <input
            value={repl}
            onChange={(e) => setRepl(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                replaceOne();
              }
            }}
          />
        </div>
      )}
      <div className="ed-find-opts">
        <label>
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
          区分大小写
        </label>
        <label>
          <input type="checkbox" checked={whole} onChange={(e) => setWhole(e.target.checked)} />
          全字匹配
        </label>
        <span className="sub">范围：正文、表格、页眉页脚</span>
      </div>
      <div className="ed-find-foot">
        <span className="sub ellipsis grow">{msg}</span>
        {p.mode === 'replace' && (
          <>
            <button className="btn" disabled={!query} onClick={replaceEvery}>
              全部替换
            </button>
            <button className="btn" disabled={!query} onClick={replaceOne}>
              替换
            </button>
          </>
        )}
        <button className="btn" disabled={!query} onClick={() => next(true)}>
          上一个
        </button>
        <button className="btn primary" disabled={!query} onClick={() => next()}>
          下一个
        </button>
      </div>
    </div>
  );
}

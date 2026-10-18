/** 查找和替换（Ctrl+F / Ctrl+H）：浮在表格右上角，不挡住表格操作。 */
import { useEffect, useRef, useState } from 'react';
import { host } from '@wz/editor-shared';
import { findAll, findNext, type FindOptions, type Hit } from '../core/find';
import type { Workbook } from '../core/workbook';
import { colName } from '@wz/editor-shared/xlsx';
import { Icon } from '@wz/editor-shared/Icon';

interface Props {
  book: Workbook;
  si: number;
  /** 当前活动单元格（从它后面开始找） */
  at: Hit;
  mode: 'find' | 'replace';
  /** 模型改动计数：改了内容后结果要重新找 */
  version: number;
  onMode: (m: 'find' | 'replace') => void;
  onGoto: (h: Hit) => void;
  onReplace: (h: Hit, o: FindOptions, repl: string) => boolean;
  onReplaceAll: (o: FindOptions, repl: string) => number;
  onClose: () => void;
}

export function FindReplace(p: Props) {
  const [query, setQuery] = useState('');
  const [repl, setRepl] = useState('');
  const [book, setBook] = useState(false);
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [lookIn, setLookIn] = useState<'formulas' | 'values'>('formulas');
  const [list, setList] = useState<Hit[] | null>(null);
  const [msg, setMsg] = useState('');
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => {
    input.current?.focus();
    input.current?.select();
  }, [p.mode]);
  useEffect(() => setList(null), [query, book, matchCase, whole, lookIn, p.version]);

  const opts = (): FindOptions => ({ query, matchCase, whole, lookIn: p.mode === 'replace' ? 'formulas' : lookIn, si: book ? undefined : p.si });
  const next = (back = false) => {
    const hits = findAll(p.book, opts());
    if (!hits.length) return setMsg('找不到要查找的内容');
    const h = findNext(hits, p.at, back)!;
    setMsg(`第 ${hits.findIndex((x) => x.si === h.si && x.r === h.r && x.c === h.c) + 1} / ${hits.length} 个`);
    p.onGoto(h);
  };
  const replaceOne = () => {
    const o = opts();
    const hits = findAll(p.book, o);
    if (!hits.length) return setMsg('找不到要替换的内容');
    // 活动单元格匹配就先替换它，然后跳到下一个（同 Excel）
    const cur = hits.find((h) => h.si === p.at.si && h.r === p.at.r && h.c === p.at.c);
    if (cur) {
      p.onReplace(cur, o, repl);
      const rest = findAll(p.book, o);
      const h = findNext(rest, cur);
      if (h) p.onGoto(h);
      setMsg(rest.length ? `还有 ${rest.length} 个` : '已全部替换');
    } else p.onGoto(findNext(hits, p.at)!);
  };
  const replaceAll = () => {
    const n = p.onReplaceAll(opts(), repl);
    setMsg(n ? `已替换 ${n} 处` : '找不到要替换的内容');
    if (n) host().toast(`已替换 ${n} 处（可以撤销）`);
  };

  return (
    <div className="ed-find" role="dialog" aria-label="查找和替换" onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
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
          <input value={repl} onChange={(e) => setRepl(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), replaceOne())} />
        </div>
      )}
      <div className="ed-find-opts">
        <select className="ed-select" value={book ? 'book' : 'sheet'} onChange={(e) => setBook(e.target.value === 'book')}>
          <option value="sheet">当前工作表</option>
          <option value="book">整个工作簿</option>
        </select>
        {p.mode === 'find' && (
          <select className="ed-select" value={lookIn} onChange={(e) => setLookIn(e.target.value as 'formulas' | 'values')}>
            <option value="formulas">查找公式</option>
            <option value="values">查找值</option>
          </select>
        )}
        <label>
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
          区分大小写
        </label>
        <label>
          <input type="checkbox" checked={whole} onChange={(e) => setWhole(e.target.checked)} />
          单元格匹配
        </label>
      </div>
      <div className="ed-find-foot">
        <span className="sub ellipsis grow">{msg}</span>
        {p.mode === 'replace' ? (
          <>
            <button className="btn" disabled={!query} onClick={replaceAll}>
              全部替换
            </button>
            <button className="btn" disabled={!query} onClick={replaceOne}>
              替换
            </button>
          </>
        ) : (
          <button className="btn" disabled={!query} onClick={() => setList(findAll(p.book, opts()))}>
            查找全部
          </button>
        )}
        <button className="btn primary" disabled={!query} onClick={() => next()}>
          查找下一个
        </button>
      </div>
      {list && (
        <div className="ed-find-list">
          <div className="sub">共 {list.length} 个</div>
          {list.slice(0, 500).map((h) => (
            <button key={`${h.si}:${h.r}:${h.c}`} onClick={() => p.onGoto(h)}>
              <span className="mono">
                {book ? `${p.book.sheets[h.si].name}!` : ''}
                {colName(h.c)}
                {h.r + 1}
              </span>
              <span className="ellipsis">{p.book.sheets[h.si].cells[h.r]?.[h.c]?.t}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

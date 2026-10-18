/**
 * 单元格里的选择器：单选（可输入筛选）、多选（勾选，确定后写成逗号分隔文字）、日期（月历）。
 * 只在工作包里有字段类型的列出现；位置由表格视图给（贴着活动单元格下方）。
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import { OPTION_PALETTE, formatDate, joinMulti, parseDateText, splitMulti, type ColumnMeta, type SelectOption } from '@wz/editor-shared';
import './columns.css';

export interface PickerProps {
  meta: ColumnMeta;
  options: SelectOption[];
  /** 单元格现在的文字 */
  value: string;
  at: { left: number; top: number };
  /** 写入单元格 */
  onPick: (text: string) => void;
  /** 输入了不在选项里的文字并选择“新建选项”：由表格视图把它加进列属性 */
  onCreate?: (text: string) => void;
  /** 打开列属性（管理选项） */
  onManage?: () => void;
  onClose: () => void;
}

const colorOf = (options: SelectOption[], v: string): string => {
  const i = options.findIndex((o) => o.v === v);
  return i < 0 ? '#94a3b8' : (options[i].color ?? OPTION_PALETTE[i % OPTION_PALETTE.length]);
};

export function ColumnPicker(props: PickerProps) {
  if (props.meta.type === 'date') return <DatePicker {...props} />;
  return <ListPicker {...props} multi={props.meta.type === 'multi'} />;
}

function useOutside(ref: React.RefObject<HTMLElement | null>, onClose: () => void) {
  useEffect(() => {
    // 捕获阶段：表格区域会拦截鼠标事件，冒泡阶段收不到
    const down = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose();
    };
    document.addEventListener('mousedown', down, true);
    return () => document.removeEventListener('mousedown', down, true);
  }, [ref, onClose]);
}

function ListPicker({ meta, options, value, at, onPick, onCreate, onManage, onClose, multi }: PickerProps & { multi: boolean }) {
  const root = useRef<HTMLDivElement>(null);
  const input = useRef<HTMLInputElement>(null);
  const [q, setQ] = useState('');
  const [idx, setIdx] = useState(0);
  const [picked, setPicked] = useState<string[]>(() => (multi ? splitMulti(value) : []));
  useOutside(root, onClose);
  useEffect(() => input.current?.focus(), []);

  const list = useMemo(() => {
    const t = q.trim().toLowerCase();
    return t ? options.filter((o) => o.v.toLowerCase().includes(t)) : options;
  }, [options, q]);
  const canCreate = !!onCreate && q.trim() !== '' && !options.some((o) => o.v === q.trim());
  const rows = canCreate ? list.length + 1 : list.length;

  const toggle = (v: string) => setPicked((p) => (p.includes(v) ? p.filter((x) => x !== v) : [...p, v]));
  const choose = (i: number) => {
    if (i >= list.length) {
      const text = q.trim();
      onCreate?.(text);
      if (multi) {
        setPicked((p) => [...p, text]);
        setQ('');
      } else onPick(text);
      return;
    }
    if (multi) toggle(list[i].v);
    else onPick(list[i].v);
  };

  return (
    <div ref={root} className="pk" style={at} role="listbox" aria-multiselectable={multi} aria-label={meta.title || '选项'} onMouseDown={(e) => e.stopPropagation()}>
      <input
        ref={input}
        className="pk-input"
        aria-label="筛选选项"
        placeholder={multi ? '输入筛选或新建选项' : '输入筛选'}
        value={q}
        onChange={(e) => (setQ(e.target.value), setIdx(0))}
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            e.preventDefault();
            e.stopPropagation();
            onClose();
          } else if (e.key === 'ArrowDown') {
            e.preventDefault();
            setIdx((i) => Math.min(Math.max(0, rows - 1), i + 1));
          } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            setIdx((i) => Math.max(0, i - 1));
          } else if (e.key === 'Enter') {
            e.preventDefault();
            if (multi && (e.ctrlKey || q === '')) onPick(joinMulti(picked));
            else if (rows) choose(Math.min(idx, rows - 1));
            else if (multi) onPick(joinMulti(picked));
          }
        }}
      />
      <div className="pk-list">
        {list.map((o, i) => {
          const on = multi ? picked.includes(o.v) : o.v === value;
          return (
            <button key={o.v} role="option" aria-selected={on} className={`pk-opt${i === idx ? ' hot' : ''}${on ? ' on' : ''}`} onMouseEnter={() => setIdx(i)} onClick={() => choose(i)}>
              {multi && <span className={`pk-box${on ? ' on' : ''}`}>{on ? '✓' : ''}</span>}
              <span className="pk-dot" style={{ background: colorOf(options, o.v) }} />
              <span className="pk-text">{o.v}</span>
              {!multi && on && <span className="pk-tick">✓</span>}
            </button>
          );
        })}
        {canCreate && (
          <button className={`pk-opt create${idx === list.length ? ' hot' : ''}`} onMouseEnter={() => setIdx(list.length)} onClick={() => choose(list.length)}>
            + 新建选项「{q.trim()}」
          </button>
        )}
        {!rows && <div className="pk-empty">{options.length ? '没有匹配的选项' : '还没有选项'}</div>}
      </div>
      {multi ? (
        <div className="pk-foot">
          <span className="sub">已选 {picked.length} 项</span>
          <span className="grow" />
          {onManage && (
            <button className="cp-link" onClick={onManage}>
              管理选项
            </button>
          )}
          <button className="btn primary sm" onClick={() => onPick(joinMulti(picked))}>
            确定
          </button>
        </div>
      ) : (
        <div className="pk-foot">
          <span className="sub">{rows ? '↑↓ 选择，回车确认' : ''}</span>
          <span className="grow" />
          {value && (
            <button className="cp-link" onClick={() => onPick('')}>
              清除
            </button>
          )}
          {onManage && (
            <button className="cp-link" onClick={onManage}>
              管理选项
            </button>
          )}
        </div>
      )}
    </div>
  );
}

const WEEK = ['一', '二', '三', '四', '五', '六', '日'];

function DatePicker({ value, at, onPick, onClose }: PickerProps) {
  const root = useRef<HTMLDivElement>(null);
  useOutside(root, onClose);
  const cur = parseDateText(value);
  const today = new Date();
  const [month, setMonth] = useState(() => new Date((cur ?? today).getFullYear(), (cur ?? today).getMonth(), 1));
  const [text, setText] = useState(value);

  // 6 周 × 7 天，周一开头
  const days = useMemo(() => {
    const first = new Date(month.getFullYear(), month.getMonth(), 1);
    const lead = (first.getDay() + 6) % 7;
    return Array.from({ length: 42 }, (_, i) => new Date(month.getFullYear(), month.getMonth(), 1 - lead + i));
  }, [month]);
  const same = (a: Date | null, b: Date) => !!a && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  const go = (n: number) => setMonth(new Date(month.getFullYear(), month.getMonth() + n, 1));

  return (
    <div ref={root} className="pk pk-date" style={at} role="dialog" aria-label="选择日期" onMouseDown={(e) => e.stopPropagation()}>
      <div className="pk-dhead">
        <button aria-label="上个月" onClick={() => go(-1)}>
          ‹
        </button>
        <span>
          {month.getFullYear()} 年 {month.getMonth() + 1} 月
        </span>
        <button aria-label="下个月" onClick={() => go(1)}>
          ›
        </button>
      </div>
      <div className="pk-grid">
        {WEEK.map((w) => (
          <span key={w} className="pk-wk">
            {w}
          </span>
        ))}
        {days.map((d) => (
          <button
            key={d.getTime()}
            className={`pk-day${d.getMonth() !== month.getMonth() ? ' dim' : ''}${same(cur, d) ? ' on' : ''}${same(today, d) ? ' today' : ''}`}
            onClick={() => onPick(formatDate(d))}
          >
            {d.getDate()}
          </button>
        ))}
      </div>
      <div className="pk-foot">
        <input
          className="pk-input sm"
          aria-label="日期"
          placeholder="2026-10-06"
          value={text}
          autoFocus
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Escape') {
              e.preventDefault();
              e.stopPropagation();
              onClose();
            } else if (e.key === 'Enter') {
              e.preventDefault();
              const d = parseDateText(text);
              if (d) onPick(formatDate(d));
              else if (text.trim() === '') onPick('');
            }
          }}
        />
        <button className="cp-link" onClick={() => onPick(formatDate(today))}>
          今天
        </button>
        <button className="cp-link" onClick={() => onPick('')}>
          清除
        </button>
      </div>
    </div>
  );
}

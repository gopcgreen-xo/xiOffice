/**
 * 自动筛选的列下拉：升序 / 降序、按值勾选（带搜索）、一个自定义条件（等于 / 包含 / 大于 …）、清除。
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import { host } from '@wz/editor-shared';
import type { ColumnFilter } from '../core/workbook';

type Op = 'equal' | 'notEqual' | 'contains' | 'notContains' | 'begins' | 'ends' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual';

const OPS: { op: Op; label: string }[] = [
  { op: 'equal', label: '等于' },
  { op: 'notEqual', label: '不等于' },
  { op: 'contains', label: '包含' },
  { op: 'notContains', label: '不包含' },
  { op: 'begins', label: '开头是' },
  { op: 'ends', label: '结尾是' },
  { op: 'greaterThan', label: '大于' },
  { op: 'greaterThanOrEqual', label: '大于或等于' },
  { op: 'lessThan', label: '小于' },
  { op: 'lessThanOrEqual', label: '小于或等于' },
];

/** 界面上的条件 → 文件里的 customFilter（包含 = 等于 *x*，同 Excel） */
function toRule(op: Op, val: string): NonNullable<ColumnFilter['custom']>['rules'][number] {
  switch (op) {
    case 'contains':
      return { op: 'equal', val: `*${val}*` };
    case 'notContains':
      return { op: 'notEqual', val: `*${val}*` };
    case 'begins':
      return { op: 'equal', val: `${val}*` };
    case 'ends':
      return { op: 'equal', val: `*${val}` };
    default:
      return { op, val };
  }
}

function fromRule(r: NonNullable<ColumnFilter['custom']>['rules'][number]): { op: Op; val: string } {
  const m = r.val.match(/^\*(.*)\*$/);
  if (r.op === 'equal' && m) return { op: 'contains', val: m[1] };
  if (r.op === 'notEqual' && m) return { op: 'notContains', val: m[1] };
  if (r.op === 'equal' && r.val.endsWith('*')) return { op: 'begins', val: r.val.slice(0, -1) };
  if (r.op === 'equal' && r.val.startsWith('*')) return { op: 'ends', val: r.val.slice(1) };
  return { op: r.op, val: r.val };
}

interface Props {
  x: number;
  y: number;
  /** 不是日期的格子的值 */
  values: { text: string; count: number }[];
  /** 日期格子：年 → 月 → 个数 */
  dates: Map<number, Map<number, number>>;
  /** 出现的底色 / 字体颜色（null = 无填充 / 自动）→ 个数 */
  fills: Map<string | null, number>;
  colors: Map<string | null, number>;
  current?: ColumnFilter;
  onSort: (desc: boolean) => void;
  onApply: (f: ColumnFilter | null) => void;
  onClose: () => void;
}

export function FilterMenu(p: Props) {
  const all = useMemo(() => p.values.map((v) => v.text), [p.values]);
  const [checked, setChecked] = useState<Set<string>>(() => new Set(p.current?.values ?? all));
  const [query, setQuery] = useState('');
  const first = p.current?.custom?.rules[0] ? fromRule(p.current.custom.rules[0]) : null;
  const [op, setOp] = useState<Op>(first?.op ?? 'contains');
  const [val, setVal] = useState(first?.val ?? '');
  const [mode, setMode] = useState<'values' | 'custom' | 'color'>(p.current?.custom ? 'custom' : p.current?.color ? 'color' : 'values');
  // 日期勾选：键 "年" 或 "年-月"；默认全选，有条件时按条件
  const allMonths = useMemo(() => [...p.dates.entries()].flatMap(([y, ms]) => [...ms.keys()].map((m) => `${y}-${m}`)), [p.dates]);
  const [months, setMonths] = useState<Set<string>>(() => {
    const ds = p.current?.dates;
    if (!p.current || (!ds && !p.current.values)) return new Set(allMonths);
    return new Set(allMonths.filter((k) => {
      const [y, m] = k.split('-').map(Number);
      return ds?.some((g) => g.y === y && (g.m === undefined || g.m === m));
    }));
  });
  const toggleMonths = (keys: string[]) => {
    const on = keys.every((k) => months.has(k));
    const next = new Set(months);
    for (const k of keys) {
      if (on) next.delete(k);
      else next.add(k);
    }
    setMonths(next);
  };
  const showColors = p.fills.size > 1 || p.colors.size > 1;
  const box = useRef<HTMLDivElement>(null);

  useEffect(() => {
    host().pushPopup();
    const close = (e: MouseEvent) => {
      if (!box.current?.contains(e.target as Node)) p.onClose();
    };
    window.addEventListener('mousedown', close);
    return () => {
      host().popPopup();
      window.removeEventListener('mousedown', close);
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const shown = p.values.filter((v) => !query || v.text.toLowerCase().includes(query.toLowerCase()));
  const allOn = shown.every((v) => checked.has(v.text));
  const toggleAll = () => {
    const next = new Set(checked);
    for (const v of shown) {
      if (allOn) next.delete(v.text);
      else next.add(v.text);
    }
    setChecked(next);
  };
  const apply = () => {
    if (mode === 'custom') {
      if (!val.trim()) p.onApply(null);
      else p.onApply({ custom: { and: true, rules: [toRule(op, val.trim())] } });
      return;
    }
    // 搜索时只保留搜到的（同 Excel）
    const picked = query ? shown.filter((v) => checked.has(v.text)).map((v) => v.text) : all.filter((t) => checked.has(t));
    const allDates = allMonths.every((k) => months.has(k));
    if (picked.length === all.length && allDates) return p.onApply(null);
    if (!allMonths.length) return p.onApply({ values: picked });
    // 一年的月份全选时记成整年
    const dates: { y: number; m?: number }[] = [];
    for (const [y, ms] of p.dates) {
      const keys = [...ms.keys()].map((m) => `${y}-${m}`);
      if (keys.every((k) => months.has(k))) dates.push({ y });
      else for (const m of ms.keys()) if (months.has(`${y}-${m}`)) dates.push({ y, m });
    }
    p.onApply({ values: picked, dates });
  };

  return (
    <div ref={box} className="ctx-menu sheet-filter" role="dialog" aria-label="筛选" style={{ left: p.x, top: p.y }} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
      <button role="menuitem" onClick={() => p.onSort(false)}>
        升序
      </button>
      <button role="menuitem" onClick={() => p.onSort(true)}>
        降序
      </button>
      <div className="ctx-sep" />
      <div className="sheet-filter-tabs">
        <button className={mode === 'values' ? 'on' : ''} onClick={() => setMode('values')}>
          按值筛选
        </button>
        <button className={mode === 'custom' ? 'on' : ''} onClick={() => setMode('custom')}>
          条件筛选
        </button>
        {showColors && (
          <button className={mode === 'color' ? 'on' : ''} onClick={() => setMode('color')}>
            按颜色
          </button>
        )}
      </div>
      {mode === 'values' ? (
        <>
          <input className="sheet-filter-search" placeholder="搜索" value={query} autoFocus onChange={(e) => setQuery(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), apply())} />
          <div className="sheet-filter-list">
            {[...p.dates.entries()].map(([y, ms]) => {
              const keys = [...ms.keys()].map((m) => `${y}-${m}`);
              const n = [...ms.values()].reduce((a, b) => a + b, 0);
              return (
                <div key={y}>
                  <label>
                    <input type="checkbox" checked={keys.every((k) => months.has(k))} onChange={() => toggleMonths(keys)} />
                    <span className="grow">{y} 年</span>
                    <span className="sub mono">{n}</span>
                  </label>
                  {[...ms.entries()]
                    .sort((a, b) => a[0] - b[0])
                    .map(([m, cnt]) => (
                      <label key={m} style={{ paddingLeft: 18 }}>
                        <input type="checkbox" checked={months.has(`${y}-${m}`)} onChange={() => toggleMonths([`${y}-${m}`])} />
                        <span className="grow">{m} 月</span>
                        <span className="sub mono">{cnt}</span>
                      </label>
                    ))}
                </div>
              );
            })}
            {!!p.values.length && (
            <label>
              <input type="checkbox" checked={allOn} onChange={toggleAll} />
              （全选）
            </label>
            )}
            {shown.map((v) => (
              <label key={v.text}>
                <input
                  type="checkbox"
                  checked={checked.has(v.text)}
                  onChange={() => {
                    const next = new Set(checked);
                    if (next.has(v.text)) next.delete(v.text);
                    else next.add(v.text);
                    setChecked(next);
                  }}
                />
                <span className="ellipsis grow">{v.text === '' ? '（空白）' : v.text}</span>
                <span className="sub mono">{v.count}</span>
              </label>
            ))}
          </div>
        </>
      ) : mode === 'color' ? (
        <div className="sheet-filter-list">
          {(
            [
              ['按单元格颜色', p.fills, true],
              ['按字体颜色', p.colors, false],
            ] as const
          ).map(([label, map, cell]) =>
            map.size > 1 ? (
              <div key={label}>
                <div className="sub" style={{ padding: '4px 6px' }}>
                  {label}
                </div>
                {[...map.entries()].map(([color, n]) => {
                  const on = p.current?.color?.cell === cell && (p.current.color.color ?? null) === color;
                  return (
                    <button
                      key={color ?? 'none'}
                      role="menuitem"
                      className={on ? 'on' : undefined}
                      onClick={() => p.onApply({ color: { cell, color } })}
                    >
                      <span className="sheet-filter-swatch" style={{ background: color ?? 'transparent', ...(cell ? {} : { background: 'transparent', color: color ?? 'var(--text)' }) }}>
                        {cell ? '' : 'A'}
                      </span>
                      <span className="grow">{color ? color.toUpperCase() : cell ? '无填充' : '自动'}</span>
                      <span className="sub mono">{n}</span>
                    </button>
                  );
                })}
              </div>
            ) : null,
          )}
        </div>
      ) : (
        <div className="sheet-filter-custom">
          <select className="ed-select" value={op} onChange={(e) => setOp(e.target.value as Op)}>
            {OPS.map((o) => (
              <option key={o.op} value={o.op}>
                {o.label}
              </option>
            ))}
          </select>
          <input value={val} autoFocus placeholder="值（可用 * ? 通配）" onChange={(e) => setVal(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), apply())} />
        </div>
      )}
      <div className="sheet-filter-foot">
        <button className="btn" onClick={() => p.onApply(null)} disabled={!p.current}>
          清除筛选
        </button>
        <span className="grow" />
        <button className="btn" onClick={p.onClose}>
          取消
        </button>
        <button className="btn primary" onClick={apply}>
          确定
        </button>
      </div>
    </div>
  );
}

/**
 * 数据 → 转为智能列：按每列现有内容推断字段类型，列出预览表，逐列确认 / 修改后一次应用（一步撤销）。
 * 单元格内容不变，只给列加字段类型（存进工作包），已有数据按类型设数字格式。
 */
import { useEffect, useMemo, useState } from 'react';
import { COLUMN_TYPE_GLYPHS, COLUMN_TYPE_LABELS, host, type ColumnType } from '@wz/editor-shared';
import { colName } from '@wz/editor-shared/xlsx';
import { columnValues, inferColumns, type Inference } from '../core/colmeta';
import { typeFormat } from '../core/columns';
import type { SheetData, Workbook } from '../core/workbook';
import './columns.css';

const TYPES: ColumnType[] = ['text', 'number', 'money', 'percent', 'date', 'select', 'multi', 'check', 'link', 'email', 'phone'];

function describe(x: Inference): { opts: string[]; detail: string } {
  if (x.type === 'select' || x.type === 'multi') return { opts: (x.options ?? []).slice(0, 8), detail: (x.options?.length ?? 0) > 8 ? `…共 ${x.options!.length} 项` : '' };
  if (x.type === 'text') return { opts: [], detail: '不设置' };
  if (x.type === 'number') return { opts: [], detail: x.integer ? '只允许整数' : '不限小数位' };
  if (x.type === 'check') return { opts: [], detail: '是 = 勾选，否 = 空框' };
  const fmt = typeFormat({ type: x.type, fmt: x.fmt });
  return { opts: [], detail: fmt ?? '' };
}

interface Props {
  book: Workbook;
  sheet: SheetData;
  /** 表头所在行（1 起）的初值 */
  header: number;
  onApply: (header: number, list: Inference[]) => void;
  onClose: () => void;
}

export function SmartColumnsDialog({ book, sheet, header: initHeader, onApply, onClose }: Props) {
  const [header, setHeader] = useState(initHeader);
  const [seed, setSeed] = useState(0);
  const inferred = useMemo(() => inferColumns(book, sheet, header), [book, sheet, header, seed]); // eslint-disable-line react-hooks/exhaustive-deps
  /** 用户改过的类型（列号 → 类型）和不应用的列 */
  const [edited, setEdited] = useState<Record<number, ColumnType>>({});
  const [off, setOff] = useState<Record<number, boolean>>({});

  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    document.addEventListener('keydown', onKey, true);
    return () => document.removeEventListener('keydown', onKey, true);
  }, [onClose]);

  /** 用户把类型改成别的：选项 / 格式按新类型重新取 */
  const resolved = useMemo<Inference[]>(
    () =>
      inferred.map((x) => {
        const t = edited[x.col];
        if (!t || t === x.type) return x;
        const out: Inference = { ...x, type: t, reason: `${x.reason}（已改为${COLUMN_TYPE_LABELS[t]}）`, detail: '' };
        delete out.options;
        delete out.integer;
        delete out.fmt;
        delete out.warn;
        if (t === 'select' || t === 'multi') out.options = columnValues(sheet, header, x.col, t === 'multi', 12);
        if (t === 'money') out.fmt = '#,##0.00';
        if (t === 'percent') out.fmt = '0.00%';
        if (t === 'date') out.fmt = 'yyyy-mm-dd';
        if (t === 'number') out.integer = false;
        return out;
      }),
    [inferred, edited, sheet, header],
  );

  const picked = resolved.filter((x) => x.type !== 'text' && !off[x.col]);
  const maxHeader = Math.max(1, Math.min(10, sheet.rows));

  return (
    <div className="sg-mask" onMouseDown={onClose}>
      <div className="sg" role="dialog" aria-label="转为智能列" onMouseDown={(e) => e.stopPropagation()}>
        <div className="sg-head">
          <div>
            <div className="sg-title">转为智能列</div>
            <div className="sg-sub">按每列现有内容推断类型，逐列确认后一次应用。单元格内容不变，只给列加字段类型（保存在工作包里），可撤销。</div>
          </div>
          <button className="cp-x" aria-label="关闭" onClick={onClose}>
            ×
          </button>
        </div>
        <div className="sg-bar">
          <span>
            数据区 {colName(0)}
            {header}:{colName(Math.max(0, sheet.cols - 1))}
            {sheet.rows}
          </span>
          <label>
            表头在第
            <select className="ed-select" value={header} onChange={(e) => (setHeader(Number(e.target.value)), setEdited({}), setOff({}))}>
              {Array.from({ length: maxHeader }, (_, i) => (
                <option key={i} value={i + 1}>
                  {i + 1}
                </option>
              ))}
            </select>
            行
          </label>
          <span>共 {Math.max(0, sheet.rows - header)} 行数据</span>
          <span className="grow" />
          <button className="cp-link" onClick={() => (setEdited({}), setOff({}), setSeed((n) => n + 1))}>
            重新推断
          </button>
        </div>
        <div className="sg-table" role="table">
          <div className="sg-row sg-th" role="row">
            <span>应用</span>
            <span>列</span>
            <span>表头</span>
            <span>类型</span>
            <span>选项 / 格式</span>
            <span>推断依据</span>
          </div>
          {resolved.length === 0 && <div className="sg-empty">表头行里没有文字，请换一个表头行。</div>}
          {resolved.map((x) => {
            const on = x.type !== 'text' && !off[x.col];
            const d = describe(x);
            return (
              <div key={x.col} className={`sg-row${on ? '' : ' off'}`} role="row">
                <span>
                  <input type="checkbox" aria-label="应用此列" checked={on} disabled={x.type === 'text'} onChange={(e) => setOff({ ...off, [x.col]: !e.target.checked })} />
                </span>
                <span className="sub">{colName(x.col)}</span>
                <span className="sg-head-cell">{x.head}</span>
                <span className="sg-type">
                  <span className="cp-glyph">{COLUMN_TYPE_GLYPHS[x.type]}</span>
                  <select
                    aria-label={`${x.head} 的类型`}
                    className={`ed-select${edited[x.col] && edited[x.col] !== inferred.find((i) => i.col === x.col)?.type ? ' edited' : ''}`}
                    value={x.type}
                    onChange={(e) => setEdited({ ...edited, [x.col]: e.target.value as ColumnType })}
                  >
                    {TYPES.map((t) => (
                      <option key={t} value={t}>
                        {COLUMN_TYPE_LABELS[t]}
                      </option>
                    ))}
                  </select>
                </span>
                <span className="sg-opts">
                  {d.opts.map((o) => (
                    <span key={o} className="sg-tag">
                      {o}
                    </span>
                  ))}
                  {d.detail && <span className="sub">{d.detail}</span>}
                </span>
                <span className="sg-why">
                  <span>{x.reason}</span>
                  {x.warn && <span className="sg-warn">{x.warn}</span>}
                </span>
              </div>
            );
          })}
        </div>
        <div className="sg-foot">
          <span className="sub">
            将设置 <b>{picked.length}</b> 列，文本列不做改动
          </span>
          <span className="grow" />
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" disabled={picked.length === 0} onClick={() => onApply(header, picked)}>
            应用 {picked.length} 列
          </button>
        </div>
      </div>
    </div>
  );
}

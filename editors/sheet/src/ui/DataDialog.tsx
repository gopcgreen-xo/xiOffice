/** 数据工具对话框：删除重复项（选比较的列、是否有标题行）、分列（分隔符）。 */
import { useEffect, useMemo, useState } from 'react';
import { host } from '@wz/editor-shared';
import { colName } from '@wz/editor-shared/xlsx';
import { splitText, type SplitOptions } from '../core/data';
import type { Range, SheetData } from '../core/workbook';
import { Icon } from '@wz/editor-shared/Icon';

interface DupProps {
  sheet: SheetData;
  range: Range;
  header: boolean;
  onOk: (cols: number[], header: boolean) => void;
  onClose: () => void;
}

function usePopup() {
  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
}

function Frame(p: { title: string; range: Range; onClose: () => void; children: React.ReactNode }) {
  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-sort" role="dialog" aria-label={p.title} onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
        <div className="sheet-detail-head">
          <strong>{p.title}</strong>
          <span className="sub mono">
            {colName(p.range.c0)}
            {p.range.r0 + 1}:{colName(p.range.c1)}
            {p.range.r1 + 1}
          </span>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        {p.children}
      </div>
    </div>
  );
}

export function DuplicatesDialog(p: DupProps) {
  usePopup();
  const cols = Array.from({ length: p.range.c1 - p.range.c0 + 1 }, (_, i) => p.range.c0 + i);
  const [header, setHeader] = useState(p.header);
  const [picked, setPicked] = useState<Set<number>>(() => new Set(cols));
  const label = (c: number) => {
    const h = p.sheet.cells[p.range.r0]?.[c]?.t;
    return header && h ? `${h}（${colName(c)} 列）` : `${colName(c)} 列`;
  };
  const toggle = (c: number) => setPicked((s) => (s.has(c) ? new Set([...s].filter((x) => x !== c)) : new Set([...s, c])));
  return (
    <Frame title="删除重复项" range={p.range} onClose={p.onClose}>
      <div className="sub" style={{ padding: '4px 0' }}>
        按勾选的列比较，这些列都相同的行只保留第一行：
      </div>
      <div className="sheet-data-cols">
        {cols.map((c) => (
          <label key={c}>
            <input type="checkbox" checked={picked.has(c)} onChange={() => toggle(c)} />
            {label(c)}
          </label>
        ))}
      </div>
      <div className="sheet-detail-foot">
        <button className="btn" onClick={() => setPicked(new Set(picked.size === cols.length ? [] : cols))}>
          {picked.size === cols.length ? '全不选' : '全选'}
        </button>
        <label className="sheet-sort-header">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
          数据包含标题
        </label>
        <span className="grow" />
        <button className="btn" onClick={p.onClose}>
          取消
        </button>
        <button className="btn primary" disabled={!picked.size} onClick={() => p.onOk(cols.filter((c) => picked.has(c)), header)}>
          确定
        </button>
      </div>
    </Frame>
  );
}

const DELIMS: [string, string][] = [
  ['Tab 键', '\t'],
  ['逗号', ','],
  ['分号', ';'],
  ['空格', ' '],
];

interface SplitProps {
  sheet: SheetData;
  range: Range;
  onOk: (o: SplitOptions) => void;
  onClose: () => void;
}

export function SplitDialog(p: SplitProps) {
  usePopup();
  const [on, setOn] = useState<Set<string>>(() => new Set(['\t', ',']));
  const [other, setOther] = useState('');
  const [merge, setMerge] = useState(false);
  const opts: SplitOptions = useMemo(() => ({ delims: [...on, ...(other ? [...other] : [])], merge }), [on, other, merge]);
  // 预览前 5 行
  const preview = useMemo(() => {
    const out: string[][] = [];
    for (let r = p.range.r0; r <= p.range.r1 && out.length < 5; r++) {
      const v = p.sheet.cells[r]?.[p.range.c0]?.v;
      if (typeof v === 'string') out.push(splitText(v, opts));
    }
    return out;
  }, [opts]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <Frame title="分列" range={p.range} onClose={p.onClose}>
      <div className="sheet-data-cols">
        {DELIMS.map(([label, ch]) => (
          <label key={label}>
            <input type="checkbox" checked={on.has(ch)} onChange={() => setOn((s) => (s.has(ch) ? new Set([...s].filter((x) => x !== ch)) : new Set([...s, ch])))} />
            {label}
          </label>
        ))}
        <label>
          其他
          <input className="input" style={{ width: 48 }} maxLength={4} value={other} onChange={(e) => setOther(e.target.value)} placeholder="如 / 、" />
        </label>
        <label>
          <input type="checkbox" checked={merge} onChange={(e) => setMerge(e.target.checked)} />
          连续分隔符视为单个处理
        </label>
      </div>
      <table className="sheet-data-preview">
        <tbody>
          {preview.map((row, i) => (
            <tr key={i}>
              {row.map((v, j) => (
                <td key={j}>{v}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      <div className="sheet-detail-foot">
        <span className="sub">拆出的内容会覆盖右边的列</span>
        <span className="grow" />
        <button className="btn" onClick={p.onClose}>
          取消
        </button>
        <button className="btn primary" disabled={!opts.delims.length} onClick={() => p.onOk(opts)}>
          确定
        </button>
      </div>
    </Frame>
  );
}

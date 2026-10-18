/** 自定义排序：多个排序条件（列 + 升序 / 降序），数据是否包含标题行。 */
import { useEffect, useState } from 'react';
import { host } from '@wz/editor-shared';
import { colName } from '@wz/editor-shared/xlsx';
import type { SortKey } from '../core/sortFilter';
import type { Range, SheetData } from '../core/workbook';
import { Icon } from '@wz/editor-shared/Icon';

interface Props {
  sheet: SheetData;
  range: Range;
  header: boolean;
  onOk: (keys: SortKey[], header: boolean) => void;
  onClose: () => void;
}

export function SortDialog(p: Props) {
  const [header, setHeader] = useState(p.header);
  const [keys, setKeys] = useState<SortKey[]>([{ c: p.range.c0, desc: false }]);

  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);

  const label = (c: number) => {
    const h = p.sheet.cells[p.range.r0]?.[c]?.t;
    return header && h ? `${h}（${colName(c)} 列）` : `${colName(c)} 列`;
  };
  const cols = Array.from({ length: p.range.c1 - p.range.c0 + 1 }, (_, i) => p.range.c0 + i);

  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-sort" role="dialog" aria-label="排序" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
        <div className="sheet-detail-head">
          <strong>排序</strong>
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
        {keys.map((k, i) => (
          <div key={i} className="sheet-sort-row">
            <span className="sub">{i === 0 ? '主要关键字' : '次要关键字'}</span>
            <select className="ed-select" value={k.c} onChange={(e) => setKeys(keys.map((x, j) => (j === i ? { ...x, c: Number(e.target.value) } : x)))}>
              {cols.map((c) => (
                <option key={c} value={c}>
                  {label(c)}
                </option>
              ))}
            </select>
            <select className="ed-select" value={k.desc ? 'desc' : 'asc'} onChange={(e) => setKeys(keys.map((x, j) => (j === i ? { ...x, desc: e.target.value === 'desc' } : x)))}>
              <option value="asc">升序</option>
              <option value="desc">降序</option>
            </select>
            {i > 0 && (
              <button className="icon-btn" aria-label="删除条件" onClick={() => setKeys(keys.filter((_, j) => j !== i))}>
                <Icon name="close" size={12} />
              </button>
            )}
          </div>
        ))}
        <div className="sheet-detail-foot">
          <button className="btn" disabled={keys.length >= cols.length || keys.length >= 6} onClick={() => setKeys([...keys, { c: cols.find((c) => !keys.some((k) => k.c === c)) ?? p.range.c0, desc: false }])}>
            添加条件
          </button>
          <label className="sheet-sort-header">
            <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
            数据包含标题
          </label>
          <span className="grow" />
          <button className="btn" onClick={p.onClose}>
            取消
          </button>
          <button className="btn primary" onClick={() => p.onOk(keys, header)}>
            确定
          </button>
        </div>
      </div>
    </div>
  );
}

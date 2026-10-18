/** 数据透视表对话框：给每个字段选“不用 / 行 / 列 / 值”（值再选汇总方式），选要不要总计。 */
import { useEffect, useState } from 'react';
import { host } from '@wz/editor-shared';
import { Icon } from '@wz/editor-shared/Icon';
import { colName } from '@wz/editor-shared/xlsx';
import { AGG_LABEL, type Agg, type PivotConfig } from '../core/pivot';
import type { Range } from '../core/workbook';

type Role = 'none' | 'row' | 'col' | 'value';

const AGGS = Object.keys(AGG_LABEL) as Agg[];

export function PivotDialog(p: { range: Range; fields: string[]; initial: PivotConfig; onOk: (cfg: PivotConfig) => void; onClose: () => void }) {
  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  const roleOf = (f: number): Role => (p.initial.rows.includes(f) ? 'row' : p.initial.cols.includes(f) ? 'col' : p.initial.values.some((v) => v.field === f) ? 'value' : 'none');
  const [roles, setRoles] = useState<Role[]>(() => p.fields.map((_, f) => roleOf(f)));
  const [aggs, setAggs] = useState<Agg[]>(() => p.fields.map((_, f) => p.initial.values.find((v) => v.field === f)?.agg ?? 'sum'));
  const [totRows, setTotRows] = useState(p.initial.totals.rows);
  const [totCols, setTotCols] = useState(p.initial.totals.cols);

  const pick = (role: Role) => p.fields.map((_, f) => f).filter((f) => roles[f] === role);
  const cfg = (): PivotConfig => ({
    rows: pick('row'),
    cols: pick('col'),
    values: pick('value').map((f) => ({ field: f, agg: aggs[f] })),
    totals: { rows: totRows, cols: totCols },
  });
  const ok = () => {
    const c = cfg();
    if (!c.rows.length && !c.cols.length && !c.values.length) return host().toast('请至少选一个字段');
    if (c.cols.length > 1) return host().toast('列字段最多选一个');
    p.onOk(c);
  };
  const set = <T,>(arr: T[], i: number, v: T) => arr.map((x, j) => (i === j ? v : x));

  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-sort" role="dialog" aria-label="数据透视表" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
        <div className="sheet-detail-head">
          <strong>数据透视表</strong>
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
        <div className="sub" style={{ padding: '4px 0' }}>
          数据区域第一行是字段名。给每个字段选它的用途，结果放进新工作表，是 Excel 里能刷新、能拖拽字段的真透视表（列字段最多一个；改了源数据后可在“排序和筛选”菜单里刷新）：
        </div>
        <div className="pivot-fields">
          {p.fields.map((name, f) => (
            <div key={f} className="pivot-field">
              <span className="ellipsis grow" title={name}>
                {name}
              </span>
              <select aria-label={`${name} 的用途`} className="ed-select" value={roles[f]} onChange={(e) => setRoles(set(roles, f, e.target.value as Role))}>
                <option value="none">不用</option>
                <option value="row">行</option>
                <option value="col">列</option>
                <option value="value">值</option>
              </select>
              {roles[f] === 'value' && (
                <select aria-label={`${name} 的汇总方式`} className="ed-select" value={aggs[f]} onChange={(e) => setAggs(set(aggs, f, e.target.value as Agg))}>
                  {AGGS.map((a) => (
                    <option key={a} value={a}>
                      {AGG_LABEL[a]}
                    </option>
                  ))}
                </select>
              )}
            </div>
          ))}
        </div>
        <div className="sheet-data-cols">
          <label>
            <input type="checkbox" checked={totRows} onChange={(e) => setTotRows(e.target.checked)} />
            显示总计行
          </label>
          <label>
            <input type="checkbox" checked={totCols} onChange={(e) => setTotCols(e.target.checked)} />
            显示总计列
          </label>
        </div>
        <div className="sheet-detail-foot">
          <span className="grow" />
          <button className="btn" onClick={p.onClose}>
            取消
          </button>
          <button className="btn primary" onClick={ok}>
            确定
          </button>
        </div>
      </div>
    </div>
  );
}

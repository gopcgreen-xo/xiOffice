/** 图表对话框：插入 / 编辑图表（类型、标题、数据区域及系列方向），带实时预览。 */
import { useEffect, useMemo, useState } from 'react';
import { host } from '@wz/editor-shared';
import { Icon } from '@wz/editor-shared/Icon';
import { colName } from '@wz/editor-shared/xlsx';
import { CHART_TYPES, guessSource, seriesFromRange, type ChartSource } from '../core/charts';
import type { ChartSpec } from '../core/drawing';
import { parseRef, type Range, type Workbook } from '../core/workbook';
import { SheetChart } from './SheetChart';

interface Props {
  book: Workbook;
  si: number;
  /** 新建时的数据区域（选区）；编辑时没有 */
  range?: Range;
  /** 编辑时的现有图表 */
  initial?: ChartSpec;
  onOk: (spec: ChartSpec) => void;
  onClose: () => void;
}

const rangeText = (g: Range) => `${colName(g.c0)}${g.r0 + 1}${g.r1 === g.r0 && g.c1 === g.c0 ? '' : `:${colName(g.c1)}${g.r1 + 1}`}`;

export function ChartDialog(p: Props) {
  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  const sheet = p.book.sheets[p.si];
  const editing = !!p.initial;
  const [type, setType] = useState<ChartSpec['type']>(p.initial?.type === 'other' ? 'col' : (p.initial?.type ?? 'col'));
  const [title, setTitle] = useState(p.initial?.title ?? '');
  // 编辑已有图表时，数据区域留空 = 保持现有系列不变
  const [text, setText] = useState(p.range ? rangeText(p.range) : '');
  const [src, setSrc] = useState<ChartSource>(() => (p.range ? guessSource(sheet, p.range) : { header: true, labels: true, byRows: false }));

  const g = useMemo(() => (text.trim() ? parseRef(text.trim().replace(/^.*!/, '').replace(/\$/g, '')) : null), [text]);
  const series = useMemo(() => (g ? seriesFromRange(p.book, p.si, g, src) : (p.initial?.series ?? [])), [g, src, p.book, p.si, p.initial]);
  const spec: ChartSpec = { type, ...(title.trim() && { title: title.trim() }), series };
  const bad = text.trim() && !g ? '数据区域不对，请写成 A1:C10 的样子' : !series.length ? '这片数据画不出图表：至少要有一列数字' : null;

  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-chart-dlg" role="dialog" aria-label={editing ? '编辑图表' : '插入图表'} onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
        <div className="sheet-detail-head">
          <strong>{editing ? '编辑图表' : '插入图表'}</strong>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        <div className="chart-types" role="radiogroup" aria-label="图表类型">
          {CHART_TYPES.map((t) => (
            <button key={t.type} role="radio" aria-checked={type === t.type} className={`btn${type === t.type ? ' primary' : ''}`} onClick={() => setType(t.type)}>
              {t.label}
            </button>
          ))}
        </div>
        <div className="names-form">
          <label>
            <span>标题</span>
            <input aria-label="图表标题" className="input" value={title} onChange={(e) => setTitle(e.target.value)} placeholder="（可不填）" />
          </label>
          <label>
            <span>数据区域</span>
            <input aria-label="数据区域" className="input mono" value={text} onChange={(e) => setText(e.target.value)} placeholder={editing ? '留空 = 保持现有数据' : 'A1:C10'} />
          </label>
        </div>
        {(!editing || g) && (
          <div className="sheet-data-cols">
            <label>
              <input type="checkbox" checked={src.header} onChange={(e) => setSrc({ ...src, header: e.target.checked })} />
              {src.byRows ? '第一列是系列名称' : '第一行是系列名称'}
            </label>
            <label>
              <input type="checkbox" checked={src.labels} onChange={(e) => setSrc({ ...src, labels: e.target.checked })} />
              {src.byRows ? '第一行是分类标签' : '第一列是分类标签'}
            </label>
            <label>
              <input type="checkbox" checked={src.byRows} onChange={(e) => setSrc({ ...src, byRows: e.target.checked })} />
              按行取系列
            </label>
          </div>
        )}
        {bad && <div className="ed-warn">{bad}</div>}
        <div className="chart-preview">{!bad && <SheetChart spec={spec} w={440} h={250} book={p.book} />}</div>
        <div className="sheet-detail-foot">
          <span className="sub">{series.length ? `${series.length} 个系列` : ''}</span>
          <span className="grow" />
          <button className="btn" onClick={p.onClose}>
            取消
          </button>
          <button className="btn primary" disabled={!!bad} onClick={() => p.onOk(spec)}>
            确定
          </button>
        </div>
      </div>
    </div>
  );
}

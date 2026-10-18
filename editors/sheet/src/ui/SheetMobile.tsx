/**
 * 表格编辑器的手机底部工具条（docs/ANDROID.md §6.3）。
 * - 查看模式：当前单元格的地址和内容（公式显示原文）、选区统计、查找。
 * - 编辑模式：底部编辑栏（和桌面 fx 栏同一套编辑逻辑，输入公式时点格子插入引用）、撤销 / 重做、选区模式、
 *   格式（加粗、颜色、填充、对齐、数字格式、边框）、行列（插入 / 删除）、数据（排序、筛选）、复制 / 粘贴 / 清除。
 * 按钮都不抢输入框焦点（onPointerDown 取消默认），编辑栏打开的键盘不会收起。
 */
import { useState, type Ref, type ReactNode } from 'react';
import { Icon, type IconName } from '@wz/editor-shared/Icon';
import type { StyleChange } from '../core/styles';
import type { BorderKind } from '../core/format';

type RowsOp = 'rowAbove' | 'rowBelow' | 'colLeft' | 'colRight' | 'delRow' | 'delCol';
type Panel = 'format' | 'rows' | 'data' | null;

export interface SheetMobileProps {
  edit: boolean;
  onEdit(on: boolean): void;
  range: boolean;
  onRange(on: boolean): void;
  rangeText: string;
  barProps: Record<string, unknown> & { ref: Ref<HTMLTextAreaElement>; value: string };
  editing: boolean;
  onCommit(): void;
  onCancel(): void;
  canUndo: boolean;
  canRedo: boolean;
  onUndo(): void;
  onRedo(): void;
  st: { bold?: boolean; italic?: boolean; underline?: boolean; color?: string; fill?: string; align?: string };
  fmt: string;
  onStyle(ch: StyleChange, label: string): void;
  onBorder(kind: BorderKind): void;
  onNumFmt(code: string): void;
  onRows(op: RowsOp): void;
  onSort(desc: boolean): void;
  filtered: boolean;
  onFilter(): void;
  onCopy(): void;
  onPaste(): void;
  onClear(): void;
  address: string;
  content: string;
  stats: { count: number; sum: string | null } | null;
  query: string;
  hits: number;
  hitIdx: number;
  onQuery(q: string): void;
  onHit(dir: 1 | -1): void;
}

/** 不抢焦点的按钮（键盘保持打开） */
const keep = { onPointerDown: (e: React.PointerEvent) => e.preventDefault(), onMouseDown: (e: React.MouseEvent) => e.preventDefault() };

function Tool({ icon, glyph, label, on, disabled, onClick }: { icon?: IconName; glyph?: string; label: string; on?: boolean; disabled?: boolean; onClick(): void }) {
  return (
    <button className={`m-tool${on ? ' on' : ''}`} disabled={disabled} onClick={onClick} {...keep}>
      {icon ? <Icon name={icon} size={20} /> : <span className="m-tool-glyph">{glyph}</span>}
      <span>{label}</span>
    </button>
  );
}

function Chip({ children, on, onClick, title, style }: { children: ReactNode; on?: boolean; onClick(): void; title?: string; style?: React.CSSProperties }) {
  return (
    <button className={`sheet-m-chip${on ? ' on' : ''}`} title={title} aria-label={title} style={style} onClick={onClick} {...keep}>
      {children}
    </button>
  );
}

const COLORS = ['#000000', '#c00000', '#0070c0', '#00b050', '#ed7d31', '#7030a0'];
const FILLS = ['#ffff00', '#c6efce', '#ffc7ce', '#ddebf7', '#fce4d6', '#d9d9d9'];
const NUM_FMTS: [string, string][] = [
  ['General', '常规'],
  ['0.00', '数字'],
  ['¥#,##0.00', '金额'],
  ['0%', '百分比'],
  ['yyyy-mm-dd', '日期'],
  ['@', '文本'],
];

export function SheetMobileBar(p: SheetMobileProps) {
  const [finding, setFinding] = useState(false);
  const [panel, setPanel] = useState<Panel>(null);
  const toggle = (x: Panel) => setPanel(panel === x ? null : x);

  const findBar = (
    <form
      className="sheet-m-find"
      onSubmit={(e) => {
        e.preventDefault();
        p.onHit(1);
      }}
    >
      <input autoFocus value={p.query} placeholder="查找" enterKeyHint="search" onChange={(e) => p.onQuery(e.target.value)} />
      <span className="sheet-m-count mono">{p.query ? `${p.hitIdx + 1 || '–'}/${p.hits}` : ''}</span>
      <button type="button" aria-label="上一个" disabled={!p.hits} onClick={() => p.onHit(-1)}>
        <Icon name="arrowUp" size={18} />
      </button>
      <button type="button" aria-label="下一个" disabled={!p.hits} onClick={() => p.onHit(1)}>
        <Icon name="arrowDown" size={18} />
      </button>
      <button
        type="button"
        aria-label="关闭查找"
        onClick={() => {
          p.onQuery('');
          setFinding(false);
        }}
      >
        <Icon name="close" size={18} />
      </button>
    </form>
  );

  if (!p.edit) {
    return (
      <div className="sheet-m">
        {finding ? (
          findBar
        ) : (
          <div className="sheet-m-cell" role="status">
            <span className="sheet-m-addr mono">{p.address}</span>
            <span className="sheet-m-text">{p.content || <span className="sub">（空）</span>}</span>
            {p.stats && p.stats.count > 1 && (
              <span className="sheet-m-stats mono">
                计数 {p.stats.count}
                {p.stats.sum !== null && ` · 求和 ${p.stats.sum}`}
              </span>
            )}
          </div>
        )}
        <div className="m-toolbar" role="toolbar" aria-label="表格工具">
          <Tool icon="edit" label="编辑" onClick={() => (setFinding(false), p.onEdit(true))} />
          <Tool icon="search" label="查找" on={finding} onClick={() => setFinding(!finding)} />
          <Tool icon="copy" label="复制" onClick={p.onCopy} />
        </div>
      </div>
    );
  }

  return (
    <div className="sheet-m">
      {panel === 'format' && (
        <div className="sheet-m-panel" role="group" aria-label="格式">
          <div className="sheet-m-row">
            <Chip on={!!p.st.bold} title="加粗" onClick={() => p.onStyle({ bold: !p.st.bold }, '加粗')}>
              <b>B</b>
            </Chip>
            <Chip on={!!p.st.italic} title="倾斜" onClick={() => p.onStyle({ italic: !p.st.italic }, '倾斜')}>
              <i>I</i>
            </Chip>
            <Chip on={!!p.st.underline} title="下划线" onClick={() => p.onStyle({ underline: !p.st.underline }, '下划线')}>
              <u>U</u>
            </Chip>
            <span className="sheet-m-sep" />
            {(['left', 'center', 'right'] as const).map((a) => (
              <Chip key={a} on={p.st.align === a} title={a === 'left' ? '左对齐' : a === 'center' ? '居中' : '右对齐'} onClick={() => p.onStyle({ align: p.st.align === a ? null : a }, '对齐')}>
                {a === 'left' ? '左' : a === 'center' ? '中' : '右'}
              </Chip>
            ))}
          </div>
          <div className="sheet-m-row">
            <span className="sheet-m-label">文字</span>
            {COLORS.map((c) => (
              <Chip key={c} title={`文字颜色 ${c}`} on={p.st.color === c} onClick={() => p.onStyle({ color: c }, '文字颜色')} style={{ color: c }}>
                A
              </Chip>
            ))}
          </div>
          <div className="sheet-m-row">
            <span className="sheet-m-label">填充</span>
            <Chip title="无填充" onClick={() => p.onStyle({ fill: null }, '填充')}>
              无
            </Chip>
            {FILLS.map((c) => (
              <Chip key={c} title={`填充 ${c}`} on={p.st.fill === c} onClick={() => p.onStyle({ fill: c }, '填充')} style={{ background: c }}>
                {' '}
              </Chip>
            ))}
          </div>
          <div className="sheet-m-row">
            <span className="sheet-m-label">数字</span>
            {NUM_FMTS.map(([code, label]) => (
              <Chip key={code} on={p.fmt === code} onClick={() => p.onNumFmt(code)}>
                {label}
              </Chip>
            ))}
          </div>
          <div className="sheet-m-row">
            <span className="sheet-m-label">边框</span>
            <Chip onClick={() => p.onBorder('all')}>全部</Chip>
            <Chip onClick={() => p.onBorder('outside')}>外框</Chip>
            <Chip onClick={() => p.onBorder('none')}>无</Chip>
          </div>
        </div>
      )}
      {panel === 'rows' && (
        <div className="sheet-m-panel" role="group" aria-label="行和列">
          <div className="sheet-m-row">
            <Chip onClick={() => p.onRows('rowAbove')}>上方插入行</Chip>
            <Chip onClick={() => p.onRows('rowBelow')}>下方插入行</Chip>
            <Chip onClick={() => p.onRows('delRow')}>删除行</Chip>
          </div>
          <div className="sheet-m-row">
            <Chip onClick={() => p.onRows('colLeft')}>左侧插入列</Chip>
            <Chip onClick={() => p.onRows('colRight')}>右侧插入列</Chip>
            <Chip onClick={() => p.onRows('delCol')}>删除列</Chip>
          </div>
        </div>
      )}
      {panel === 'data' && (
        <div className="sheet-m-panel" role="group" aria-label="数据">
          <div className="sheet-m-row">
            <Chip onClick={() => p.onSort(false)}>升序</Chip>
            <Chip onClick={() => p.onSort(true)}>降序</Chip>
            <Chip on={p.filtered} onClick={p.onFilter}>
              {p.filtered ? '取消筛选' : '筛选'}
            </Chip>
            <Chip onClick={p.onClear}>清除内容</Chip>
          </div>
        </div>
      )}
      {finding ? (
        findBar
      ) : (
        <div className="sheet-m-edit">
          <span className="sheet-m-addr mono">{p.range ? p.rangeText : p.address}</span>
          <textarea className="sheet-m-input" aria-label="编辑栏" rows={1} spellCheck={false} enterKeyHint="done" {...p.barProps} />
          {p.editing && (
            <>
              <button className="sheet-m-ok" aria-label="确定" onClick={p.onCommit} {...keep}>
                <Icon name="check" size={20} />
              </button>
              <button className="sheet-m-x" aria-label="取消" onClick={p.onCancel} {...keep}>
                <Icon name="close" size={20} />
              </button>
            </>
          )}
        </div>
      )}
      <div className="m-toolbar" role="toolbar" aria-label="表格编辑">
        <Tool icon="check" label="完成" onClick={() => (setPanel(null), p.onEdit(false))} />
        <Tool icon="undo" label="撤销" disabled={!p.canUndo} onClick={p.onUndo} />
        <Tool icon="redo" label="重做" disabled={!p.canRedo} onClick={p.onRedo} />
        <Tool icon="selectAll" label="选区" on={p.range} onClick={() => p.onRange(!p.range)} />
        <Tool glyph="A" label="格式" on={panel === 'format'} onClick={() => toggle('format')} />
        <Tool icon="viewLines" label="行列" on={panel === 'rows'} onClick={() => toggle('rows')} />
        <Tool icon="sort" label="数据" on={panel === 'data'} onClick={() => toggle('data')} />
        <Tool icon="copy" label="复制" onClick={p.onCopy} />
        <Tool icon="paste" label="粘贴" onClick={p.onPaste} />
        <Tool icon="search" label="查找" on={finding} onClick={() => setFinding(!finding)} />
      </div>
    </div>
  );
}

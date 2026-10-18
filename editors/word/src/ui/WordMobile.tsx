/**
 * Word 编辑器的手机底部工具条（docs/ANDROID.md §6.2）。
 * - 查看（默认，不弹键盘）：编辑、目录、查找、适应宽度、页码。
 * - 编辑：完成、撤销 / 重做、选择（手指拖动选文字，不滚动）、文字格式、段落、表格（光标在表格里时）、图片、复制 / 剪切 / 粘贴。
 * 按钮都不抢文档输入框的焦点（onPointerDown 取消默认），键盘不会收起。
 * 审阅、套打按阶段加在这里。
 */
import { useState, type ReactNode } from 'react';
import { Icon, type IconName } from '@wz/editor-shared/Icon';
import type { ParaProps, RunProps } from '../core/model';

/** 手机上的工作模式：查看（默认，不弹键盘）/ 编辑 / 审阅 */
export type MobileMode = 'view' | 'edit' | 'review';
export type TableOp = 'rowAbove' | 'rowBelow' | 'colLeft' | 'colRight' | 'delRow' | 'delCol';

/** 审阅命令（WordEditor 里现成的审阅命令） */
export interface MobileReviewApi {
  trackOn: boolean;
  setTrack(on: boolean): void;
  prev(): void;
  next(): void;
  accept(): void;
  reject(): void;
  acceptAll(): void;
  rejectAll(): void;
  /** 批注条数 */
  comments: number;
  commentsOpen: boolean;
  toggleComments(): void;
  newComment(): void;
  /** “最终状态”：按全部接受后的样子显示（只看，不改文档） */
  final: boolean;
  setFinal(on: boolean): void;
}

/** 编辑命令（WordEditor 里现成的命令函数） */
export interface MobileEditApi {
  run: RunProps;
  para: ParaProps;
  /** 当前编号预设（'' = 没有编号） */
  preset: string;
  style: string;
  inTable: boolean;
  canUndo: boolean;
  canRedo: boolean;
  hasSelection: boolean;
  selecting: boolean;
  setSelecting(on: boolean): void;
  setRun(patch: Partial<RunProps>): void;
  setPara(patch: Partial<ParaProps>): void;
  grow(dir: 1 | -1): void;
  numbering(preset: string): void;
  paraStyle(id: string): void;
  table(op: TableOp): void;
  undo(): void;
  redo(): void;
  insertImage(): void;
  clip(kind: 'copy' | 'cut' | 'paste'): void;
}

type Panel = 'text' | 'para' | 'table' | null;

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
    <button className={`word-m-chip${on ? ' on' : ''}`} title={title} aria-label={title} aria-pressed={on} style={style} onClick={onClick} {...keep}>
      {children}
    </button>
  );
}

const COLORS = ['#000000', '#c00000', '#0070c0', '#00b050', '#ed7d31', '#7030a0'];
/** 突出显示（RRGGBB） */
const HIGHLIGHTS = ['FFFF00', '00FF00', '00FFFF', 'FF00FF'];
const STYLES: [string, string][] = [
  ['Normal', '正文'],
  ['Heading1', '标题 1'],
  ['Heading2', '标题 2'],
  ['Heading3', '标题 3'],
];

export function WordMobileBar(props: {
  mode: MobileMode;
  onMode(m: MobileMode): void;
  page: number;
  pages: number;
  fit: boolean;
  navOpen: boolean;
  onNav(): void;
  onFind(): void;
  onFit(): void;
  ed: MobileEditApi;
  rev: MobileReviewApi;
  /** 文档受保护（只读 / 只能批注）时不能进编辑 */
  readOnly?: string | null;
  /** 有套打入口时（手机版）：用当前文档当模板套打 */
  onMerge?: () => void;
}) {
  const { page, pages, fit, navOpen, ed, mode } = props;
  const [panel, setPanel] = useState<Panel>(null);
  const toggle = (p: Panel) => setPanel(panel === p ? null : p);
  const r = ed.run;

  const rv = props.rev;
  if (mode === 'review') {
    return (
      <>
        <div className="word-m-indicator" role="status">
          {page + 1} / {pages}
        </div>
        <div className="m-toolbar word-m-bar" role="toolbar" aria-label="审阅">
          <Tool icon="check" label="完成" onClick={() => props.onMode('view')} />
          <Tool icon="edit" label={rv.trackOn ? '修订:开' : '修订:关'} on={rv.trackOn} onClick={() => rv.setTrack(!rv.trackOn)} />
          <Tool icon="arrowUp" label="上一处" disabled={rv.final} onClick={rv.prev} />
          <Tool icon="arrowDown" label="下一处" disabled={rv.final} onClick={rv.next} />
          <Tool icon="check" label="接受" disabled={rv.final} onClick={rv.accept} />
          <Tool icon="close" label="拒绝" disabled={rv.final} onClick={rv.reject} />
          <Tool glyph="✓✓" label="全部接受" disabled={rv.final} onClick={rv.acceptAll} />
          <Tool glyph="✕✕" label="全部拒绝" disabled={rv.final} onClick={rv.rejectAll} />
          <Tool icon="bell" label={rv.comments ? `批注 ${rv.comments}` : '批注'} on={rv.commentsOpen} disabled={rv.final} onClick={rv.toggleComments} />
          <Tool icon="plus" label="新批注" disabled={rv.final || !ed.hasSelection} onClick={rv.newComment} />
          <Tool icon="info" label="最终状态" on={rv.final} onClick={() => rv.setFinal(!rv.final)} />
          <Tool icon="edit" label="编辑" disabled={rv.final || !!props.readOnly} onClick={() => props.onMode('edit')} />
        </div>
      </>
    );
  }

  if (mode !== 'edit') {
    return (
      <>
        <div className="word-m-indicator" role="status">
          {page + 1} / {pages}
        </div>
        <div className="m-toolbar word-m-bar" role="toolbar" aria-label="Word 工具">
          <Tool icon="edit" label="编辑" disabled={!!props.readOnly} onClick={() => props.onMode('edit')} />
          <Tool icon="contacts" label="审阅" onClick={() => props.onMode('review')} />
          <Tool icon="viewLines" label="目录" on={navOpen} onClick={props.onNav} />
          <Tool icon="search" label="查找" onClick={props.onFind} />
          <Tool icon="expand" label="适应宽度" on={fit} onClick={props.onFit} />
          {props.onMerge && <Tool icon="printer" label="套打" onClick={props.onMerge} />}
        </div>
      </>
    );
  }

  return (
    <>
      {panel === 'text' && (
        <div className="word-m-panel" role="group" aria-label="文字格式">
          <div className="word-m-row">
            <Chip title="加粗" on={r.bold} onClick={() => ed.setRun({ bold: !r.bold })}>
              <b>B</b>
            </Chip>
            <Chip title="倾斜" on={r.italic} onClick={() => ed.setRun({ italic: !r.italic })}>
              <i>I</i>
            </Chip>
            <Chip title="下划线" on={r.underline} onClick={() => ed.setRun({ underline: !r.underline })}>
              <u>U</u>
            </Chip>
            <Chip title="删除线" on={!!r.strike} onClick={() => ed.setRun({ strike: !r.strike })}>
              <s>S</s>
            </Chip>
            <span className="word-m-sep" />
            <Chip title="字号减小" onClick={() => ed.grow(-1)}>
              A-
            </Chip>
            <span className="word-m-size mono">{r.size}</span>
            <Chip title="字号增大" onClick={() => ed.grow(1)}>
              A+
            </Chip>
          </div>
          <div className="word-m-row">
            <span className="word-m-label">颜色</span>
            {COLORS.map((c) => (
              <Chip key={c} title={`字体颜色 ${c}`} on={r.color.toLowerCase() === c.slice(1)} onClick={() => ed.setRun({ color: c.slice(1) })} style={{ color: c }}>
                A
              </Chip>
            ))}
          </div>
          <div className="word-m-row">
            <span className="word-m-label">高亮</span>
            <Chip title="无高亮" on={!r.highlight} onClick={() => ed.setRun({ highlight: undefined })}>
              无
            </Chip>
            {HIGHLIGHTS.map((k) => (
              <Chip key={k} title={`高亮 #${k}`} on={r.highlight?.toUpperCase() === k} onClick={() => ed.setRun({ highlight: k })} style={{ background: `#${k}` }}>
                {' '}
              </Chip>
            ))}
          </div>
        </div>
      )}
      {panel === 'para' && (
        <div className="word-m-panel" role="group" aria-label="段落">
          <div className="word-m-row">
            {(['left', 'center', 'right', 'both'] as const).map((a) => (
              <Chip key={a} title={a === 'left' ? '左对齐' : a === 'center' ? '居中' : a === 'right' ? '右对齐' : '两端对齐'} on={ed.para.align === a} onClick={() => ed.setPara({ align: a })}>
                {a === 'left' ? '左' : a === 'center' ? '中' : a === 'right' ? '右' : '两端'}
              </Chip>
            ))}
          </div>
          <div className="word-m-row">
            <span className="word-m-label">列表</span>
            <Chip on={ed.preset === 'bullet'} onClick={() => ed.numbering(ed.preset === 'bullet' ? '' : 'bullet')}>
              ● 项目符号
            </Chip>
            <Chip on={ed.preset === 'decimal'} onClick={() => ed.numbering(ed.preset === 'decimal' ? '' : 'decimal')}>
              1. 编号
            </Chip>
            <Chip on={!ed.preset} onClick={() => ed.numbering('')}>
              无
            </Chip>
          </div>
          <div className="word-m-row">
            <span className="word-m-label">样式</span>
            {STYLES.map(([id, label]) => (
              <Chip key={id} on={ed.style === id} onClick={() => ed.paraStyle(id)}>
                {label}
              </Chip>
            ))}
          </div>
        </div>
      )}
      {panel === 'table' && ed.inTable && (
        <div className="word-m-panel" role="group" aria-label="表格">
          <div className="word-m-row">
            <Chip onClick={() => ed.table('rowAbove')}>上方插入行</Chip>
            <Chip onClick={() => ed.table('rowBelow')}>下方插入行</Chip>
            <Chip onClick={() => ed.table('delRow')}>删除行</Chip>
          </div>
          <div className="word-m-row">
            <Chip onClick={() => ed.table('colLeft')}>左侧插入列</Chip>
            <Chip onClick={() => ed.table('colRight')}>右侧插入列</Chip>
            <Chip onClick={() => ed.table('delCol')}>删除列</Chip>
          </div>
        </div>
      )}
      <div className="m-toolbar word-m-bar" role="toolbar" aria-label="Word 编辑">
        <Tool icon="check" label="完成" onClick={() => (setPanel(null), props.onMode('view'))} />
        {rv.trackOn && <Tool icon="contacts" label="审阅" onClick={() => (setPanel(null), props.onMode('review'))} />}
        <Tool icon="undo" label="撤销" disabled={!ed.canUndo} onClick={ed.undo} />
        <Tool icon="redo" label="重做" disabled={!ed.canRedo} onClick={ed.redo} />
        <Tool icon="cursor" label="选择" on={ed.selecting} onClick={() => ed.setSelecting(!ed.selecting)} />
        <Tool glyph="A" label="文字" on={panel === 'text'} onClick={() => toggle('text')} />
        <Tool icon="viewLines" label="段落" on={panel === 'para'} onClick={() => toggle('para')} />
        {ed.inTable && <Tool icon="selectAll" label="表格" on={panel === 'table'} onClick={() => toggle('table')} />}
        <Tool icon="upload" label="图片" onClick={ed.insertImage} />
        <Tool icon="copy" label="复制" disabled={!ed.hasSelection} onClick={() => ed.clip('copy')} />
        <Tool icon="cut" label="剪切" disabled={!ed.hasSelection} onClick={() => ed.clip('cut')} />
        <Tool icon="paste" label="粘贴" onClick={() => ed.clip('paste')} />
      </div>
    </>
  );
}

/**
 * 底部工作表标签：点击切换、双击改名、右键菜单（插入 / 删除 / 重命名 / 复制 / 隐藏 / 取消隐藏 / 标签颜色）、
 * 拖动调整顺序、“+”新建。改动通过回调交给编辑器。
 */
import { useEffect, useRef, useState } from 'react';
import { host } from '@wz/editor-shared';
import { STANDARD_COLORS } from '../core/format';
import type { SheetData } from '../core/workbook';
import { Icon } from '@wz/editor-shared/Icon';

interface Props {
  sheets: SheetData[];
  current: number;
  onSelect: (i: number) => void;
  onAdd: () => void;
  /** 返回错误信息（不合法时），成功返回 null */
  onRename: (i: number, name: string) => string | null;
  onDelete: (i: number) => void;
  onCopy: (i: number) => void;
  onMove: (from: number, to: number) => void;
  onHide: (i: number) => void;
  onUnhide: (i: number) => void;
  onTabColor: (i: number, color: string | null) => void;
  onDone: () => void;
}

export function SheetTabs(p: Props) {
  const [menu, setMenu] = useState<{ i: number; x: number; y: number; sub?: 'color' | 'unhide' } | null>(null);
  const [renaming, setRenaming] = useState<{ i: number; text: string } | null>(null);
  const [drag, setDrag] = useState<{ from: number; over: number } | null>(null);
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!menu) return;
    host().pushPopup();
    const close = () => setMenu(null);
    window.addEventListener('click', close);
    window.addEventListener('blur', close);
    return () => {
      host().popPopup();
      window.removeEventListener('click', close);
      window.removeEventListener('blur', close);
    };
  }, [menu]);

  useEffect(() => {
    if (renaming) requestAnimationFrame(() => input.current?.select());
  }, [renaming?.i]); // eslint-disable-line react-hooks/exhaustive-deps

  const commitRename = () => {
    if (!renaming) return;
    const err = p.onRename(renaming.i, renaming.text);
    if (err) {
      host().toast(err);
      input.current?.focus();
      return;
    }
    setRenaming(null);
    p.onDone();
  };
  const act = (fn: () => void) => () => {
    setMenu(null);
    fn();
    p.onDone();
  };
  const hidden = p.sheets.map((s, i) => ({ s, i })).filter((x) => x.s.hidden);

  return (
    <>
      <div className="sheet-tabs" role="tablist" aria-label="工作表" onDragOver={(e) => e.preventDefault()}>
        {p.sheets.map((s, i) =>
          s.hidden ? null : renaming?.i === i ? (
            <input
              key={i}
              ref={input}
              className="sheet-tab-input"
              value={renaming.text}
              aria-label="工作表名称"
              autoFocus
              onChange={(e) => setRenaming({ i, text: e.target.value })}
              onBlur={commitRename}
              onKeyDown={(e) => {
                // 阻止默认动作：否则焦点回到表格后这次回车会被当成输入
                if (e.key === 'Enter') {
                  e.preventDefault();
                  commitRename();
                }
                if (e.key === 'Escape') {
                  e.preventDefault();
                  setRenaming(null);
                  p.onDone();
                }
              }}
            />
          ) : (
            <button
              key={i}
              role="tab"
              aria-selected={i === p.current}
              className={`sheet-tab${i === p.current ? ' on' : ''}${drag?.over === i && drag.from !== i ? ' drop' : ''}`}
              style={s.tabColor ? ({ '--tab-color': s.tabColor } as React.CSSProperties) : undefined}
              draggable
              onDragStart={(e) => {
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', s.name);
                setDrag({ from: i, over: i });
              }}
              onDragEnter={() => drag && setDrag({ ...drag, over: i })}
              onDragEnd={() => setDrag(null)}
              onDrop={(e) => {
                e.preventDefault();
                if (drag && drag.from !== i) p.onMove(drag.from, i);
                setDrag(null);
              }}
              onClick={() => {
                p.onSelect(i);
                p.onDone();
              }}
              onDoubleClick={() => setRenaming({ i, text: s.name })}
              onContextMenu={(e) => {
                e.preventDefault();
                e.stopPropagation();
                // 不切换工作表：切换会重建标签栏，菜单就没了；菜单只作用于点中的这张表
                setMenu({ i, x: e.clientX, y: e.clientY });
              }}
            >
              <span className={`dot${s.tabColor ? ' color' : ''}`} />
              {s.name}
            </button>
          ),
        )}
        <button className="icon-btn sheet-tab-add" title="新建工作表（Shift+F11）" aria-label="新建工作表" onClick={() => (p.onAdd(), p.onDone())}>
          <Icon name="plus" size={14} />
        </button>
      </div>

      {menu && (
        <div className="ctx-menu" role="menu" style={{ left: menu.x, top: Math.max(8, menu.y - 300) }} onClick={(e) => e.stopPropagation()} onMouseDown={(e) => e.preventDefault()}>
          {!menu.sub && (
            <>
              <button role="menuitem" onClick={act(p.onAdd)}>
                插入工作表
              </button>
              <button role="menuitem" onClick={act(() => p.onDelete(menu.i))}>
                删除
              </button>
              <button role="menuitem" onClick={() => (setMenu(null), setRenaming({ i: menu.i, text: p.sheets[menu.i].name }))}>
                重命名
              </button>
              <button role="menuitem" onClick={act(() => p.onCopy(menu.i))}>
                复制工作表
              </button>
              <div className="ctx-sep" />
              <button role="menuitem" onClick={() => setMenu({ ...menu, sub: 'color' })}>
                工作表标签颜色<span className="grow" />›
              </button>
              <div className="ctx-sep" />
              <button role="menuitem" onClick={act(() => p.onHide(menu.i))}>
                隐藏
              </button>
              <button role="menuitem" disabled={!hidden.length} onClick={() => setMenu({ ...menu, sub: 'unhide' })}>
                取消隐藏…<span className="grow" />›
              </button>
            </>
          )}
          {menu.sub === 'color' && (
            <div className="sheet-palette">
              <button className="sheet-palette-wide" onClick={act(() => p.onTabColor(menu.i, null))}>
                <span className="swatch none" />
                无颜色
              </button>
              <div className="sheet-palette-grid">
                {STANDARD_COLORS.map((c) => (
                  <button key={c} className="swatch" style={{ background: c }} title={c} onClick={act(() => p.onTabColor(menu.i, c))} />
                ))}
              </div>
            </div>
          )}
          {menu.sub === 'unhide' &&
            hidden.map(({ s, i }) => (
              <button key={i} role="menuitem" onClick={act(() => p.onUnhide(i))}>
                {s.name}
              </button>
            ))}
        </div>
      )}
    </>
  );
}

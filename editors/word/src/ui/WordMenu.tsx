/** 右键菜单：点菜单外面、按 Esc、窗口失焦或滚动都会关闭；有子菜单的项鼠标移上去在右边展开。 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';

export interface MenuItem {
  label: string;
  /** 右边灰色的快捷键说明 */
  keys?: string;
  onClick?: () => void;
  disabled?: boolean;
  /** 在这一项上面画一条分隔线 */
  sep?: boolean;
  sub?: MenuItem[];
}

function MenuPanel({ x, y, items, onClose, sub = false }: { x: number; y: number; items: MenuItem[]; onClose: () => void; sub?: boolean }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ left: x, top: y });
  const [open, setOpen] = useState<{ i: number; x: number; y: number } | null>(null);
  // 不超出窗口：右边放不下往左移，下面放不下往上移
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({ left: Math.max(4, Math.min(x, window.innerWidth - r.width - 4)), top: Math.max(4, Math.min(y, window.innerHeight - r.height - 4)) });
  }, [x, y]);
  return (
    <div ref={ref} className={`ctx-menu word-menu${sub ? ' ctx-sub-panel' : ''}`} role="menu" style={pos} onContextMenu={(e) => e.preventDefault()} onMouseDown={(e) => e.preventDefault()}>
      {items.map((it, i) => (
        <div key={i}>
          {it.sep && <div className="ctx-sep" />}
          <button
            type="button"
            role="menuitem"
            disabled={it.disabled}
            className={open?.i === i ? 'on' : undefined}
            aria-haspopup={it.sub ? 'menu' : undefined}
            onMouseEnter={(e) => {
              if (!it.sub) return setOpen(null);
              const r = e.currentTarget.getBoundingClientRect();
              setOpen({ i, x: r.right + 2, y: r.top - 6 });
            }}
            onClick={(e) => {
              if (it.sub) {
                const r = e.currentTarget.getBoundingClientRect();
                return setOpen({ i, x: r.right + 2, y: r.top - 6 });
              }
              onClose();
              it.onClick?.();
            }}
          >
            <span className="word-menu-label">{it.label}</span>
            {it.keys && <span className="word-menu-keys">{it.keys}</span>}
            {it.sub && <span className="ctx-sub-arrow">›</span>}
          </button>
        </div>
      ))}
      {open && items[open.i]?.sub && <MenuPanel x={open.x} y={open.y} items={items[open.i].sub!} onClose={onClose} sub />}
    </div>
  );
}

export function WordMenu({ x, y, items, onClose }: { x: number; y: number; items: MenuItem[]; onClose: () => void }) {
  const box = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const down = (e: MouseEvent) => {
      if (!box.current?.contains(e.target as Node)) onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      e.preventDefault();
      e.stopPropagation();
      onClose();
    };
    window.addEventListener('mousedown', down, true);
    window.addEventListener('keydown', key, true);
    window.addEventListener('blur', onClose);
    window.addEventListener('wheel', onClose, { passive: true });
    return () => {
      window.removeEventListener('mousedown', down, true);
      window.removeEventListener('keydown', key, true);
      window.removeEventListener('blur', onClose);
      window.removeEventListener('wheel', onClose);
    };
  }, [onClose]);
  return (
    <div ref={box}>
      <MenuPanel x={x} y={y} items={items} onClose={onClose} />
    </div>
  );
}

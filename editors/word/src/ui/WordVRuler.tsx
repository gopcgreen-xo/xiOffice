/**
 * 纵向标尺：和光标所在页对齐，显示上下页边距（灰色）和厘米刻度；拖动页边距线改上 / 下页边距，
 * 松开后一次性修改（一步撤销），拖动时显示一条横贯页面的参考线。取值规则在 core/ruler.ts（marginPatchV）。
 */
import { useRef, useState } from 'react';
import type { PageSetup } from '../core/model';
import { rulerTicks, snap } from '../core/ruler';

const SNAP_PT = 72 / 2.54 / 10; // 0.1 厘米

type Drag = { side: 'top' | 'bottom'; pt: number };

export function WordVRuler({
  ppt,
  pageTop,
  setup,
  onMargin,
}: {
  ppt: number;
  /** 页面上边在标尺里的纵坐标（像素，可以为负：页面顶部已经滚出视野） */
  pageTop: number;
  setup: PageSetup;
  onMargin: (side: 'top' | 'bottom', pagePt: number) => void;
}) {
  const bar = useRef<HTMLDivElement>(null);
  const [drag, setDrag] = useState<Drag | null>(null);
  const dragRef = useRef<Drag | null>(null);
  const set = (d: Drag | null) => {
    dragRef.current = d;
    setDrag(d);
  };
  const px = (pt: number) => pageTop + pt * ppt;
  const pagePt = (e: { clientY: number; altKey: boolean }) => {
    const r = bar.current!.getBoundingClientRect();
    const v = (e.clientY - r.top - pageTop) / ppt;
    return e.altKey ? v : snap(v - setup.marginTop, SNAP_PT) + setup.marginTop;
  };
  const start = (side: 'top' | 'bottom') => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.setPointerCapture(e.pointerId);
    set({ side, pt: pagePt(e) });
  };
  const end = () => {
    const d = dragRef.current;
    set(null);
    if (d) onMargin(d.side, d.pt);
  };

  const tm = drag?.side === 'top' ? drag.pt : setup.marginTop;
  const bm = drag?.side === 'bottom' ? drag.pt : setup.height - setup.marginBottom;
  const ticks = rulerTicks(setup.marginTop, 0, setup.height);

  return (
    <>
      <div
        ref={bar}
        className="word-vruler"
        aria-label="纵向标尺"
        onPointerMove={(e) => dragRef.current && set({ ...dragRef.current, pt: pagePt(e) })}
        onPointerUp={end}
        onPointerCancel={() => set(null)}
      >
        <div className="word-vruler-page" style={{ top: px(0), height: setup.height * ppt }}>
          <div className="word-ruler-margin" style={{ top: 0, height: tm * ppt }} />
          <div className="word-ruler-margin" style={{ top: bm * ppt, height: (setup.height - bm) * ppt }} />
        </div>
        {ticks.map((t) => (
          <span key={t.cm} className={`word-vruler-tick ${t.major ? 'major' : 'half'}`} style={{ top: px(t.pt) }}>
            {t.major && t.cm !== 0 ? <i>{Math.abs(t.cm)}</i> : null}
          </span>
        ))}
        <div className="word-vruler-edge" title="拖动调整上页边距" style={{ top: px(tm) - 3 }} onPointerDown={start('top')} />
        <div className="word-vruler-edge" title="拖动调整下页边距" style={{ top: px(bm) - 3 }} onPointerDown={start('bottom')} />
      </div>
      {drag && <div className="word-vruler-guide" style={{ top: px(drag.pt) }} />}
    </>
  );
}

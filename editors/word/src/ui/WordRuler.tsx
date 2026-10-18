/**
 * 水平标尺：和光标所在页对齐，显示页边距（灰色）、厘米刻度、当前段落的缩进标记和制表位。
 * 拖动标记 → 松开后一次性改格式（一步撤销）；拖动时显示一条贯穿页面的参考线。
 * 取值规则和换算在 core/ruler.ts（纯函数，有测试）。
 */
import { useRef, useState } from 'react';
import type { PageSetup, ParaProps, TabStop } from '../core/model';
import { addTab, indentPatch, moveTab, rulerMarks, rulerTicks, snap, tabAt, type IndentMarker, type RulerArea, type TabAlign } from '../core/ruler';

/** 标尺高度（像素），CSS 里同值 */
export const RULER_H = 24;

const SNAP_PT = 72 / 2.54 / 10; // 0.1 厘米
const TAB_NEXT: Record<TabAlign, TabAlign> = { left: 'center', center: 'right', right: 'left' };
const TAB_LABEL: Record<TabAlign, string> = { left: '左对齐制表位', center: '居中制表位', right: '右对齐制表位' };

type Drag =
  | { kind: 'indent'; marker: IndentMarker; pt: number }
  | { kind: 'margin'; side: 'left' | 'right'; pt: number }
  | { kind: 'tab'; index: number; pt: number; off: boolean };

export function WordRuler({
  ppt,
  pageLeft,
  setup,
  area,
  para,
  fontSize,
  onIndent,
  onMargin,
  onTabs,
}: {
  ppt: number;
  /** 页面左边在标尺里的横坐标（像素） */
  pageLeft: number;
  setup: PageSetup;
  /** 光标所在段落的区域；光标不在正文段落里（表格、页眉页脚）时为 null，只显示页边距 */
  area: RulerArea | null;
  para: ParaProps | null;
  fontSize: number;
  onIndent: (patch: Partial<ParaProps>) => void;
  onMargin: (side: 'left' | 'right', pagePt: number) => void;
  onTabs: (tabs: TabStop[]) => void;
}) {
  const bar = useRef<HTMLDivElement>(null);
  const [drag, setDrag] = useState<Drag | null>(null);
  const [tabKind, setTabKind] = useState<TabAlign>('left');
  const dragRef = useRef<Drag | null>(null);

  const px = (pagePt: number) => pageLeft + pagePt * ppt;
  const pagePt = (e: { clientX: number; altKey: boolean }) => {
    const r = bar.current!.getBoundingClientRect();
    const v = (e.clientX - r.left - pageLeft) / ppt;
    return e.altKey ? v : snap(v - setup.marginLeft, SNAP_PT) + setup.marginLeft;
  };
  const set = (d: Drag | null) => {
    dragRef.current = d;
    setDrag(d);
  };

  const marks = para && area ? rulerMarks(para, area.w, fontSize) : null;
  const tabs = para?.tabs ?? [];

  // 拖动中的显示位置：按拖动的那个标记算出新格式，再换回各标记的位置（拖悬挂标记时首行标记不动，拖左缩进时两个一起动）
  const dragged =
    drag?.kind === 'indent' && area && para ? indentPatch(drag.marker, drag.pt - area.x, para, area, fontSize) : null;
  const live = para && area ? rulerMarks(dragged ? { ...para, ...dragged } : para, area.w, fontSize) : null;

  const startIndent = (marker: IndentMarker) => (e: React.PointerEvent) => {
    if (!area) return;
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.setPointerCapture(e.pointerId);
    set({ kind: 'indent', marker, pt: pagePt(e) });
  };
  const startMargin = (side: 'left' | 'right') => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.setPointerCapture(e.pointerId);
    set({ kind: 'margin', side, pt: pagePt(e) });
  };
  const startTab = (index: number) => (e: React.PointerEvent) => {
    if (!area) return;
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.setPointerCapture(e.pointerId);
    set({ kind: 'tab', index, pt: pagePt(e), off: false });
  };
  /** 点标尺空白处：新建制表位（继续拖可以调位置） */
  const startNewTab = (e: React.PointerEvent) => {
    if (!area || !para) return;
    const pt = pagePt(e) - area.x;
    if (pt < 0 || pt > area.w) return;
    const hit = tabAt(tabs, pt, 4 / ppt);
    e.preventDefault();
    e.currentTarget.setPointerCapture(e.pointerId);
    set({ kind: 'tab', index: hit >= 0 ? hit : -1, pt: pagePt(e), off: false });
  };

  const move = (e: React.PointerEvent) => {
    const d = dragRef.current;
    if (!d) return;
    const r = bar.current!.getBoundingClientRect();
    const off = e.clientY < r.top - 12 || e.clientY > r.bottom + 28;
    set(d.kind === 'tab' ? { ...d, pt: pagePt(e), off } : { ...d, pt: pagePt(e) });
  };

  const end = () => {
    const d = dragRef.current;
    set(null);
    if (!d) return;
    if (d.kind === 'margin') return onMargin(d.side, d.pt);
    if (!area || !para) return;
    if (d.kind === 'indent') {
      const patch = indentPatch(d.marker, d.pt - area.x, para, area, fontSize);
      if (patch) onIndent(patch);
      return;
    }
    const pos = d.pt - area.x;
    if (d.index < 0) {
      const next = addTab(tabs, pos, tabKind, area.w);
      if (next) onTabs(next);
    } else {
      onTabs(moveTab(tabs, d.index, d.off ? null : pos, area.w));
    }
  };

  const ticks = rulerTicks(setup.marginLeft, 0, setup.width);
  const lm = drag?.kind === 'margin' && drag.side === 'left' ? drag.pt : setup.marginLeft;
  const rm = drag?.kind === 'margin' && drag.side === 'right' ? drag.pt : setup.width - setup.marginRight;
  const guide = drag ? px(drag.pt) : null;

  const first = live?.first ?? 0;
  const left = live?.left ?? 0;
  const hang = left;
  const right = live?.right ?? 0;
  const ax = area?.x ?? 0;

  return (
    <>
      <div className="word-ruler" aria-label="标尺">
        <button
          className="word-ruler-tabkind"
          title={`${TAB_LABEL[tabKind]}（点击切换；在标尺上单击添加制表位，拖出标尺删除）`}
          disabled={!area}
          onClick={() => setTabKind(TAB_NEXT[tabKind])}
        >
          <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
            <path d="M6 1v8" stroke="currentColor" strokeWidth="1.5" fill="none" />
            <path d={tabKind === 'left' ? 'M6 9h5' : tabKind === 'right' ? 'M1 9h5' : 'M1 9h10'} stroke="currentColor" strokeWidth="1.5" fill="none" />
          </svg>
        </button>
        <div ref={bar} className="word-ruler-bar" onPointerMove={move} onPointerUp={end} onPointerCancel={() => set(null)}>
          <div className="word-ruler-page" style={{ left: px(0), width: setup.width * ppt }}>
            <div className="word-ruler-margin" style={{ left: 0, width: lm * ppt }} />
            <div className="word-ruler-margin" style={{ left: rm * ppt, width: (setup.width - rm) * ppt }} />
          </div>
          <div className="word-ruler-track" style={{ left: px(lm), width: Math.max(0, (rm - lm) * ppt) }} onPointerDown={startNewTab} />
          {ticks.map((t) =>
            t.major ? (
              <span key={t.cm} className="word-ruler-tick major" style={{ left: px(t.pt) }}>
                {Math.abs(t.cm) > 0 ? <i>{Math.abs(t.cm)}</i> : null}
              </span>
            ) : (
              <span key={t.cm} className="word-ruler-tick half" style={{ left: px(t.pt) }} />
            ),
          )}
          <div className="word-ruler-edge" title="拖动调整页边距" style={{ left: px(lm) - 3 }} onPointerDown={startMargin('left')} />
          <div className="word-ruler-edge" title="拖动调整页边距" style={{ left: px(rm) - 3 }} onPointerDown={startMargin('right')} />
          {marks && area && (
            <>
              {tabs.map((t, i) => {
                const pos = drag?.kind === 'tab' && drag.index === i ? drag.pt - area.x : t.pos;
                const off = drag?.kind === 'tab' && drag.index === i && drag.off;
                return (
                  <button
                    key={`${i}-${t.pos}`}
                    className={`word-ruler-tab ${t.align}${off ? ' off' : ''}`}
                    title={`${TAB_LABEL[t.align]}：拖动调整位置，拖出标尺删除`}
                    style={{ left: px(ax + pos) - 5 }}
                    onPointerDown={startTab(i)}
                  />
                );
              })}
              {drag?.kind === 'tab' && drag.index < 0 && <span className={`word-ruler-tab ${tabKind}`} style={{ left: px(drag.pt) - 5 }} />}
              <button className="word-ruler-mark first" title="首行缩进" style={{ left: px(ax + first) - 5 }} onPointerDown={startIndent('first')} />
              <button className="word-ruler-mark hanging" title="悬挂缩进" style={{ left: px(ax + hang) - 5 }} onPointerDown={startIndent('hanging')} />
              <button className="word-ruler-mark left" title="左缩进" style={{ left: px(ax + left) - 5 }} onPointerDown={startIndent('left')} />
              <button className="word-ruler-mark right" title="右缩进" style={{ left: px(ax + right) - 5 }} onPointerDown={startIndent('right')} />
            </>
          )}
        </div>
      </div>
      {guide !== null && <div className="word-ruler-guide" style={{ left: guide + 28 }} />}
    </>
  );
}

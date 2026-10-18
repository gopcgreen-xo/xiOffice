/**
 * 手机版的触摸手势（docs/ANDROID.md §7），Word / Excel / PDF 共用：
 * - 双指缩放：手指移动时只给内容加 CSS 缩放（不重新排版 / 绘制，跟手），松开后才提交新的缩放比例，
 *   并调整滚动位置让两指中点下的内容留在原处；
 * - 单指滚动交给浏览器（容器设 touch-action: pan-x pan-y，浏览器自己的页面缩放关掉）。
 */
import { useEffect, useRef, type RefObject } from 'react';

export interface PinchOptions {
  min: number;
  max: number;
  /** 提交新的缩放（已限制在 min–max） */
  onZoom(zoom: number): void;
}

/** 两指缩放后新的缩放比例（限制范围、保留两位小数） */
export function pinchZoom(start: number, startDist: number, dist: number, min: number, max: number): number {
  if (startDist <= 0) return start;
  const z = start * (dist / startDist);
  return Math.round(Math.min(max, Math.max(min, z)) * 100) / 100;
}

/**
 * 缩放后保持焦点不动的滚动位置：焦点在内容里的坐标按比例放大，再减去它在视口里的位置。
 * offset = 内容在滚动区域里的起点（页面居中时左右的空白），不随缩放变化的那部分。
 */
export function anchorScroll(scroll: number, focal: number, k: number, offset = 0): number {
  return Math.max(0, (scroll + focal - offset) * k + offset - focal);
}

const dist = (t: TouchList) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
const mid = (t: TouchList) => [(t[0].clientX + t[1].clientX) / 2, (t[0].clientY + t[1].clientY) / 2] as const;

/**
 * 给滚动容器加双指缩放。content = 被缩放的内容元素（预览时加 transform）；zoom = 当前缩放。
 * 返回一个 ref，表示是否正在缩放（期间别处理单击）。
 */
export function usePinchZoom(scroller: RefObject<HTMLElement | null>, content: RefObject<HTMLElement | null>, zoom: number, opts: PinchOptions | null) {
  const pinching = useRef(false);
  const live = useRef({ zoom, opts });
  live.current = { zoom, opts };
  useEffect(() => {
    const sc = scroller.current;
    if (!sc || !opts) return;
    let start: { d: number; z: number; fx: number; fy: number; sl: number; st: number } | null = null;
    let k = 1;
    const onStart = (e: TouchEvent) => {
      if (e.touches.length !== 2) return;
      const r = sc.getBoundingClientRect();
      const [x, y] = mid(e.touches);
      start = { d: dist(e.touches), z: live.current.zoom, fx: x - r.left, fy: y - r.top, sl: sc.scrollLeft, st: sc.scrollTop };
      k = 1;
      pinching.current = true;
    };
    const onMove = (e: TouchEvent) => {
      if (!start || e.touches.length !== 2) return;
      e.preventDefault();
      const o = live.current.opts!;
      const z = pinchZoom(start.z, start.d, dist(e.touches), o.min, o.max);
      k = z / start.z;
      const el = content.current;
      if (el) {
        el.style.transformOrigin = `${start.sl + start.fx}px ${start.st + start.fy}px`;
        el.style.transform = `scale(${k})`;
      }
    };
    const onEnd = (e: TouchEvent) => {
      if (!start || e.touches.length >= 2) return;
      const s = start;
      start = null;
      const el = content.current;
      if (el) {
        el.style.transform = '';
        el.style.transformOrigin = '';
      }
      const o = live.current.opts!;
      if (Math.abs(k - 1) > 0.01) {
        o.onZoom(Math.round(s.z * k * 100) / 100);
        // 新缩放渲染后再定位滚动
        requestAnimationFrame(() =>
          requestAnimationFrame(() => {
            sc.scrollLeft = anchorScroll(s.sl, s.fx, k);
            sc.scrollTop = anchorScroll(s.st, s.fy, k);
          }),
        );
      }
      setTimeout(() => (pinching.current = false), 50);
    };
    sc.addEventListener('touchstart', onStart, { passive: true });
    sc.addEventListener('touchmove', onMove, { passive: false });
    sc.addEventListener('touchend', onEnd);
    sc.addEventListener('touchcancel', onEnd);
    return () => {
      sc.removeEventListener('touchstart', onStart);
      sc.removeEventListener('touchmove', onMove);
      sc.removeEventListener('touchend', onEnd);
      sc.removeEventListener('touchcancel', onEnd);
    };
  }, [scroller, content, !!opts]);
  return pinching;
}

/** 按容器宽度适应页面宽度的缩放：pageWidthPx = 缩放 1 时的页面宽度（像素），pad = 两侧留白 */
export function fitWidthZoom(containerWidth: number, pageWidthPx: number, pad = 12, min = 0.2, max = 3): number {
  if (containerWidth <= 0 || pageWidthPx <= 0) return 1;
  return Math.round(Math.min(max, Math.max(min, (containerWidth - pad * 2) / pageWidthPx)) * 1000) / 1000;
}

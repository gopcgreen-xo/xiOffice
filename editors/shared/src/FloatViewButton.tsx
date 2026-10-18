import { host } from './host';
import { Icon } from './Icon';

/**
 * 编辑器工具条里的“画中画”按钮：框选显示区，浮成置顶的小窗口浏览（工作台的 host().floatView）。
 * 没有这个能力、文件不在磁盘上（新建未保存、工作包里的虚拟路径）时不显示；悬浮窗口自己里面也不显示（样式隐藏）。
 */
export function FloatViewButton({ tabId, path, className = 'icon-btn' }: { tabId: string; path?: string | null; className?: string }) {
  const h = host();
  if (!h.floatView || !path || path.includes('!/')) return null;
  return (
    <button className={`${className} ed-float-btn`} title="画中画: 框选后悬浮显示" aria-label="画中画" onClick={() => h.floatView!(tabId)}>
      <Icon name="pip" size={16} />
    </button>
  );
}

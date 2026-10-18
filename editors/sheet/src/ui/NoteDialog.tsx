/** 批注编辑对话框：多行文字；清空文字再确定 = 删除批注。 */
import { useEffect, useState } from 'react';
import { host } from '@wz/editor-shared';
import { Icon } from '@wz/editor-shared/Icon';

interface Props {
  /** 格子地址（标题里显示） */
  cell: string;
  initial: string;
  author: string;
  /** 已有批注时才有“删除” */
  exists: boolean;
  onOk: (text: string) => void;
  onDelete: () => void;
  onClose: () => void;
}

export function NoteDialog(p: Props) {
  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  const [text, setText] = useState(p.initial);
  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-sort" role="dialog" aria-label="批注" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => (e.key === 'Escape' ? p.onClose() : e.key === 'Enter' && (e.ctrlKey || e.metaKey) && p.onOk(text))}>
        <div className="sheet-detail-head">
          <strong>{p.exists ? '编辑批注' : '插入批注'}</strong>
          <span className="sub mono">{p.cell}</span>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        <textarea aria-label="批注内容" className="input" style={{ width: '100%', height: 140, resize: 'vertical', padding: 8, boxSizing: 'border-box' }} autoFocus value={text} onChange={(e) => setText(e.target.value)} placeholder="输入批注内容" />
        <div className="sub" style={{ padding: '4px 0' }}>
          {p.author ? `作者：${p.author}　` : ''}Ctrl+Enter 确定
        </div>
        <div className="sheet-detail-foot">
          {p.exists && (
            <button className="btn" onClick={p.onDelete}>
              删除批注
            </button>
          )}
          <span className="grow" />
          <button className="btn" onClick={p.onClose}>
            取消
          </button>
          <button className="btn primary" onClick={() => p.onOk(text)}>
            确定
          </button>
        </div>
      </div>
    </div>
  );
}

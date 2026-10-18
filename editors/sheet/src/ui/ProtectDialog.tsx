/** 保护对话框：保护工作表（密码 + 允许的操作）、撤销保护（输入密码）、保护 / 撤销工作簿结构保护。 */
import { useEffect, useState } from 'react';
import { host } from '@wz/editor-shared';
import { Icon } from '@wz/editor-shared/Icon';
import { ALLOW_KEYS, allowOf, DEFAULT_ALLOW, type SheetAllow } from '../core/protect';

export type ProtectKind = 'sheet' | 'sheet-off' | 'book' | 'book-off';

const ALLOW_LABEL: Record<keyof SheetAllow, string> = {
  selectLockedCells: '选定锁定的单元格',
  selectUnlockedCells: '选定未锁定的单元格',
  formatCells: '设置单元格格式',
  formatColumns: '设置列格式',
  formatRows: '设置行格式',
  insertColumns: '插入列',
  insertRows: '插入行',
  insertHyperlinks: '插入超链接',
  deleteColumns: '删除列',
  deleteRows: '删除行',
  sort: '排序',
  autoFilter: '使用自动筛选',
  pivotTables: '使用数据透视表',
  objects: '编辑对象（图片、图表）',
};

interface Props {
  kind: ProtectKind;
  title: string;
  /** 保护工作表时的初始选项 */
  initial?: Record<string, string>;
  /** 确认：保护时给密码和选项；撤销时给输入的密码。返回错误文字（留在对话框里提示）或 null（成功） */
  onOk: (password: string, allow: SheetAllow) => Promise<string | null> | string | null;
  onClose: () => void;
}

export function ProtectDialog(p: Props) {
  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  const off = p.kind === 'sheet-off' || p.kind === 'book-off';
  const [pw, setPw] = useState('');
  const [pw2, setPw2] = useState('');
  const [allow, setAllow] = useState<SheetAllow>(() => (p.initial ? allowOf(p.initial) : DEFAULT_ALLOW));
  const [err, setErr] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const ok = async () => {
    if (!off && pw !== pw2) return setErr('两次输入的密码不一样');
    setBusy(true);
    const bad = await p.onOk(pw, allow);
    setBusy(false);
    if (bad) setErr(bad);
  };

  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-sort" role="dialog" aria-label={p.title} onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => (e.key === 'Escape' ? p.onClose() : e.key === 'Enter' && !busy && void ok())}>
        <div className="sheet-detail-head">
          <strong>{p.title}</strong>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        {p.kind === 'sheet' && (
          <div className="sub" style={{ padding: '4px 0' }}>
            保护后只能修改“未锁定”的单元格（右键 → 锁定 / 取消锁定单元格）。密码可以不设，只是防止误改，不是加密。
          </div>
        )}
        {p.kind === 'book' && (
          <div className="sub" style={{ padding: '4px 0' }}>
            保护后不能新建、删除、重命名、移动、隐藏或取消隐藏工作表。密码可以不设。
          </div>
        )}
        <div className="names-form">
          <label>
            <span>密码</span>
            <input aria-label="密码" type="password" className="input" value={pw} autoFocus onChange={(e) => (setPw(e.target.value), setErr(null))} placeholder={off ? '输入密码' : '（可不填）'} />
          </label>
          {!off && (
            <label>
              <span>再输一次</span>
              <input aria-label="确认密码" type="password" className="input" value={pw2} onChange={(e) => (setPw2(e.target.value), setErr(null))} />
            </label>
          )}
        </div>
        {p.kind === 'sheet' && (
          <>
            <div className="sub" style={{ padding: '8px 0 2px' }}>
              允许所有用户进行：（在 Excel 里生效；本编辑器里受保护的工作表只能改未锁定的单元格）
            </div>
            <div className="sheet-data-cols">
              {ALLOW_KEYS.map((k) => (
                <label key={k} style={{ width: 180 }}>
                  <input type="checkbox" checked={allow[k]} onChange={(e) => setAllow({ ...allow, [k]: e.target.checked })} />
                  {ALLOW_LABEL[k]}
                </label>
              ))}
            </div>
          </>
        )}
        {err && <div className="ed-warn">{err}</div>}
        <div className="sheet-detail-foot">
          <span className="grow" />
          <button className="btn" onClick={p.onClose}>
            取消
          </button>
          <button className="btn primary" disabled={busy} onClick={() => void ok()}>
            {busy ? '正在验证…' : '确定'}
          </button>
        </div>
      </div>
    </div>
  );
}

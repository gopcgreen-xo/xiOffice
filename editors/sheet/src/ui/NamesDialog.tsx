/** 名称管理器：列出定义名称，新建 / 修改 / 删除；引用位置可以一键取当前选区。 */
import { useEffect, useState } from 'react';
import { host } from '@wz/editor-shared';
import { Icon } from '@wz/editor-shared/Icon';
import { nameProblem, nameTarget, refersToProblem, type DefinedName } from '../core/names';
import type { Workbook } from '../core/workbook';

interface Props {
  book: Workbook;
  /** 当前工作表序号（“范围”默认选它的工作表） */
  si: number;
  /** 当前选区写成 Sheet1!$A$1:$B$5 */
  selectionRef: string;
  /** 打开时就在新建状态，并带上这些初值（“用选区定义名称”） */
  draft?: { name?: string; f?: string };
  onSave: (def: DefinedName, replace?: DefinedName) => string | null;
  onDelete: (def: DefinedName) => void;
  onGoto: (def: DefinedName) => void;
  onClose: () => void;
}

const scopeText = (book: Workbook, local: number | undefined) => (local === undefined ? '工作簿' : (book.sheets[local]?.name ?? '?'));

export function NamesDialog(p: Props) {
  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  const names = p.book.names ?? [];
  // editing = 正在修改的那一项；null = 新建
  const [editing, setEditing] = useState<DefinedName | null>(null);
  const [name, setName] = useState(p.draft?.name ?? '');
  const [f, setF] = useState(p.draft?.f ?? p.selectionRef);
  const [scope, setScope] = useState<string>('book');
  const [shown, setShown] = useState(!!p.draft || names.length === 0);

  const local = scope === 'book' ? undefined : Number(scope);
  const problem = name ? (nameProblem(p.book, name, local, editing ?? undefined) ?? refersToProblem(f)) : null;
  const pick = (n: DefinedName | null) => {
    setEditing(n);
    setShown(true);
    setName(n?.name ?? '');
    setF(n ? `=${n.f}` : p.selectionRef);
    setScope(n?.local === undefined ? 'book' : String(n.local));
  };
  const save = () => {
    if (!name.trim()) return host().toast('请输入名称');
    const bad = p.onSave({ name: name.trim(), f, ...(local !== undefined && { local }) }, editing ?? undefined);
    if (bad) return host().toast(bad);
    setEditing(null);
    setShown(false);
  };

  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-names" role="dialog" aria-label="名称管理器" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
        <div className="sheet-detail-head">
          <strong>名称管理器</strong>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        <div className="names-list" role="list">
          {names.length === 0 && <div className="sub" style={{ padding: 8 }}>还没有定义名称。名称可以代表一个区域、常量或公式，在公式里直接用，如 =SUM(销售额)。</div>}
          {names.map((n) => (
            <div key={`${n.local ?? ''}|${n.name}`} role="listitem" className={`names-row${editing === n ? ' on' : ''}`} onClick={() => pick(n)} onDoubleClick={() => nameTarget(p.book, n.name, p.si) && p.onGoto(n)}>
              <span className="names-name mono ellipsis" title={n.name}>
                {n.name}
              </span>
              <span className="names-f mono ellipsis grow" title={n.f}>
                {n.f}
              </span>
              <span className="sub names-scope ellipsis">{scopeText(p.book, n.local)}</span>
              <button
                className="icon-btn"
                aria-label={`删除 ${n.name}`}
                title="删除"
                onClick={(e) => {
                  e.stopPropagation();
                  p.onDelete(n);
                  if (editing === n) pick(null);
                }}
              >
                <Icon name="close" size={11} />
              </button>
            </div>
          ))}
        </div>
        {!shown && (
          <div className="sheet-detail-foot">
            <button className="btn primary" onClick={() => pick(null)}>
              新建…
            </button>
            <span className="grow" />
            <button className="btn" onClick={p.onClose}>
              关闭
            </button>
          </div>
        )}
        {shown && (
          <div className="names-form">
            <label>
              <span>名称</span>
              <input aria-label="名称" className="input" value={name} autoFocus onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && save()} />
            </label>
            <label>
              <span>范围</span>
              <select aria-label="范围" className="ed-select" value={scope} onChange={(e) => setScope(e.target.value)}>
                <option value="book">工作簿</option>
                {p.book.sheets.map((s, i) => (
                  <option key={i} value={i}>
                    {s.name}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>引用位置</span>
              <span className="names-ref">
                <input aria-label="引用位置" className="input mono" value={f} onChange={(e) => setF(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && save()} />
                <button className="btn" title="取当前选中的区域" onClick={() => setF(p.selectionRef)}>
                  取选区
                </button>
              </span>
            </label>
            {problem && <div className="ed-warn">{problem}</div>}
            <div className="sheet-detail-foot">
              {editing && nameTarget(p.book, editing.name, p.si) && (
                <button className="btn" onClick={() => p.onGoto(editing)}>
                  转到
                </button>
              )}
              <span className="grow" />
              <button className="btn" onClick={() => (setShown(names.length === 0), setEditing(null))}>
                取消
              </button>
              <button className="btn primary" disabled={!!problem || !name} onClick={save}>
                {editing ? '修改' : '添加'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/**
 * Word 编辑器的几个对话框：超链接、书签、水印、样式、模板与批量生成。
 * 只管界面和输入检查；对文档的修改、保存都由 WordEditor 通过回调完成。
 */
import { useMemo, useState } from 'react';
import { basename, dirname, host } from '@wz/editor-shared';
import { readXlsx } from '@wz/editor-shared/xlsx';
import { checkBookmarkName, linkLabel, normalizeUrl, type BookmarkInfo, type LinkTarget } from '../core/links';
import { outputName, parseCsv, tableFromJson, tableFromRows, type DataTable, type PlaceholderInfo } from '../core/merge';
import { styleLabel, type StyleInfo } from '../core/styles';
import { isPictureWatermark, type PictureWatermark, type Watermark } from '../core/model';
import { loadImages, registerImage } from '../core/images';
import { needsPassword, protectLabel, type DocProtection, type ProtectMode } from '../core/protect';

export function Dialog({ title, width = 460, onClose, children, foot }: { title: string; width?: number; onClose: () => void; children: React.ReactNode; foot: React.ReactNode }) {
  return (
    <div className="backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div
        className="dialog"
        role="dialog"
        aria-modal="true"
        aria-label={title}
        style={{ width }}
        onKeyDown={(e) => {
          e.stopPropagation();
          if (e.key === 'Escape') onClose();
        }}
      >
        <h3>{title}</h3>
        {children}
        <div className="dialog-foot">{foot}</div>
      </div>
    </div>
  );
}

// ---------------- 超链接 ----------------

export interface LinkDialogProps {
  /** 选中的文字（没有选中时为空，要自己填显示文字） */
  text: string;
  /** 已有的链接（编辑时） */
  link?: LinkTarget;
  bookmarks: string[];
  onOk: (r: { text: string; link: LinkTarget }) => void;
  onRemove?: () => void;
  onClose: () => void;
}

export function LinkDialog({ text, link, bookmarks, onOk, onRemove, onClose }: LinkDialogProps) {
  const [kind, setKind] = useState<'url' | 'anchor'>(link?.anchor ? 'anchor' : 'url');
  const [url, setUrl] = useState(link?.url ?? '');
  const [anchor, setAnchor] = useState(link?.anchor ?? bookmarks[0] ?? '');
  const [shown, setShown] = useState(text);
  const readOnlyRid = !!link?.rid && !!link.url && url === link.url;
  const submit = () => {
    if (kind === 'url') {
      const u = normalizeUrl(url);
      if (!u) return host().toast('请输入链接地址');
      onOk({ text: shown || u, link: readOnlyRid ? { rid: link!.rid, url: link!.url } : { url: u } });
    } else {
      if (!anchor) return host().toast('文档里还没有书签：先到“书签”里添加');
      onOk({ text: shown || anchor, link: { anchor } });
    }
  };
  return (
    <Dialog
      title={link ? '编辑超链接' : '插入超链接'}
      onClose={onClose}
      foot={
        <>
          {onRemove && (
            <button className="btn" style={{ marginRight: 'auto' }} onClick={onRemove}>
              取消链接
            </button>
          )}
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" onClick={submit}>
            确定
          </button>
        </>
      }
    >
      <div style={{ display: 'grid', gap: 10 }}>
        <label className="label">
          要显示的文字
          <input className="field" value={shown} placeholder={text ? '' : '（可不填，默认显示链接地址）'} onChange={(e) => setShown(e.target.value)} autoFocus={!text} />
        </label>
        <label className="label">
          链接到
          <select className="field" value={kind} onChange={(e) => setKind(e.target.value as 'url' | 'anchor')}>
            <option value="url">网页、邮箱或文件地址</option>
            <option value="anchor">本文档中的书签</option>
          </select>
        </label>
        {kind === 'url' ? (
          <label className="label">
            地址
            <input
              className="field"
              value={url}
              placeholder="https://example.com 或 name@example.com"
              onChange={(e) => setUrl(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && submit()}
              autoFocus={!!text}
            />
          </label>
        ) : (
          <label className="label">
            书签
            <select className="field" value={anchor} onChange={(e) => setAnchor(e.target.value)}>
              {!bookmarks.length && <option value="">（没有书签）</option>}
              {anchor && !bookmarks.includes(anchor) && <option value={anchor}>{anchor}（已不存在）</option>}
              {bookmarks.map((b) => (
                <option key={b}>{b}</option>
              ))}
            </select>
          </label>
        )}
        {link && <small style={{ color: 'var(--text-dim, #888)' }}>现在：{linkLabel(link)}</small>}
      </div>
    </Dialog>
  );
}

// ---------------- 书签 ----------------

export function BookmarkDialog({
  bookmarks,
  canAdd,
  suggest,
  onAdd,
  onGo,
  onDelete,
  onClose,
}: {
  bookmarks: BookmarkInfo[];
  /** 光标在正文里才能添加 */
  canAdd: boolean;
  suggest: string;
  onAdd: (name: string) => void;
  onGo: (name: string) => void;
  onDelete: (name: string) => void;
  onClose: () => void;
}) {
  const [name, setName] = useState(suggest);
  const bad = name ? checkBookmarkName(name) : null;
  const exists = bookmarks.some((b) => b.name === name);
  return (
    <Dialog
      title="书签"
      onClose={onClose}
      foot={
        <button className="btn primary" onClick={onClose}>
          关闭
        </button>
      }
    >
      <div style={{ display: 'grid', gap: 10 }}>
        <label className="label">
          书签名
          <div style={{ display: 'flex', gap: 8 }}>
            <input className="field" style={{ flex: 1 }} value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && !bad && name && canAdd && onAdd(name)} placeholder="如：第一章、付款条款" autoFocus />
            <button className="btn primary" disabled={!canAdd || !name || !!bad} title={canAdd ? '' : '请先把光标放在正文里'} onClick={() => onAdd(name)}>
              {exists ? '移到这里' : '添加'}
            </button>
          </div>
          {bad && <small style={{ color: 'var(--danger, #d33)' }}>{bad}</small>}
          {!bad && exists && <small>已有同名书签，点“移到这里”会把它移到当前选中的位置</small>}
        </label>
        <div style={{ maxHeight: 220, overflow: 'auto', border: '1px solid var(--border, #ccc)', borderRadius: 6 }}>
          {!bookmarks.length && <div style={{ padding: 12, color: 'var(--text-dim, #888)' }}>文档里还没有书签</div>}
          {bookmarks.map((b) => (
            <div key={b.name} style={{ display: 'flex', gap: 8, alignItems: 'center', padding: '6px 10px', borderBottom: '1px solid var(--border, #eee)' }}>
              <button className="link" style={{ flex: 1, textAlign: 'left' }} onClick={() => onGo(b.name)} title="跳到书签">
                {b.name}
                {b.text && <span style={{ color: 'var(--text-dim, #888)' }}>　{b.text}</span>}
              </button>
              <button className="link" onClick={() => onDelete(b.name)}>
                删除
              </button>
            </div>
          ))}
        </div>
      </div>
    </Dialog>
  );
}

// ---------------- 水印 ----------------

const WM_PRESETS = ['机密', '绝密', '草稿', '样品', '内部资料', '请勿外传', '仅供参考'];
const WM_COLORS: [string, string][] = [
  ['银灰', 'C0C0C0'],
  ['红', 'FF0000'],
  ['蓝', '0070C0'],
  ['黑', '000000'],
];

export function WatermarkDialog({ value, onOk, onClose }: { value?: Watermark; onOk: (w: Watermark | null) => void; onClose: () => void }) {
  const pic = value && isPictureWatermark(value) ? value : undefined;
  const txt = value && !isPictureWatermark(value) ? value : undefined;
  const [kind, setKind] = useState<'text' | 'image'>(pic ? 'image' : 'text');
  const [text, setText] = useState(txt?.text ?? '机密');
  const [color, setColor] = useState(txt?.colorHex ?? 'C0C0C0');
  const [opacity, setOpacity] = useState(Math.round((txt?.opacity ?? 0.5) * 100));
  const [diagonal, setDiagonal] = useState(txt?.diagonal !== false);
  const [bold, setBold] = useState(!!txt?.bold);
  const [font, setFont] = useState(txt?.fontFamily ?? '等线');
  // 图片水印
  const [image, setImage] = useState<PictureWatermark['image'] | undefined>(pic?.image);
  const [fit, setFit] = useState(pic ? pic.scale === undefined : true);
  const [scale, setScale] = useState(pic?.scale ?? 100);
  const [washout, setWashout] = useState(pic?.washout !== false);
  const [imageName, setImageName] = useState(pic ? '（文档里原有的图片）' : '');

  const pickImage = (file?: File) => {
    if (!file) return;
    void (async () => {
      if (!/^image\/(png|jpeg|gif)$/.test(file.type)) return host().toast('水印图片只支持 png、jpg、gif');
      const data = new Uint8Array(await file.arrayBuffer());
      let widthPx = 400;
      let heightPx = 300;
      try {
        const bmp = await createImageBitmap(file);
        widthPx = bmp.width;
        heightPx = bmp.height;
        bmp.close();
      } catch {
        /* 解码不了就按默认大小 */
      }
      const id = registerImage(data, file.type);
      await loadImages([id]);
      setImage({ id, widthPx, heightPx });
      setImageName(file.name);
    })();
  };
  const ok = () => {
    if (kind === 'text') onOk({ text: text.trim(), colorHex: color, opacity: opacity / 100, diagonal, ...(bold ? { bold: true } : {}), fontFamily: font });
    else if (image) onOk({ image, ...(fit ? {} : { scale }), washout });
  };
  return (
    <Dialog
      title="水印"
      onClose={onClose}
      foot={
        <>
          {value && (
            <button className="btn" style={{ marginRight: 'auto' }} onClick={() => onOk(null)}>
              删除水印
            </button>
          )}
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" disabled={kind === 'text' ? !text.trim() : !image} onClick={ok}>
            确定
          </button>
        </>
      }
    >
      <div style={{ display: 'grid', gap: 10 }}>
        <div style={{ display: 'flex', gap: 8 }} role="tablist">
          {(['text', 'image'] as const).map((k) => (
            <button key={k} role="tab" aria-selected={kind === k} className={`btn${kind === k ? ' soft' : ''}`} onClick={() => setKind(k)}>
              {k === 'text' ? '文字水印' : '图片水印'}
            </button>
          ))}
        </div>
        {kind === 'text' ? (
          <>
            <label className="label">
              文字
              <input className="field" list="wm-presets" value={text} onChange={(e) => setText(e.target.value)} autoFocus />
              <datalist id="wm-presets">
                {WM_PRESETS.map((p) => (
                  <option key={p} value={p} />
                ))}
              </datalist>
            </label>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <label className="label">
                字体
                <select className="field" value={font} onChange={(e) => setFont(e.target.value)}>
                  {['等线', '微软雅黑', '黑体', '宋体', '楷体'].map((f) => (
                    <option key={f}>{f}</option>
                  ))}
                </select>
              </label>
              <label className="label">
                颜色
                <select className="field" value={WM_COLORS.some(([, c]) => c === color) ? color : ''} onChange={(e) => e.target.value && setColor(e.target.value)}>
                  {!WM_COLORS.some(([, c]) => c === color) && <option value="">#{color}</option>}
                  {WM_COLORS.map(([n, c]) => (
                    <option key={c} value={c}>
                      {n}
                    </option>
                  ))}
                </select>
              </label>
              <label className="label">
                版式
                <select className="field" value={diagonal ? 'd' : 'h'} onChange={(e) => setDiagonal(e.target.value === 'd')}>
                  <option value="d">斜式</option>
                  <option value="h">水平</option>
                </select>
              </label>
              <label className="label">
                浓度 {opacity}%
                <input type="range" min={10} max={100} step={5} value={opacity} onChange={(e) => setOpacity(Number(e.target.value))} />
              </label>
            </div>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="checkbox" checked={bold} onChange={(e) => setBold(e.target.checked)} />
              加粗
            </label>
          </>
        ) : (
          <>
            <div className="label">
              图片
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <label className="btn" style={{ cursor: 'pointer' }}>
                  选择图片…
                  <input
                    type="file"
                    hidden
                    accept="image/png,image/jpeg,image/gif"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      e.target.value = '';
                      pickImage(f);
                    }}
                  />
                </label>
                <span style={{ color: 'var(--text-dim, #888)' }}>{imageName || '公司标识、印章样式等（png / jpg / gif）'}</span>
              </div>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <label className="label">
                大小
                <select className="field" value={fit ? 'fit' : 'scale'} onChange={(e) => setFit(e.target.value === 'fit')}>
                  <option value="fit">适合页面</option>
                  <option value="scale">按比例</option>
                </select>
              </label>
              <label className="label">
                比例 %
                <input className="field" type="number" min={10} max={500} step={10} value={scale} disabled={fit} onChange={(e) => setScale(Number(e.target.value) || 100)} />
              </label>
            </div>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="checkbox" checked={washout} onChange={(e) => setWashout(e.target.checked)} />
              冲蚀（淡化，不影响阅读正文）
            </label>
          </>
        )}
        <small style={{ color: 'var(--text-dim, #888)' }}>水印放在页眉里：默认页眉、首页不同时的首页页眉、奇偶页不同时的偶数页页眉都会加，所有页都显示。</small>
      </div>
    </Dialog>
  );
}

// ---------------- 样式 ----------------

export function StylesDialog({
  styles,
  currentPara,
  currentChar,
  hasSelection,
  onApply,
  onCreate,
  onRedefine,
  onRename,
  onDelete,
  onClose,
}: {
  styles: StyleInfo[];
  currentPara?: string;
  currentChar?: string;
  /** 选中了文字（字符样式只能用在选中的文字上） */
  hasSelection: boolean;
  onApply: (id: string, type: 'paragraph' | 'character') => void;
  onCreate: (name: string, type: 'paragraph' | 'character') => string | null;
  onRedefine: (id: string) => void;
  /** 返回错误提示，成功返回 null */
  onRename: (id: string, name: string) => string | null;
  onDelete: (id: string) => string | null;
  onClose: () => void;
}) {
  const [tab, setTab] = useState<'paragraph' | 'character'>('paragraph');
  const [all, setAll] = useState(false);
  const [name, setName] = useState('');
  const shown = useMemo(() => styles.filter((s) => s.type === tab && (all || s.quick || s.custom || s.id === currentPara || s.id === currentChar)), [styles, tab, all, currentPara, currentChar]);
  const cur = tab === 'paragraph' ? currentPara : currentChar;
  const [pick, setPick] = useState<string | null>(null);
  const picked = shown.find((s) => s.id === (pick ?? cur)) ?? null;
  const create = () => {
    const err = onCreate(name, tab);
    if (err) host().toast(err);
    else setName('');
  };
  return (
    <Dialog
      title="样式"
      width={520}
      onClose={onClose}
      foot={
        <button className="btn primary" onClick={onClose}>
          关闭
        </button>
      }
    >
      <div style={{ display: 'grid', gap: 10 }}>
        <div style={{ display: 'flex', gap: 8 }} role="tablist">
          {(['paragraph', 'character'] as const).map((t) => (
            <button key={t} role="tab" aria-selected={tab === t} className={`btn${tab === t ? ' soft' : ''}`} onClick={() => (setTab(t), setPick(null))}>
              {t === 'paragraph' ? '段落样式' : '字符样式'}
            </button>
          ))}
          <label style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 6 }}>
            <input type="checkbox" checked={all} onChange={(e) => setAll(e.target.checked)} />
            显示全部样式
          </label>
        </div>
        <div style={{ maxHeight: 240, overflow: 'auto', border: '1px solid var(--border, #ccc)', borderRadius: 6 }}>
          {!shown.length && <div style={{ padding: 12, color: 'var(--text-dim, #888)' }}>还没有{tab === 'paragraph' ? '段落' : '字符'}样式：用下面的“新建”把选中内容的格式记成样式（标题样式在“常用”页的样式框里选）</div>}
          {shown.map((s) => (
            <div
              key={s.id}
              onClick={() => setPick(s.id)}
              onDoubleClick={() => onApply(s.id, s.type)}
              style={{ padding: '6px 10px', cursor: 'pointer', display: 'flex', gap: 8, background: picked?.id === s.id ? 'var(--accent-soft, rgba(64,128,255,.15))' : undefined, borderBottom: '1px solid var(--border, #eee)' }}
            >
              <span style={{ flex: 1 }}>{styleLabel(s)}</span>
              {s.custom && <small style={{ color: 'var(--text-dim, #888)' }}>自定义</small>}
              {s.id === cur && <small>当前</small>}
            </div>
          ))}
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button className="btn primary" disabled={!picked || (tab === 'character' && !hasSelection)} title={tab === 'character' && !hasSelection ? '先选中要套用的文字' : '双击样式也可以套用'} onClick={() => picked && onApply(picked.id, picked.type)}>
            套用
          </button>
          <button className="btn" disabled={!picked} title="把选中位置的格式记为这个样式，用到它的段落一起变" onClick={async () => picked && (await host().confirm(`用当前选中位置的格式更新样式“${styleLabel(picked)}”？\n\n所有用到这个样式的${tab === 'paragraph' ? '段落' : '文字'}都会跟着变（各自单独设过的格式保留）。`)) && onRedefine(picked.id)}>
            用选中内容更新样式
          </button>
          {picked?.custom && (
            <>
              <button
                className="btn"
                onClick={() =>
                  void host()
                    .askText('样式新名称', styleLabel(picked))
                    .then((n) => {
                      if (n === null) return;
                      const err = onRename(picked.id, n);
                      if (err) host().toast(err);
                    })
                }
              >
                重命名
              </button>
              <button
                className="btn"
                title="用到它的段落改用它基于的样式"
                onClick={async () => {
                  if (!(await host().confirm(`删除样式“${styleLabel(picked)}”？\n\n用到它的${tab === 'paragraph' ? '段落' : '文字'}会改用它基于的样式（各自单独设过的格式保留）。`))) return;
                  const err = onDelete(picked.id);
                  if (err) host().toast(err);
                  else setPick(null);
                }}
              >
                删除
              </button>
            </>
          )}
        </div>
        <div className="label">
          用选中内容新建{tab === 'paragraph' ? '段落' : '字符'}样式
          <div style={{ display: 'flex', gap: 8 }}>
            <input className="field" style={{ flex: 1 }} value={name} placeholder="样式名，如：公文正文" onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && name.trim() && create()} />
            <button className="btn" disabled={!name.trim() || (tab === 'character' && !hasSelection)} onClick={create}>
              新建
            </button>
          </div>
          <small style={{ color: 'var(--text-dim, #888)' }}>
            {tab === 'paragraph' ? '以光标所在段落的格式为准，新样式立即用在选中的段落上。' : '以选中文字的格式（与段落样式不同的部分）为准。'}
          </small>
        </div>
      </div>
    </Dialog>
  );
}

// ---------------- 限制编辑 ----------------

export function ProtectDialog({
  protection,
  hasSelection,
  permCount,
  onStart,
  onStop,
  onAddPerm,
  onClearPerms,
  onClose,
}: {
  /** 现在生效的保护（没有为 undefined） */
  protection?: DocProtection;
  /** 选中了文字（才能设成允许编辑的区域） */
  hasSelection: boolean;
  permCount: number;
  /** 返回错误提示，成功返回 null */
  onStart: (mode: ProtectMode, password: string) => Promise<string | null>;
  onStop: (password: string) => Promise<string | null>;
  onAddPerm: () => void;
  onClearPerms: () => void;
  onClose: () => void;
}) {
  const [mode, setMode] = useState<ProtectMode>('readOnly');
  const [pw, setPw] = useState('');
  const [pw2, setPw2] = useState('');
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState('');
  const run = async (fn: () => Promise<string | null>) => {
    setBusy(true);
    setErr('');
    const e = await fn();
    setBusy(false);
    if (e) setErr(e);
    else onClose();
  };
  const on = !!protection?.enforced;
  return (
    <Dialog
      title="限制编辑"
      onClose={() => !busy && onClose()}
      foot={
        <>
          <button className="btn" disabled={busy} onClick={onClose}>
            关闭
          </button>
          {on ? (
            <button className="btn primary" disabled={busy} onClick={() => void run(() => onStop(pw))}>
              {busy ? '正在检查…' : '停止保护'}
            </button>
          ) : (
            <button
              className="btn primary"
              disabled={busy}
              onClick={() => {
                if (pw !== pw2) return setErr('两次输入的密码不一样');
                void run(() => onStart(mode, pw));
              }}
            >
              {busy ? '正在设置…' : '启动保护'}
            </button>
          )}
        </>
      }
    >
      <div style={{ display: 'grid', gap: 10 }}>
        {on ? (
          <>
            <div>
              这份文档现在受保护：<b>{protectLabel(protection!)}</b>
              {permCount > 0 && `，有 ${permCount} 处允许编辑的区域`}。
            </div>
            {needsPassword(protection!) && (
              <label className="label">
                密码
                <input className="field" type="password" value={pw} autoFocus onChange={(e) => setPw(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && void run(() => onStop(pw))} />
              </label>
            )}
          </>
        ) : (
          <>
            <label className="label">
              限制方式
              <select className="field" value={mode} onChange={(e) => setMode(e.target.value as ProtectMode)}>
                <option value="readOnly">只读：不能修改（允许编辑的区域除外）</option>
                <option value="comments">只能批注：可以加批注，不能改正文</option>
                <option value="trackedChanges">只能修订：所有修改都记为修订，不能关闭或接受</option>
              </select>
            </label>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <label className="label">
                密码（可不设）
                <input className="field" type="password" value={pw} onChange={(e) => setPw(e.target.value)} />
              </label>
              <label className="label">
                再输入一遍
                <input className="field" type="password" value={pw2} onChange={(e) => setPw2(e.target.value)} />
              </label>
            </div>
            <div className="label">
              允许编辑的区域（只读 / 只能批注时仍可以改）
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <button className="btn" disabled={!hasSelection} title={hasSelection ? '' : '先在正文里选中文字'} onClick={onAddPerm}>
                  把选中内容设为允许编辑
                </button>
                <button className="btn" disabled={!permCount} onClick={onClearPerms}>
                  清除全部（{permCount}）
                </button>
              </div>
              <small style={{ color: 'var(--text-dim, #888)' }}>适合合同模板：只放开要填的地方。Word 里也认这些区域。</small>
            </div>
            <small style={{ color: 'var(--text-dim, #888)' }}>设了密码，停止保护时要输入；密码按 Word 的方式加密保存，忘记了没法找回。</small>
          </>
        )}
        {err && <div style={{ color: 'var(--danger, #d33)' }}>{err}</div>}
      </div>
    </Dialog>
  );
}

// ---------------- 模板与批量生成 ----------------

/** 读数据文件：csv / txt / json / xlsx */
async function loadTable(file: File): Promise<DataTable> {
  const name = file.name.toLowerCase();
  if (name.endsWith('.xlsx') || name.endsWith('.xlsm')) {
    const sheets = readXlsx(new Uint8Array(await file.arrayBuffer()));
    if (!sheets.length) throw new Error('工作簿里没有工作表');
    return tableFromRows(sheets[0].rows);
  }
  const buf = new Uint8Array(await file.arrayBuffer());
  // 编码：BOM / UTF-8，不是合法 UTF-8 时按 GBK（Excel 另存的 csv 常是 GBK）
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(buf);
  } catch {
    text = new TextDecoder('gbk').decode(buf);
  }
  if (name.endsWith('.json')) return tableFromJson(text);
  return tableFromRows(parseCsv(text));
}

export function MergeDialog({
  fields,
  tabPath,
  generate,
  generateAll,
  onInsertField,
  onInsertNext,
  onClose,
}: {
  /** 文档里用到的占位符 */
  fields: PlaceholderInfo[];
  tabPath?: string;
  /** 用一条数据生成填好的文档字节 */
  generate: (records: Record<string, string>[], start: number) => Promise<{ bytes: Uint8Array | null; consumed: number }>;
  /** 所有数据合并成一个文档 */
  generateAll: (records: Record<string, string>[]) => Promise<{ bytes: Uint8Array; count: number; skipped: number }>;
  onInsertField: (name: string, asWordField: boolean) => void;
  /** 在光标处插入“下一条记录”（NEXT）规则域 */
  onInsertNext: () => void;
  onClose: () => void;
}) {
  const [table, setTable] = useState<DataTable | null>(null);
  const [fileName, setFileName] = useState('');
  const [rec, setRec] = useState(0);
  const [tmpl, setTmpl] = useState(() => {
    const stem = tabPath ? basename(tabPath).replace(/\.(docx|doc|dot)$/i, '') : '文档';
    return `${stem}_{#}`;
  });
  const [wordField, setWordField] = useState(true);
  const [combined, setCombined] = useState(false);
  const [busy, setBusy] = useState<string | null>(null);
  const [newField, setNewField] = useState('');
  const values = table?.records[rec] ?? {};
  const unmatched = table ? fields.filter((f) => !table.columns.includes(f.key)) : [];
  const dir = tabPath ? dirname(tabPath) : undefined;

  const pickFile = (f?: File) => {
    if (!f) return;
    void (async () => {
      try {
        const t = await loadTable(f);
        if (!t.columns.length) throw new Error('没有读到列名（第一行应是列名）');
        setTable(t);
        setFileName(f.name);
        setRec(0);
        if (!t.records.length) host().toast('文件里只有列名，没有数据行');
      } catch (e) {
        host().toast(`读取数据文件失败：${e instanceof Error ? e.message : String(e)}`);
      }
    })();
  };

  const run = async () => {
    if (!table || !table.records.length) return;
    if (!dir) return host().toast('请先保存这份模板（生成的文件放在模板所在的文件夹）');
    if (unmatched.length && !(await host().confirm(`模板里有 ${unmatched.length} 个字段在数据里找不到：${unmatched.map((u) => u.key).join('、')}。\n\n这些字段会原样保留在生成的文档里。继续吗？`))) return;
    const s = host();
    const existing = new Set(((await s.guard(s.listDir(dir))) ?? []).map((e) => e.name.toLowerCase().replace(/\.docx$/, '')));
    const sep = dir.includes('\\') ? '\\' : '/';
    const records = table.records;
    let done = 0;
    let skipped = 0;
    try {
      if (combined) {
        setBusy('正在合并…');
        const r = await generateAll(records);
        const name = outputName(tmpl, records[0] ?? {}, 0, existing);
        if ((await s.guard(s.writeBytes(`${dir}${sep}${name}.docx`, r.bytes).then(() => true))) === undefined) return;
        done = r.count;
        skipped = r.skipped;
        s.toast(`已把 ${done} 条数据合并成一个文档：${name}.docx${skipped ? `（按条件跳过 ${skipped} 条）` : ''}`, 5000);
        onClose();
        return;
      }
      // 每条数据一个文件；模板里有 NEXT 时一份文档会用掉好几条
      let files = 0;
      for (let i = 0; i < records.length; ) {
        setBusy(`正在生成 ${i + 1} / ${records.length}…`);
        const r = await generate(records, i);
        if (r.bytes) {
          const name = outputName(tmpl, records[i], i, existing);
          if ((await s.guard(s.writeBytes(`${dir}${sep}${name}.docx`, r.bytes).then(() => true))) === undefined) return;
          files++;
        } else skipped++;
        i += Math.max(1, r.consumed);
      }
      done = files;
    } finally {
      setBusy(null);
    }
    s.toast(`已生成 ${done} 份文档，放在 ${basename(dir)} 文件夹里${skipped ? `（按条件跳过 ${skipped} 条）` : ''}`, 5000);
    onClose();
  };

  return (
    <Dialog
      title="模板与批量生成"
      width={640}
      onClose={() => !busy && onClose()}
      foot={
        <>
          <button className="btn" disabled={!!busy} onClick={onClose}>
            关闭
          </button>
          <button className="btn primary" disabled={!!busy || !table?.records.length} onClick={() => void run()}>
            {busy ?? (combined ? `合并 ${table?.records.length ?? 0} 条数据` : `生成 ${table?.records.length ?? 0} 份文档`)}
          </button>
        </>
      }
    >
      <div style={{ display: 'grid', gap: 12 }}>
        <div>
          <b>1. 在模板里放字段</b>
          <div style={{ color: 'var(--text-dim, #888)', margin: '4px 0 6px' }}>
            文字写成 <code>{'{{字段名}}'}</code>，或插入 Word 的邮件合并域；可以放在正文、表格、页眉页脚里。
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
            {table ? (
              <select className="field" style={{ width: 200 }} value="" onChange={(e) => e.target.value && onInsertField(e.target.value, wordField)}>
                <option value="">从数据列插入字段…</option>
                {table.columns.map((c) => (
                  <option key={c}>{c}</option>
                ))}
              </select>
            ) : null}
            <input className="field" style={{ width: 160 }} value={newField} placeholder="或输入字段名" onChange={(e) => setNewField(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && newField.trim() && (onInsertField(newField.trim(), wordField), setNewField(''))} />
            <button className="btn" disabled={!newField.trim()} onClick={() => (onInsertField(newField.trim(), wordField), setNewField(''))}>
              插入
            </button>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="checkbox" checked={wordField} onChange={(e) => setWordField(e.target.checked)} />
              插入为 Word 邮件合并域
            </label>
            <button className="btn" title="一份文档里放多条数据（标签、名录）：在这里换下一条数据" onClick={onInsertNext}>
              插入“下一条记录”
            </button>
          </div>
          <div style={{ marginTop: 8 }}>
            文档里现有字段：
            {fields.length ? fields.map((f) => <span key={f.key} className="tag" style={{ marginRight: 6 }}>{`${f.key}${f.count > 1 ? ` ×${f.count}` : ''}`}</span>) : <span style={{ color: 'var(--text-dim, #888)' }}>（还没有）</span>}
          </div>
        </div>
        <div>
          <b>2. 选数据文件</b>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 6 }}>
            <label className="btn" style={{ cursor: 'pointer' }}>
              选择文件…
              <input
                type="file"
                hidden
                accept=".xlsx,.xlsm,.csv,.txt,.json"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  e.target.value = '';
                  pickFile(f);
                }}
              />
            </label>
            <span style={{ color: 'var(--text-dim, #888)' }}>
              {table ? `${fileName}：${table.records.length} 条数据，${table.columns.length} 列` : 'Excel、csv 或 json；第一行是列名，每行一条数据'}
            </span>
          </div>
        </div>
        {table && table.records.length > 0 && (
          <div>
            <b>3. 核对</b>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center', margin: '6px 0' }}>
              <button className="btn" disabled={rec <= 0} onClick={() => setRec(rec - 1)}>
                上一条
              </button>
              <span>
                第 {rec + 1} / {table.records.length} 条
              </span>
              <button className="btn" disabled={rec >= table.records.length - 1} onClick={() => setRec(rec + 1)}>
                下一条
              </button>
            </div>
            <div style={{ maxHeight: 150, overflow: 'auto', border: '1px solid var(--border, #ccc)', borderRadius: 6 }}>
              {fields.map((f) => (
                <div key={f.key} style={{ display: 'flex', gap: 10, padding: '4px 10px', borderBottom: '1px solid var(--border, #eee)' }}>
                  <span style={{ width: 150, color: table.columns.includes(f.key) ? undefined : 'var(--danger, #d33)' }}>{f.key}</span>
                  <span style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{table.columns.includes(f.key) ? values[f.key] || '（空）' : '数据里没有这一列'}</span>
                </div>
              ))}
              {!fields.length && <div style={{ padding: 10, color: 'var(--text-dim, #888)' }}>模板里还没有字段</div>}
            </div>
            <div style={{ display: 'flex', gap: 18, margin: '10px 0 0' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <input type="radio" checked={!combined} onChange={() => setCombined(false)} />
                每条数据一个文件
              </label>
              <label style={{ display: 'flex', alignItems: 'center', gap: 6 }} title="每条数据从新的一页开始；模板只能有一节">
                <input type="radio" checked={combined} onChange={() => setCombined(true)} />
                合并成一个文档
              </label>
            </div>
            <label className="label" style={{ marginTop: 10 }}>
              {combined ? '文件名（{{列名}} 取第一条数据的值）' : '生成的文件名（{{列名}} 取这条数据的值，{#} 是序号）'}
              <input className="field" value={tmpl} onChange={(e) => setTmpl(e.target.value)} />
            </label>
            <small style={{ color: 'var(--text-dim, #888)' }}>
              例：{outputName(tmpl, values, rec, new Set())}.docx，保存在{dir ? ` ${basename(dir)} 文件夹` : '模板所在文件夹（模板要先保存）'}，不会覆盖已有文件。
            </small>
          </div>
        )}
      </div>
    </Dialog>
  );
}

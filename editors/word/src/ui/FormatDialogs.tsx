/**
 * 字体、段落、边框对话框，以及工具条上可以直接输入的字号框。
 * 换算和检查在 core/format.ts（有测试）；这里只管界面。对话框只把改过的项交给调用方，没动的项保持原值。
 */
import { useState } from 'react';
import { Dialog } from './WordDialogs';
import { checkParaForm, lineAtFor, paraFormOf, paraPatchOf, parseFontSize, parseLength, type LineChoice, type ParaForm } from '../core/format';
import { CN_SIZES, type Align, type BorderStyle, type CellEdge, type ParaProps, type RunProps } from '../core/model';
import { EDITOR_FONTS_EA, EDITOR_FONTS_LATIN } from '../core/render';
import type { BorderTarget } from '../core/ops';

const SIZE_NAMES = [...CN_SIZES.map(([n]) => n), '8', '9', '10', '10.5', '11', '12', '14', '16', '18', '20', '22', '24', '26', '28', '36', '48', '72'];

/** 字号的显示：正好是中文字号的用中文名 */
export const sizeName = (size: number) => CN_SIZES.find(([, v]) => v === size)?.[0] ?? String(size);

/** 工具条上的字号框：可以从列表选，也可以直接输入（“小四”“13.5”），回车或离开时生效 */
export function FontSizeBox({ size, onPick }: { size: number; onPick: (size: number) => void }) {
  const [text, setText] = useState<string | null>(null);
  const shown = text ?? sizeName(size);
  const commit = (t: string) => {
    setText(null);
    const v = parseFontSize(t);
    if (v !== null && v !== size) onPick(v);
  };
  return (
    <span className="word-size-box">
      <input
        className="ed-select"
        style={{ width: 70 }}
        aria-label="字号"
        title="字号：可以直接输入，如 小四、13.5（1–1638 磅）"
        list="word-size-list"
        value={shown}
        onFocus={(e) => e.currentTarget.select()}
        onChange={(e) => {
          const v = e.target.value;
          setText(v);
          // 从下拉列表里点选的直接生效
          if (SIZE_NAMES.includes(v)) commit(v);
        }}
        onKeyDown={(e) => {
          e.stopPropagation();
          if (e.key === 'Enter') commit(shown);
          if (e.key === 'Escape') setText(null);
        }}
        onBlur={() => text !== null && commit(text)}
      />
      <datalist id="word-size-list">
        {SIZE_NAMES.map((n) => (
          <option key={n} value={n} />
        ))}
      </datalist>
    </span>
  );
}

// ---------------- 字体 ----------------

type FontStyle = 'regular' | 'italic' | 'bold' | 'boldItalic';
const FONT_STYLES: [FontStyle, string][] = [
  ['regular', '常规'],
  ['italic', '倾斜'],
  ['bold', '加粗'],
  ['boldItalic', '加粗 倾斜'],
];

export function FontDialog({ run, onOk, onClose }: { run: RunProps; onOk: (patch: Partial<RunProps>) => void; onClose: () => void }) {
  const [ea, setEa] = useState(run.fontEA);
  const [latin, setLatin] = useState(run.fontLatin);
  const [style, setStyle] = useState<FontStyle>(run.bold ? (run.italic ? 'boldItalic' : 'bold') : run.italic ? 'italic' : 'regular');
  const [size, setSize] = useState(sizeName(run.size));
  const [autoColor, setAutoColor] = useState(run.color === 'auto');
  const [color, setColor] = useState(run.color === 'auto' ? '000000' : run.color);
  const [underline, setUnderline] = useState(run.underline);
  const [strike, setStrike] = useState(!!run.strike);
  const [vert, setVert] = useState<RunProps['vertAlign'] | ''>(run.vertAlign ?? '');
  const sp = run.spacing ?? 0;
  const [spacingKind, setSpacingKind] = useState<'normal' | 'expand' | 'condense'>(sp > 0 ? 'expand' : sp < 0 ? 'condense' : 'normal');
  const [spacingBy, setSpacingBy] = useState(String(Math.abs(sp) || 1));

  const sizeVal = parseFontSize(size);
  const spacingVal = parseLength(spacingBy, 'pt');
  const spacingOk = spacingKind === 'normal' || (!!spacingVal && spacingVal.unit === 'pt' && spacingVal.value >= 0 && spacingVal.value <= 1584);
  const ok = () => {
    if (sizeVal === null) return;
    const next: Partial<RunProps> = {
      fontEA: ea.trim() || run.fontEA,
      fontLatin: latin.trim() || run.fontLatin,
      bold: style === 'bold' || style === 'boldItalic',
      italic: style === 'italic' || style === 'boldItalic',
      size: sizeVal,
      color: autoColor ? 'auto' : color.toUpperCase(),
      underline,
      strike,
      vertAlign: vert || undefined,
      spacing: spacingKind === 'normal' ? undefined : (spacingKind === 'expand' ? 1 : -1) * (spacingVal?.value ?? 0) || undefined,
    };
    const patch: Partial<RunProps> = {};
    for (const [k, v] of Object.entries(next) as [keyof RunProps, unknown][]) {
      const old = k === 'strike' ? !!run.strike : run[k];
      if (v !== old) (patch as Record<string, unknown>)[k] = v;
    }
    onOk(patch);
  };
  const previewSize = Math.min(48, Math.max(9, (sizeVal ?? run.size) * (4 / 3)));
  return (
    <Dialog
      title="字体"
      width={520}
      onClose={onClose}
      foot={
        <>
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" disabled={sizeVal === null || !spacingOk} onClick={ok}>
            确定
          </button>
        </>
      }
    >
      <div className="word-form">
        <label className="label">
          中文字体
          <input className="field" list="word-font-ea" value={ea} onChange={(e) => setEa(e.target.value)} />
        </label>
        <label className="label">
          西文字体
          <input className="field" list="word-font-latin" value={latin} onChange={(e) => setLatin(e.target.value)} />
        </label>
        <datalist id="word-font-ea">
          {EDITOR_FONTS_EA.map((f) => (
            <option key={f} value={f} />
          ))}
        </datalist>
        <datalist id="word-font-latin">
          {EDITOR_FONTS_LATIN.map((f) => (
            <option key={f} value={f} />
          ))}
        </datalist>
        <label className="label">
          字形
          <select className="field" value={style} onChange={(e) => setStyle(e.target.value as FontStyle)}>
            {FONT_STYLES.map(([v, l]) => (
              <option key={v} value={v}>
                {l}
              </option>
            ))}
          </select>
        </label>
        <label className="label">
          字号
          <input className="field" list="word-size-list-dlg" value={size} aria-invalid={sizeVal === null} onChange={(e) => setSize(e.target.value)} />
          <datalist id="word-size-list-dlg">
            {SIZE_NAMES.map((n) => (
              <option key={n} value={n} />
            ))}
          </datalist>
        </label>
        <div className="label">
          字体颜色
          <span className="word-form-inline">
            <label className="word-check">
              <input type="checkbox" checked={autoColor} onChange={(e) => setAutoColor(e.target.checked)} />
              自动
            </label>
            <input type="color" aria-label="字体颜色" disabled={autoColor} value={`#${color}`} onChange={(e) => setColor(e.target.value.slice(1))} />
          </span>
        </div>
        <div className="label">
          字符间距
          <span className="word-form-inline">
            <select className="field" value={spacingKind} onChange={(e) => setSpacingKind(e.target.value as typeof spacingKind)}>
              <option value="normal">标准</option>
              <option value="expand">加宽</option>
              <option value="condense">紧缩</option>
            </select>
            <input className="field" style={{ width: 80 }} disabled={spacingKind === 'normal'} value={spacingBy} aria-invalid={!spacingOk} onChange={(e) => setSpacingBy(e.target.value)} />
            磅
          </span>
        </div>
        <div className="label" style={{ gridColumn: '1 / -1' }}>
          效果
          <span className="word-form-inline">
            <label className="word-check">
              <input type="checkbox" checked={underline} onChange={(e) => setUnderline(e.target.checked)} />
              下划线
            </label>
            <label className="word-check">
              <input type="checkbox" checked={strike} onChange={(e) => setStrike(e.target.checked)} />
              删除线
            </label>
            <label className="word-check">
              <input type="checkbox" checked={vert === 'superscript'} onChange={(e) => setVert(e.target.checked ? 'superscript' : '')} />
              上标
            </label>
            <label className="word-check">
              <input type="checkbox" checked={vert === 'subscript'} onChange={(e) => setVert(e.target.checked ? 'subscript' : '')} />
              下标
            </label>
          </span>
        </div>
      </div>
      <div className="word-preview" aria-label="预览">
        <span
          style={{
            fontFamily: `"${latin}", "${ea}"`,
            fontSize: previewSize,
            fontWeight: style === 'bold' || style === 'boldItalic' ? 700 : 400,
            fontStyle: style === 'italic' || style === 'boldItalic' ? 'italic' : 'normal',
            color: autoColor ? undefined : `#${color}`,
            textDecoration: [underline && 'underline', strike && 'line-through'].filter(Boolean).join(' ') || undefined,
            letterSpacing: spacingKind === 'normal' ? undefined : `${((spacingKind === 'expand' ? 1 : -1) * (spacingVal?.value ?? 0) * 4) / 3}px`,
            verticalAlign: vert === 'superscript' ? 'super' : vert === 'subscript' ? 'sub' : undefined,
          }}
        >
          微软卓越 AaBbCc 123
        </span>
      </div>
    </Dialog>
  );
}

// ---------------- 段落 ----------------

const ALIGNS: [Align, string][] = [
  ['left', '左对齐'],
  ['center', '居中'],
  ['right', '右对齐'],
  ['both', '两端对齐'],
  ['distribute', '分散对齐'],
];
const LINE_CHOICES: [LineChoice, string][] = [
  ['single', '单倍行距'],
  ['1.5', '1.5 倍行距'],
  ['double', '2 倍行距'],
  ['atLeast', '最小值'],
  ['exact', '固定值'],
  ['multiple', '多倍行距'],
];

/** 段落对话框：charPt = 光标所在段落一个字符多宽（缩进按字符换算） */
export function ParagraphDialog({ props, charPt, onOk, onClose }: { props: ParaProps; charPt: number; onOk: (patch: Partial<ParaProps>) => void; onClose: () => void }) {
  const [init] = useState(() => paraFormOf(props, charPt));
  const [f, setF] = useState<ParaForm>(init);
  const [tab, setTab] = useState<'indent' | 'pages'>('indent');
  const set = (patch: Partial<ParaForm>) => setF((x) => ({ ...x, ...patch }));
  const error = checkParaForm(f);
  const text = (k: 'left' | 'right' | 'specialBy' | 'before' | 'after' | 'lineAt', label: string, disabled = false, hint?: string) => (
    <label className="label">
      {label}
      <input className="field" value={f[k]} disabled={disabled} title={hint} onChange={(e) => set({ [k]: e.target.value })} />
    </label>
  );
  const check = (k: 'snapToGrid' | 'widowControl' | 'keepNext' | 'keepLines' | 'pageBreakBefore', label: string) => (
    <label className="word-check">
      <input type="checkbox" checked={f[k]} onChange={(e) => set({ [k]: e.target.checked })} />
      {label}
    </label>
  );
  return (
    <Dialog
      title="段落"
      width={540}
      onClose={onClose}
      foot={
        <>
          {error && <span className="word-form-error">{error}</span>}
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" disabled={!!error} onClick={() => onOk(paraPatchOf(f, init, charPt))}>
            确定
          </button>
        </>
      }
    >
      <div style={{ display: 'flex', gap: 8 }} role="tablist">
        {(['indent', 'pages'] as const).map((k) => (
          <button key={k} role="tab" aria-selected={tab === k} className={`btn${tab === k ? ' soft' : ''}`} onClick={() => setTab(k)}>
            {k === 'indent' ? '缩进和间距' : '换行和分页'}
          </button>
        ))}
      </div>
      {tab === 'indent' ? (
        <div className="word-form">
          <label className="label">
            对齐方式
            <select className="field" value={f.align} onChange={(e) => set({ align: e.target.value as Align })}>
              {ALIGNS.map(([v, l]) => (
                <option key={v} value={v}>
                  {l}
                </option>
              ))}
            </select>
          </label>
          <label className="label">
            大纲级别
            <select className="field" value={f.outline ?? ''} onChange={(e) => set({ outline: e.target.value === '' ? null : Number(e.target.value) })}>
              <option value="">正文文本</option>
              {Array.from({ length: 9 }, (_, i) => (
                <option key={i} value={i}>
                  {i + 1} 级
                </option>
              ))}
            </select>
          </label>
          <div className="word-form-title">缩进</div>
          {text('left', '左侧', false, '例：2 字符、0.5 厘米、12 磅')}
          {text('right', '右侧')}
          <label className="label">
            特殊格式
            <select className="field" value={f.special} onChange={(e) => set({ special: e.target.value as ParaForm['special'] })}>
              <option value="none">（无）</option>
              <option value="first">首行缩进</option>
              <option value="hanging">悬挂缩进</option>
            </select>
          </label>
          {text('specialBy', '缩进值', f.special === 'none')}
          <div className="word-form-title">间距</div>
          {text('before', '段前', false, '例：0.5 行、6 磅、自动')}
          {text('after', '段后', false, '例：0.5 行、6 磅、自动')}
          <label className="label">
            行距
            <select className="field" value={f.lineRule} onChange={(e) => set({ lineRule: e.target.value as LineChoice, lineAt: lineAtFor(e.target.value as LineChoice, f) })}>
              {LINE_CHOICES.map(([v, l]) => (
                <option key={v} value={v}>
                  {l}
                </option>
              ))}
            </select>
          </label>
          {text('lineAt', '设置值', f.lineRule === 'single' || f.lineRule === '1.5' || f.lineRule === 'double', f.lineRule === 'multiple' ? '倍数，如 1.25' : '例：20 磅')}
          <div style={{ gridColumn: '1 / -1' }}>{check('snapToGrid', '如果定义了文档网格，则对齐到网格')}</div>
        </div>
      ) : (
        <div className="word-form-checks">
          {check('widowControl', '孤行控制（段落首行或末行不单独留在一页）')}
          {check('keepNext', '与下段同页')}
          {check('keepLines', '段中不分页')}
          {check('pageBreakBefore', '段前分页')}
        </div>
      )}
    </Dialog>
  );
}

// ---------------- 表格框线 ----------------

export const BORDER_TARGETS: [BorderTarget, string][] = [
  ['all', '所有框线'],
  ['outside', '外侧框线'],
  ['inside', '内部框线'],
  ['insideH', '内部横框线'],
  ['insideV', '内部竖框线'],
  ['top', '上框线'],
  ['bottom', '下框线'],
  ['left', '左框线'],
  ['right', '右框线'],
  ['none', '无框线'],
];
const BORDER_STYLES: [BorderStyle, string][] = [
  ['single', '实线'],
  ['dashed', '虚线'],
  ['dotted', '点线'],
  ['double', '双线'],
];
const BORDER_WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

/** 框线对话框：线型、粗细、颜色（记作“笔”，工具条的框线下拉也用它）+ 应用到选中格子的哪些边 */
export function BordersDialog({ pen, onOk, onClose }: { pen: CellEdge; onOk: (target: BorderTarget, pen: CellEdge) => void; onClose: () => void }) {
  const [target, setTarget] = useState<BorderTarget>('all');
  const [style, setStyle] = useState<BorderStyle>(pen.style ?? 'single');
  const [w, setW] = useState(pen.w);
  const [color, setColor] = useState(pen.color);
  const line = { w, color, ...(style !== 'single' && { style }) };
  const css = `${style === 'double' ? Math.max(3, w * 1.33) : Math.max(1, w * 1.33)}px ${style === 'dashed' ? 'dashed' : style === 'dotted' ? 'dotted' : style === 'double' ? 'double' : 'solid'} ${color === 'auto' ? '#000' : `#${color}`}`;
  return (
    <Dialog
      title="边框"
      width={460}
      onClose={onClose}
      foot={
        <>
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" onClick={() => onOk(target, line)}>
            确定
          </button>
        </>
      }
    >
      <div className="word-form">
        <label className="label" style={{ gridColumn: '1 / -1' }}>
          应用于选中的单元格
          <select className="field" value={target} onChange={(e) => setTarget(e.target.value as BorderTarget)}>
            {BORDER_TARGETS.map(([v, l]) => (
              <option key={v} value={v}>
                {l}
              </option>
            ))}
          </select>
        </label>
        <label className="label">
          线型
          <select className="field" value={style} disabled={target === 'none'} onChange={(e) => setStyle(e.target.value as BorderStyle)}>
            {BORDER_STYLES.map(([v, l]) => (
              <option key={v} value={v}>
                {l}
              </option>
            ))}
          </select>
        </label>
        <label className="label">
          宽度
          <select className="field" value={w} disabled={target === 'none'} onChange={(e) => setW(Number(e.target.value))}>
            {BORDER_WIDTHS.map((v) => (
              <option key={v} value={v}>
                {v} 磅
              </option>
            ))}
          </select>
        </label>
        <div className="label">
          颜色
          <span className="word-form-inline">
            <label className="word-check">
              <input type="checkbox" checked={color === 'auto'} disabled={target === 'none'} onChange={(e) => setColor(e.target.checked ? 'auto' : '000000')} />
              自动
            </label>
            <input type="color" aria-label="框线颜色" disabled={target === 'none' || color === 'auto'} value={`#${color === 'auto' ? '000000' : color}`} onChange={(e) => setColor(e.target.value.slice(1).toUpperCase())} />
          </span>
        </div>
        <div className="label">
          预览
          <div className="word-border-preview" style={target === 'none' ? undefined : { borderBottom: css }} />
        </div>
      </div>
    </Dialog>
  );
}

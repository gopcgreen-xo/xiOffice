/**
 * 表格编辑器的格式工具栏（同 Excel“开始”选项卡的常用部分）：
 * 字体 / 字号 / 加粗倾斜下划线删除线 / 字体颜色 / 填充色 / 边框 / 对齐 / 自动换行 / 合并 / 数字格式 / 小数位 / 格式刷 / 清除格式。
 * 只负责界面，改动通过回调交给编辑器（在一个撤销步骤里改选区）。
 */
import { useEffect, useMemo, useState, type ReactNode } from 'react';
import { host } from '@wz/editor-shared';
import type { CellStyle, StyleChange } from '../core/styles';
import { BORDER_ITEMS, formatLabel, NUMBER_FORMATS, SHEET_FONTS_EA, SHEET_FONTS_LATIN, SHEET_SIZES, STANDARD_COLORS, themePalette, type BorderKind } from '../core/format';

export type MergeKind = 'center' | 'merge' | 'across' | 'unmerge';

interface Props {
  st: CellStyle;
  /** 活动单元格的数字格式 */
  fmt: string;
  defaultFont?: string;
  defaultSize: number;
  /** 活动单元格在合并区域里 */
  merged: boolean;
  onStyle: (ch: StyleChange, label: string) => void;
  onBorder: (kind: BorderKind, color?: string) => void;
  onMerge: (kind: MergeKind) => void;
  onNumFmt: (code: string) => void;
  onCustomFmt: () => void;
  onDecimals: (dir: 1 | -1) => void;
  /** 操作完回到表格（键盘焦点） */
  onDone: () => void;
  /** 接在最后的其它常用按钮（求和、排序和筛选等），和格式按钮排在一行 */
  extra?: ReactNode;
}

type MenuKind = 'color' | 'fill' | 'border' | 'merge';

export function FormatBar(p: Props) {
  const [menu, setMenu] = useState<{ kind: MenuKind; x: number; y: number } | null>(null);
  const [lastColor, setLastColor] = useState('#c00000');
  const [lastFill, setLastFill] = useState('#ffff00');
  const [lastBorder, setLastBorder] = useState<BorderKind>('bottom');
  const palette = useMemo(() => themePalette(), []);

  useEffect(() => {
    if (!menu) return;
    host().pushPopup();
    const close = () => setMenu(null);
    window.addEventListener('click', close);
    window.addEventListener('blur', close);
    return () => {
      host().popPopup();
      window.removeEventListener('click', close);
      window.removeEventListener('blur', close);
    };
  }, [menu]);

  const open = (kind: MenuKind) => (e: React.MouseEvent) => {
    // 不让这次点击冒泡到 window，否则刚打开的菜单立刻被“点击关闭”收掉
    e.stopPropagation();
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
    setMenu((m) => (m?.kind === kind ? null : { kind, x: r.left, y: r.bottom + 4 }));
  };
  const done = (fn: () => void) => () => {
    setMenu(null);
    fn();
    p.onDone();
  };
  const font = p.st.fontName ?? p.defaultFont ?? '等线';
  const size = p.st.fontSize ?? p.defaultSize;
  const toggle = (key: 'bold' | 'italic' | 'underline' | 'strike', label: string) => done(() => p.onStyle({ [key]: !p.st[key] }, label));
  const fmtLabel = formatLabel(p.fmt);

  const colorMenu = (kind: 'color' | 'fill') => {
    const apply = (c: string | null) => {
      if (c) (kind === 'color' ? setLastColor : setLastFill)(c);
      p.onStyle(kind === 'color' ? { color: c } : { fill: c }, kind === 'color' ? '字体颜色' : '填充颜色');
    };
    return (
      <div className="sheet-palette" onClick={(e) => e.stopPropagation()}>
        <button className="sheet-palette-wide" onClick={done(() => apply(null))}>
          {kind === 'color' ? (
            <>
              <span className="swatch" style={{ background: '#000' }} />
              自动
            </>
          ) : (
            <>
              <span className="swatch none" />
              无填充
            </>
          )}
        </button>
        <div className="sheet-palette-label">主题颜色</div>
        <div className="sheet-palette-grid">
          {palette.map((row, i) =>
            row.map((c) => <button key={`${i}${c}`} className={`swatch${i === 0 ? ' first' : ''}`} style={{ background: c }} title={c} onClick={done(() => apply(c))} />),
          )}
        </div>
        <div className="sheet-palette-label">标准色</div>
        <div className="sheet-palette-grid">
          {STANDARD_COLORS.map((c) => (
            <button key={c} className="swatch" style={{ background: c }} title={c} onClick={done(() => apply(c))} />
          ))}
        </div>
        <label className="sheet-palette-wide">
          其他颜色…
          <input
            type="color"
            defaultValue={kind === 'color' ? lastColor : lastFill}
            onChange={(e) => {
              apply(e.target.value);
              setMenu(null);
              p.onDone();
            }}
          />
        </label>
      </div>
    );
  };

  return (
    <div className="sheet-fmtbar" onMouseDown={(e) => e.target instanceof HTMLButtonElement && e.preventDefault()}>
      <select className="ed-select" style={{ width: 112 }} aria-label="字体" value={font} onChange={(e) => (p.onStyle({ fontName: e.target.value }, '字体'), p.onDone())}>
        {/* 分中文和西文两组（同 Word）；不在列表里的当前字体按名字里有没有汉字排进对应的组，等线放在中文组最后 */}
        <optgroup label="中文字体">
          {[...new Set([...SHEET_FONTS_EA.filter((f) => f !== '等线'), ...(/[一-鿿]/.test(font) && font !== '等线' ? [font] : []), '等线'])].map((f) => (
            <option key={`ea-${f}`} value={f}>
              {f}
            </option>
          ))}
        </optgroup>
        <optgroup label="西文字体">
          {[...new Set([...(/[一-鿿]/.test(font) || SHEET_FONTS_EA.includes(font) ? [] : [font]), ...SHEET_FONTS_LATIN])].map((f) => (
            <option key={`la-${f}`} value={f}>
              {f}
            </option>
          ))}
        </optgroup>
      </select>
      <select className="ed-select" style={{ width: 62 }} aria-label="字号" value={size} onChange={(e) => (p.onStyle({ fontSize: Number(e.target.value) }, '字号'), p.onDone())}>
        {[...new Set([size, ...SHEET_SIZES])].sort((a, b) => a - b).map((n) => (
          <option key={n} value={n}>
            {n}
          </option>
        ))}
      </select>
      <Btn on={p.st.bold} title="加粗（Ctrl+B）" onClick={toggle('bold', '加粗')} style={{ fontWeight: 800 }}>
        B
      </Btn>
      <Btn on={p.st.italic} title="倾斜（Ctrl+I）" onClick={toggle('italic', '倾斜')} style={{ fontStyle: 'italic', fontFamily: 'serif' }}>
        I
      </Btn>
      <Btn on={p.st.underline} title="下划线（Ctrl+U）" onClick={toggle('underline', '下划线')} style={{ textDecoration: 'underline' }}>
        U
      </Btn>
      <Btn on={p.st.strike} title="删除线（Ctrl+5）" onClick={toggle('strike', '删除线')}>
        <span className="strike-ico">abc</span>
      </Btn>
      <div className="sheet-split">
        <Btn title={`字体颜色（${lastColor}）`} onClick={done(() => p.onStyle({ color: lastColor }, '字体颜色'))}>
          <span className="sheet-color-a" style={{ borderBottomColor: lastColor }}>
            A
          </span>
        </Btn>
        <button className="icon-btn sheet-caret-btn" aria-label="字体颜色" onClick={open('color')}>
          <Caret />
        </button>
      </div>
      <div className="sheet-split">
        <Btn title={`填充颜色（${lastFill}）`} onClick={done(() => p.onStyle({ fill: lastFill }, '填充颜色'))}>
          <FillIcon color={lastFill} />
        </Btn>
        <button className="icon-btn sheet-caret-btn" aria-label="填充颜色" onClick={open('fill')}>
          <Caret />
        </button>
      </div>
      <div className="sheet-split">
        <Btn title={BORDER_ITEMS.find((x) => x.kind === lastBorder)?.label ?? '边框'} onClick={done(() => p.onBorder(lastBorder))}>
          <BorderIcon kind={lastBorder} />
        </Btn>
        <button className="icon-btn sheet-caret-btn" aria-label="边框" onClick={open('border')}>
          <Caret />
        </button>
      </div>
      <span className="sep" />
      <Btn on={p.st.align === 'left'} title="左对齐" onClick={done(() => p.onStyle({ align: p.st.align === 'left' ? null : 'left' }, '对齐'))}>
        <AlignIcon h="left" />
      </Btn>
      <Btn on={p.st.align === 'center'} title="居中" onClick={done(() => p.onStyle({ align: p.st.align === 'center' ? null : 'center' }, '对齐'))}>
        <AlignIcon h="center" />
      </Btn>
      <Btn on={p.st.align === 'right'} title="右对齐" onClick={done(() => p.onStyle({ align: p.st.align === 'right' ? null : 'right' }, '对齐'))}>
        <AlignIcon h="right" />
      </Btn>
      <Btn on={p.st.valign === 'top'} title="顶端对齐" onClick={done(() => p.onStyle({ valign: 'top' }, '对齐'))}>
        <VAlignIcon v="top" />
      </Btn>
      <Btn on={!p.st.valign || p.st.valign === 'center'} title="垂直居中" onClick={done(() => p.onStyle({ valign: 'center' }, '对齐'))}>
        <VAlignIcon v="center" />
      </Btn>
      <Btn on={p.st.valign === 'bottom'} title="底端对齐" onClick={done(() => p.onStyle({ valign: 'bottom' }, '对齐'))}>
        <VAlignIcon v="bottom" />
      </Btn>
      <Btn on={p.st.wrap} title="自动换行" onClick={done(() => p.onStyle({ wrap: !p.st.wrap }, '自动换行'))}>
        <WrapIcon />
      </Btn>
      <div className="sheet-split">
        <button className={`btn sheet-merge-btn${p.merged ? ' soft on' : ''}`} title={p.merged ? '取消合并' : '合并后居中'} onClick={done(() => p.onMerge(p.merged ? 'unmerge' : 'center'))}>
          <MergeIcon />
          {p.merged ? '取消合并' : '合并居中'}
        </button>
        <button className="icon-btn sheet-caret-btn" aria-label="合并选项" onClick={open('merge')}>
          <Caret />
        </button>
      </div>
      <span className="sep" />
      <select
        className="ed-select"
        style={{ width: 110 }}
        aria-label="数字格式"
        title="数字格式；下面还有百分比、千位分隔、增减小数位数"
        value={NUMBER_FORMATS.some((f) => f.code === p.fmt) ? p.fmt : '__cur'}
        onChange={(e) => {
          const v = e.target.value;
          if (v === '__custom') p.onCustomFmt();
          else if (v === '__pct') (p.onNumFmt('0%'), p.onDone());
          else if (v === '__thou') (p.onNumFmt('#,##0.00'), p.onDone());
          else if (v === '__inc') (p.onDecimals(1), p.onDone());
          else if (v === '__dec') (p.onDecimals(-1), p.onDone());
          else if (v !== '__cur') {
            p.onNumFmt(e.target.value);
            p.onDone();
          }
        }}
      >
        {!NUMBER_FORMATS.some((f) => f.code === p.fmt) && <option value="__cur">{fmtLabel}</option>}
        {NUMBER_FORMATS.map((f) => (
          <option key={f.code} value={f.code}>
            {f.label}
          </option>
        ))}
        <option value="__custom">其他格式…</option>
        <optgroup label="快捷操作">
          <option value="__pct">百分比样式 %</option>
          <option value="__thou">千位分隔样式 ,</option>
          <option value="__inc">增加小数位数 .0→.00</option>
          <option value="__dec">减少小数位数 .00→.0</option>
        </optgroup>
      </select>

      {menu && (
        <div className="ctx-menu sheet-dropdown" role="menu" style={{ left: menu.x, top: menu.y }} onMouseDown={(e) => e.preventDefault()}>
          {menu.kind === 'color' && colorMenu('color')}
          {menu.kind === 'fill' && colorMenu('fill')}
          {menu.kind === 'border' &&
            BORDER_ITEMS.map((b) => (
              <button
                key={b.kind}
                role="menuitem"
                onClick={done(() => {
                  setLastBorder(b.kind);
                  p.onBorder(b.kind);
                })}
              >
                <BorderIcon kind={b.kind} />
                {b.label}
              </button>
            ))}
          {menu.kind === 'merge' && (
            <>
              <button role="menuitem" onClick={done(() => p.onMerge('center'))}>
                合并后居中
              </button>
              <button role="menuitem" onClick={done(() => p.onMerge('across'))}>
                跨越合并（每行分别合并）
              </button>
              <button role="menuitem" onClick={done(() => p.onMerge('merge'))}>
                合并单元格
              </button>
              <button role="menuitem" onClick={done(() => p.onMerge('unmerge'))}>
                取消单元格合并
              </button>
            </>
          )}
        </div>
      )}
      {p.extra && (
        <>
          <span className="sep" />
          {p.extra}
        </>
      )}
    </div>
  );
}

function Btn({ on, title, onClick, style, children }: { on?: boolean; title: string; onClick: () => void; style?: React.CSSProperties; children: ReactNode }) {
  return (
    <button className={`icon-btn fmt${on ? ' on' : ''}`} aria-pressed={on} title={title} aria-label={title} onClick={onClick} style={style}>
      {children}
    </button>
  );
}

// ---------- 小图标（16×16，currentColor） ----------

const Svg = ({ children }: { children: ReactNode }) => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" aria-hidden="true">
    {children}
  </svg>
);

function Caret() {
  return (
    <svg width="12" height="12" viewBox="0 0 8 8" aria-hidden="true">
      <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.5" />
    </svg>
  );
}

function AlignIcon({ h }: { h: 'left' | 'center' | 'right' }) {
  const x = (w: number) => (h === 'left' ? 2 : h === 'right' ? 14 - w : 8 - w / 2);
  return (
    <Svg>
      {[12, 8, 12, 6].map((w, i) => (
        <path key={i} d={`M${x(w)} ${3 + i * 3.3}h${w}`} />
      ))}
    </Svg>
  );
}

function VAlignIcon({ v }: { v: 'top' | 'center' | 'bottom' }) {
  const y = v === 'top' ? 5 : v === 'center' ? 8 : 11;
  return (
    <Svg>
      <path d={v === 'top' ? 'M2 2h12' : v === 'bottom' ? 'M2 14h12' : 'M2 8h3M11 8h3'} />
      <path d={`M8 ${y - 3}v6`} />
      <path d={v === 'top' ? `M5.5 ${y - 0.5}L8 ${y - 3}l2.5 2.5` : v === 'bottom' ? `M5.5 ${y + 0.5}L8 ${y + 3}l2.5-2.5` : `M6 ${y - 3}h4M6 ${y + 3}h4`} />
    </Svg>
  );
}

function WrapIcon() {
  return (
    <Svg>
      <path d="M2 3.5h12M2 8h10a2 2 0 0 1 0 4H8.5M2 12.5h4" />
      <path d="M10 10.5L8.5 12l1.5 1.5" />
    </Svg>
  );
}

function MergeIcon() {
  return (
    <Svg>
      <rect x="1.5" y="3" width="13" height="10" rx="1" />
      <path d="M4 8h8M5.5 6.5L4 8l1.5 1.5M10.5 6.5L12 8l-1.5 1.5" />
    </Svg>
  );
}

function FillIcon({ color }: { color: string }) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <path d="M3 7.5L7.5 3l4.5 4.5L7.5 12z" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinejoin="round" />
      <path d="M12.5 9.5c.8 1.1 1.2 1.8 1.2 2.3a1.2 1.2 0 0 1-2.4 0c0-.5.4-1.2 1.2-2.3z" fill="currentColor" />
      <rect x="1" y="13.5" width="14" height="2.5" fill={color} />
    </svg>
  );
}

function BorderIcon({ kind }: { kind: BorderKind }) {
  const thick = kind === 'thickOutside' || kind === 'thickBottom' ? 2.2 : 1.4;
  const solid: Record<string, string> = {
    all: 'M2 2h12v12H2zM8 2v12M2 8h12',
    outside: 'M2 2h12v12H2z',
    thickOutside: 'M2 2h12v12H2z',
    top: 'M2 2h12',
    bottom: 'M2 14h12',
    left: 'M2 2v12',
    right: 'M14 2v12',
    doubleBottom: 'M2 12.5h12M2 14.5h12',
    thickBottom: 'M2 14h12',
    none: '',
  };
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <path d="M2 2h12v12H2zM8 2v12M2 8h12" fill="none" stroke="currentColor" strokeWidth="1" strokeDasharray="1 1.5" opacity="0.45" />
      {solid[kind] && <path d={solid[kind]} fill="none" stroke="currentColor" strokeWidth={thick} />}
    </svg>
  );
}


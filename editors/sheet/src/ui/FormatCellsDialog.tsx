/**
 * “设置单元格格式”对话框（同 Excel 的 Ctrl+1）：数字 / 对齐 / 字体 / 边框 / 填充。
 * 打开时按活动单元格的格式填好；确定时只交出改过的项（没动的项不改选区里各格子原来的格式），由编辑器在一个撤销步骤里应用。
 */
import { useEffect, useMemo, useState, type ReactNode } from 'react';
import { host } from '@wz/editor-shared';
import { Icon } from '@wz/editor-shared/Icon';
import type { CellStyle, StyleChange } from '../core/styles';
import { formatNumber } from '../core/numfmt';
import {
  BORDER_ITEMS,
  buildFormat,
  DATE_CODES,
  FMT_CATEGORIES,
  FMT_DECIMALS,
  fmtParts,
  SHEET_FONTS_EA,
  SHEET_FONTS_LATIN,
  SHEET_SIZES,
  STANDARD_COLORS,
  themePalette,
  TIME_CODES,
  type BorderKind,
  type FmtCategory,
} from '../core/format';

export interface FormatCellsResult {
  style: StyleChange;
  border?: { kind: BorderKind; color?: string };
  /** 合并（true）/ 取消合并（false）；没改为 undefined */
  merge?: boolean;
}

type TabKey = 'number' | 'align' | 'font' | 'border' | 'fill';
const TABS: [TabKey, string][] = [
  ['number', '数字'],
  ['align', '对齐'],
  ['font', '字体'],
  ['border', '边框'],
  ['fill', '填充'],
];

interface Props {
  st: CellStyle;
  /** 活动单元格的数字格式代码 */
  fmt: string;
  /** 活动单元格的值（数字格式预览用） */
  value: unknown;
  date1904?: boolean;
  defaultFont: string;
  defaultSize: number;
  /** 活动单元格在合并区域里 */
  merged: boolean;
  /** 选区不止一个格子（才能合并） */
  multi: boolean;
  onOk: (r: FormatCellsResult) => void;
  onClose: () => void;
}

export function FormatCellsDialog(p: Props) {
  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  const [tab, setTab] = useState<TabKey>('number');

  // 数字
  const init = useMemo(() => fmtParts(p.fmt), [p.fmt]);
  const [cat, setCat] = useState<FmtCategory>(init.cat);
  const [decimals, setDecimals] = useState(init.decimals);
  const [thousands, setThousands] = useState(init.thousands);
  const [code, setCode] = useState(p.fmt);
  const codes = cat === 'date' ? DATE_CODES : cat === 'time' ? TIME_CODES : null;
  const numFmt = buildFormat({ cat, decimals, thousands }, code);
  const pickCat = (c: FmtCategory) => {
    setCat(c);
    if (c === 'date') setCode(DATE_CODES.includes(code) ? code : DATE_CODES[0]);
    else if (c === 'time') setCode(TIME_CODES.includes(code) ? code : TIME_CODES[0]);
    else if (c === 'custom') setCode(buildFormat({ cat, decimals, thousands }, code));
  };
  const sample = typeof p.value === 'number' ? p.value : typeof p.value === 'string' && p.value !== '' ? p.value : cat === 'date' || cat === 'time' ? 45000.5 : 1234.5678;
  let preview = '';
  try {
    preview = typeof sample === 'number' ? formatNumber(sample, numFmt, p.date1904) : sample;
  } catch {
    preview = '（格式代码有误）';
  }

  // 对齐
  const [align, setAlign] = useState<CellStyle['align'] | ''>(p.st.align ?? '');
  const [valign, setValign] = useState<CellStyle['valign']>(p.st.valign ?? 'center');
  const [wrap, setWrap] = useState(!!p.st.wrap);
  const [merge, setMerge] = useState(p.merged);

  // 字体
  const [fontName, setFontName] = useState(p.st.fontName ?? p.defaultFont);
  const [fontSize, setFontSize] = useState(p.st.fontSize ?? p.defaultSize);
  const [bold, setBold] = useState(!!p.st.bold);
  const [italic, setItalic] = useState(!!p.st.italic);
  const [underline, setUnderline] = useState(!!p.st.underline);
  const [strike, setStrike] = useState(!!p.st.strike);
  const [color, setColor] = useState<string | null>(p.st.color ?? null);

  // 边框（选一种画法；不选 = 不改）
  const [border, setBorder] = useState<BorderKind | null>(null);
  const [borderColor, setBorderColor] = useState('#000000');

  // 填充
  const [fill, setFill] = useState<string | null>(p.st.fill ?? null);
  const palette = useMemo(() => themePalette(), []);

  const ok = () => {
    const s: StyleChange = {};
    if (numFmt !== p.fmt) s.numFmt = numFmt;
    if ((align || null) !== (p.st.align ?? null)) s.align = align || null;
    if (valign !== (p.st.valign ?? 'center')) s.valign = valign ?? null;
    if (wrap !== !!p.st.wrap) s.wrap = wrap;
    if (fontName !== (p.st.fontName ?? p.defaultFont)) s.fontName = fontName;
    if (fontSize !== (p.st.fontSize ?? p.defaultSize)) s.fontSize = fontSize;
    if (bold !== !!p.st.bold) s.bold = bold;
    if (italic !== !!p.st.italic) s.italic = italic;
    if (underline !== !!p.st.underline) s.underline = underline;
    if (strike !== !!p.st.strike) s.strike = strike;
    if (color !== (p.st.color ?? null)) s.color = color;
    if (fill !== (p.st.fill ?? null)) s.fill = fill;
    p.onOk({
      style: s,
      ...(border && { border: { kind: border, ...(border !== 'none' && { color: borderColor }) } }),
      ...(merge !== p.merged && { merge }),
    });
  };

  const fontList = [...new Set([...SHEET_FONTS_EA, ...SHEET_FONTS_LATIN, fontName])];
  const sizes = [...new Set([fontSize, ...SHEET_SIZES])].sort((a, b) => a - b);

  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div
        className="sheet-detail sheet-fmtdlg"
        role="dialog"
        aria-label="设置单元格格式"
        onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Escape') p.onClose();
          else if (e.key === 'Enter' && !(e.target instanceof HTMLSelectElement)) ok();
        }}
      >
        <div className="sheet-detail-head">
          <strong>设置单元格格式</strong>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        <div className="sheet-fmtdlg-tabs" role="tablist">
          {TABS.map(([k, label]) => (
            <button key={k} role="tab" aria-selected={tab === k} className={tab === k ? 'on' : undefined} onClick={() => setTab(k)}>
              {label}
            </button>
          ))}
        </div>

        <div className="sheet-fmtdlg-body">
          {tab === 'number' && (
            <div className="sheet-fmtdlg-num">
              <div className="sheet-fmtdlg-list" role="listbox" aria-label="分类">
                {FMT_CATEGORIES.map((c) => (
                  <button key={c.cat} role="option" aria-selected={cat === c.cat} className={cat === c.cat ? 'on' : undefined} onClick={() => pickCat(c.cat)}>
                    {c.label}
                  </button>
                ))}
              </div>
              <div className="sheet-fmtdlg-col">
                <Field label="示例">
                  <span className="sheet-fmtdlg-sample mono">{preview}</span>
                </Field>
                {FMT_DECIMALS.includes(cat) && (
                  <Field label="小数位数">
                    <input className="input" type="number" min={0} max={30} value={decimals} style={{ width: 70 }} onChange={(e) => setDecimals(Math.max(0, Math.min(30, Number(e.target.value) || 0)))} />
                  </Field>
                )}
                {cat === 'number' && (
                  <label className="sheet-fmtdlg-check">
                    <input type="checkbox" checked={thousands} onChange={(e) => setThousands(e.target.checked)} />
                    使用千位分隔符（,）
                  </label>
                )}
                {codes && (
                  <Field label="类型">
                    <select className="ed-select" value={code} onChange={(e) => setCode(e.target.value)}>
                      {codes.map((c) => (
                        <option key={c} value={c}>
                          {formatNumber(45000.5, c, p.date1904)}
                        </option>
                      ))}
                    </select>
                  </Field>
                )}
                {cat === 'custom' && (
                  <Field label="类型">
                    <input className="input mono" value={code} onChange={(e) => setCode(e.target.value.replace(/\[红色\]/g, '[Red]'))} placeholder='如 0.00"吨"、yyyy-mm-dd' />
                  </Field>
                )}
                <div className="sub">
                  {cat === 'general' && '常规格式不包含任何特定的数字格式。'}
                  {cat === 'text' && '文本格式的单元格里，数字也当作文本，显示的内容和输入的完全一样。'}
                  {cat === 'accounting' && '会计格式可对一列数值进行货币符号和小数点对齐。'}
                  {cat === 'custom' && '以现有格式为基础，输入格式代码。'}
                </div>
              </div>
            </div>
          )}

          {tab === 'align' && (
            <div className="sheet-fmtdlg-col">
              <Field label="水平对齐">
                <select className="ed-select" value={align ?? ''} onChange={(e) => setAlign(e.target.value as CellStyle['align'] | '')}>
                  <option value="">常规</option>
                  <option value="left">靠左</option>
                  <option value="center">居中</option>
                  <option value="right">靠右</option>
                </select>
              </Field>
              <Field label="垂直对齐">
                <select className="ed-select" value={valign ?? 'center'} onChange={(e) => setValign(e.target.value as CellStyle['valign'])}>
                  <option value="top">靠上</option>
                  <option value="center">居中</option>
                  <option value="bottom">靠下</option>
                </select>
              </Field>
              <div className="sub">文本控制</div>
              <label className="sheet-fmtdlg-check">
                <input type="checkbox" checked={wrap} onChange={(e) => setWrap(e.target.checked)} />
                自动换行
              </label>
              <label className="sheet-fmtdlg-check">
                <input type="checkbox" checked={merge} disabled={!p.multi && !p.merged} onChange={(e) => setMerge(e.target.checked)} />
                合并单元格
              </label>
            </div>
          )}

          {tab === 'font' && (
            <div className="sheet-fmtdlg-col">
              <Field label="字体">
                <select className="ed-select" value={fontName} onChange={(e) => setFontName(e.target.value)}>
                  {fontList.map((f) => (
                    <option key={f} value={f}>
                      {f}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label="字形">
                <select
                  className="ed-select"
                  value={`${bold ? 'b' : ''}${italic ? 'i' : ''}`}
                  onChange={(e) => (setBold(e.target.value.includes('b')), setItalic(e.target.value.includes('i')))}
                >
                  <option value="">常规</option>
                  <option value="i">倾斜</option>
                  <option value="b">加粗</option>
                  <option value="bi">加粗 倾斜</option>
                </select>
              </Field>
              <Field label="字号">
                <select className="ed-select" value={fontSize} onChange={(e) => setFontSize(Number(e.target.value))}>
                  {sizes.map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label="颜色">
                <ColorPick value={color} auto="自动" onChange={setColor} />
              </Field>
              <label className="sheet-fmtdlg-check">
                <input type="checkbox" checked={underline} onChange={(e) => setUnderline(e.target.checked)} />
                下划线
              </label>
              <label className="sheet-fmtdlg-check">
                <input type="checkbox" checked={strike} onChange={(e) => setStrike(e.target.checked)} />
                删除线
              </label>
              <Field label="预览">
                <span
                  className="sheet-fmtdlg-font"
                  style={{
                    fontFamily: fontName,
                    fontSize: `${Math.min(fontSize, 28) * (4 / 3)}px`,
                    fontWeight: bold ? 700 : 400,
                    fontStyle: italic ? 'italic' : 'normal',
                    textDecoration: [underline && 'underline', strike && 'line-through'].filter(Boolean).join(' ') || 'none',
                    color: color ?? undefined,
                  }}
                >
                  微软卓越 AaBbCc
                </span>
              </Field>
            </div>
          )}

          {tab === 'border' && (
            <div className="sheet-fmtdlg-col">
              <div className="sub">选一种边框，确定后应用到选中的单元格（不选则不改边框）</div>
              <div className="sheet-fmtdlg-borders">
                {BORDER_ITEMS.map((b) => (
                  <label key={b.kind} className="sheet-fmtdlg-check">
                    <input type="radio" name="fc-border" checked={border === b.kind} onChange={() => setBorder(b.kind)} />
                    {b.label}
                  </label>
                ))}
              </div>
              <Field label="颜色">
                <ColorPick value={borderColor} onChange={(c) => setBorderColor(c ?? '#000000')} />
              </Field>
            </div>
          )}

          {tab === 'fill' && (
            <div className="sheet-fmtdlg-col">
              <Field label="背景色">
                <ColorPick value={fill} auto="无颜色" onChange={setFill} />
              </Field>
              <div className="sheet-fmtdlg-palette">
                {[...palette, STANDARD_COLORS].map((row, i) => (
                  <div key={i} className="sheet-fmtdlg-row">
                    {row.map((c, j) => (
                      <button key={j} className={`swatch${fill === c ? ' on' : ''}`} style={{ background: c }} title={c} aria-label={c} onClick={() => setFill(c)} />
                    ))}
                  </div>
                ))}
              </div>
              <Field label="示例">
                <span className="sheet-fmtdlg-fill" style={{ background: fill ?? 'transparent' }} />
              </Field>
            </div>
          )}
        </div>

        <div className="sheet-detail-foot">
          <span className="grow" />
          <button className="btn" onClick={p.onClose}>
            取消
          </button>
          <button className="btn primary" onClick={ok}>
            确定
          </button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="sheet-fmtdlg-field">
      <span className="sub">{label}</span>
      {children}
    </div>
  );
}

/** 颜色：色块 + 系统取色；auto 有值时多一个“自动 / 无颜色”按钮（选中时 value = null） */
function ColorPick({ value, auto, onChange }: { value: string | null; auto?: string; onChange: (c: string | null) => void }) {
  return (
    <span className="sheet-fmtdlg-color">
      {auto && (
        <button className={`btn${value === null ? ' soft on' : ''}`} onClick={() => onChange(null)}>
          {auto}
        </button>
      )}
      <input type="color" aria-label="选择颜色" value={value ?? '#000000'} onChange={(e) => onChange(e.target.value)} />
    </span>
  );
}

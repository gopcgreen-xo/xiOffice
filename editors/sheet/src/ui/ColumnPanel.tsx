/**
 * 列属性浮层：设置一列的字段类型（文本 / 数字 / 金额 / 百分比 / 日期 / 单选 / 多选 / 复选框 / 网址 / 邮箱 / 电话 …）、
 * 选项与颜色、数字格式和范围、必填、默认值、输入不合规时的处理。只在工作包里的表格出现；
 * 列属性保存在工作包里，不写进 xlsx。
 */
import { useEffect, useMemo, useState } from 'react';
import {
  COLUMN_TYPE_GLYPHS,
  COLUMN_TYPE_LABELS,
  OPTION_PALETTE,
  host,
  splitMulti,
  type ColumnMeta,
  type ColumnType,
  type SelectOption,
} from '@wz/editor-shared';
import { colName } from '@wz/editor-shared/xlsx';
import { cellText, columnValues } from '../core/colmeta';
import { makeMeta } from '../core/columns';
import { parseRef, type SheetData, type Workbook } from '../core/workbook';
import './columns.css';

/** 浮层上排的类型（自动编号、创建时间放在“更多”里，一般只在表单里用） */
const CHIPS: ColumnType[] = ['text', 'number', 'money', 'percent', 'date', 'select', 'multi', 'check', 'link', 'email', 'phone', 'autonum', 'created'];

const DATE_FORMATS: { fmt: string; label: string }[] = [
  { fmt: 'yyyy-mm-dd', label: '2026-10-06' },
  { fmt: 'yyyy/m/d', label: '2026/10/6' },
  { fmt: 'yyyy"年"m"月"d"日"', label: '2026年10月6日' },
];

const INVALID: { v: NonNullable<ColumnMeta['onInvalid']>; label: string }[] = [
  { v: 'stop', label: '拒绝' },
  { v: 'warn', label: '警告' },
  { v: 'info', label: '仅提示' },
];

/** 数字格式码 ⇄ 小数位数 / 千位分隔 */
export function parseNumFmt(fmt: string | undefined): { decimals?: number; thousands: boolean } {
  if (!fmt) return { thousands: false };
  const dec = /\.(0+)/.exec(fmt);
  return { ...(dec && { decimals: dec[1].length }), thousands: fmt.includes(',') };
}

export function buildNumFmt(type: ColumnType, decimals: number | undefined, thousands: boolean): string | undefined {
  if (type === 'percent') return `0${decimals ? `.${'0'.repeat(decimals)}` : ''}%`;
  if (decimals === undefined && !thousands) return undefined;
  return `${thousands ? '#,##0' : '0'}${decimals ? `.${'0'.repeat(decimals)}` : ''}`;
}

const numText = (v: number | undefined): string => (v === undefined ? '' : String(v));

export interface ColumnPanelProps {
  book: Workbook;
  sheet: SheetData;
  col: number;
  existing: ColumnMeta | undefined;
  /** 表头所在行（1 起）：还没有列属性时由表格视图按数据区猜 */
  header: number;
  /** 浮层位置（屏幕坐标，贴着列头） */
  at: { x: number; y: number };
  onApply: (meta: ColumnMeta | null) => void;
  onClose: () => void;
}

export function ColumnPanel({ sheet, col, existing, header, at, onApply, onClose }: ColumnPanelProps) {
  const title = cellText(sheet, header - 1, col).trim();
  const [type, setType] = useState<ColumnType>(existing?.type ?? 'text');
  const [options, setOptions] = useState<SelectOption[]>(existing?.options ?? []);
  const [source, setSource] = useState<'manual' | 'range'>(existing?.optionsRef ? 'range' : 'manual');
  const [ref, setRef] = useState(existing?.optionsRef ?? '');
  const [required, setRequired] = useState(!!existing?.required);
  const [def, setDef] = useState(existing?.default ?? '');
  const [note, setNote] = useState(existing?.note ?? '');
  const [onInvalid, setOnInvalid] = useState<NonNullable<ColumnMeta['onInvalid']>>(existing?.onInvalid ?? 'stop');
  const initFmt = parseNumFmt(existing?.fmt);
  const [decimals, setDecimals] = useState<string>(
    existing?.type === 'money' || existing?.type === 'percent' ? String(initFmt.decimals ?? (existing.type === 'money' ? 2 : 0)) : numText(initFmt.decimals),
  );
  const [thousands, setThousands] = useState(existing ? initFmt.thousands : true);
  const [min, setMin] = useState(numText(existing?.min));
  const [max, setMax] = useState(numText(existing?.max));
  const [integer, setInteger] = useState(!!existing?.integer);
  const [dateFmt, setDateFmt] = useState(existing?.fmt && DATE_FORMATS.some((d) => d.fmt === existing.fmt) ? existing.fmt : DATE_FORMATS[0].fmt);
  const [paletteFor, setPaletteFor] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    document.addEventListener('keydown', onKey, true);
    return () => document.removeEventListener('keydown', onKey, true);
  }, [onClose]);

  const isChoice = type === 'select' || type === 'multi';
  const isNum = type === 'number' || type === 'money' || type === 'percent';

  // 每个选项在本列里用了几次
  const counts = useMemo(() => {
    const m = new Map<string, number>();
    for (let r = header; r < sheet.rows; r++) {
      const t = cellText(sheet, r, col).trim();
      for (const v of t ? (type === 'multi' ? splitMulti(t) : [t]) : []) m.set(v, (m.get(v) ?? 0) + 1);
    }
    return m;
  }, [sheet, col, header, type]);

  const changeType = (next: ColumnType) => {
    setType(next);
    setError(null);
    // 换成单选 / 多选且还没有选项：用本列已有的取值补上
    if ((next === 'select' || next === 'multi') && options.length === 0) {
      const found = columnValues(sheet, header, col, next === 'multi');
      if (found.length <= 30) setOptions(found.map((v) => ({ v })));
    }
    if (next === 'money' && decimals === '') setDecimals('2');
    if (next === 'percent' && decimals === '') setDecimals('0');
  };

  const fillFromColumn = () => {
    const found = columnValues(sheet, header, col, type === 'multi');
    const have = new Set(options.map((o) => o.v));
    setOptions([...options, ...found.filter((v) => !have.has(v)).map((v) => ({ v }))]);
  };

  const apply = () => {
    const base: ColumnMeta = existing && !existing.unknownType ? { ...existing } : makeMeta(sheet, col, type, existing ? { id: existing.id } : {});
    const meta: ColumnMeta = { id: base.id, title, col, type };
    if (isChoice) {
      const clean = options.map((o) => ({ ...o, v: o.v.trim() })).filter((o, i, all) => o.v && all.findIndex((x) => x.v === o.v) === i);
      meta.options = clean;
      if (source === 'range') {
        const r = ref.trim();
        const bang = r.lastIndexOf('!');
        const okRef = r.startsWith('[') ? /^\[[^\]]+\]\s*.+$/.test(r) : !!r && !!parseRef(bang > 0 ? r.slice(bang + 1) : r);
        if (!okRef) return setError('选项区域的写法不对，例如 Sheet2!A2:A20，或别的表格 [供应商.xlsx]Sheet1!A2:A30');
        meta.optionsRef = r;
      }
      meta.onInvalid = onInvalid;
    }
    if (isNum) {
      const lo = min.trim() === '' ? undefined : Number(min);
      const hi = max.trim() === '' ? undefined : Number(max);
      if ((lo !== undefined && !Number.isFinite(lo)) || (hi !== undefined && !Number.isFinite(hi))) return setError('取值范围要填数字');
      if (lo !== undefined && hi !== undefined && lo > hi) return setError('最小值不能大于最大值');
      const dec = decimals.trim() === '' ? undefined : Math.max(0, Math.min(10, Math.floor(Number(decimals))));
      if (dec !== undefined && !Number.isFinite(dec)) return setError('小数位数要填数字');
      const fmt = buildNumFmt(type, dec, thousands);
      if (fmt) meta.fmt = fmt;
      if (lo !== undefined) meta.min = lo;
      if (hi !== undefined) meta.max = hi;
      if (integer) meta.integer = true;
      meta.onInvalid = onInvalid;
    }
    if (type === 'date') {
      meta.fmt = dateFmt;
      meta.onInvalid = onInvalid;
    }
    if (type === 'check' || type === 'link' || type === 'email' || type === 'phone') meta.onInvalid = onInvalid;
    if (required) meta.required = true;
    if (def.trim()) meta.default = def.trim();
    if (note.trim()) meta.note = note.trim();
    if (existing?.ext) meta.ext = existing.ext;
    if (meta.onInvalid === 'stop') delete meta.onInvalid;
    onApply(type === 'text' && !meta.required && !meta.default && !meta.note ? null : meta);
  };

  const numSample = useMemo(() => {
    const dec = decimals.trim() === '' ? undefined : Math.max(0, Math.min(10, Math.floor(Number(decimals)) || 0));
    const n = (1234.5678).toFixed(dec ?? 2);
    const [i, f] = n.split('.');
    const body = thousands ? i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : i;
    const out = f ? `${body}.${f}` : body;
    return type === 'money' ? `¥${out}` : type === 'percent' ? `${(123456.78).toFixed(dec ?? 0)}%` : out;
  }, [decimals, thousands, type]);

  const left = Math.max(8, Math.min(at.x, window.innerWidth - 392));
  const top = Math.max(8, Math.min(at.y, window.innerHeight - 420));

  return (
    <div className="cp-mask" onMouseDown={onClose}>
      <div className="cp" role="dialog" aria-label="列属性" style={{ left, top, maxHeight: window.innerHeight - top - 8 }} onMouseDown={(e) => e.stopPropagation()}>
        <div className="cp-head">
          <div>
            <div className="cp-title">列属性</div>
            <div className="cp-sub">
              {colName(col)} 列 ·「{title || '（无表头）'}」· 应用于第 {header + 1} 行起整列
            </div>
          </div>
          <button className="cp-x" aria-label="关闭" onClick={onClose}>
            ×
          </button>
        </div>

        <div className="cp-body">
          <div className="cp-sec">
            <div className="cp-lbl">字段类型</div>
            <div className="cp-chips">
              {CHIPS.map((t) => (
                <button key={t} className={`cp-chip${t === type ? ' on' : ''}`} aria-pressed={t === type} onClick={() => changeType(t)}>
                  <span className="cp-glyph">{COLUMN_TYPE_GLYPHS[t]}</span>
                  {t === 'link' ? '网址' : COLUMN_TYPE_LABELS[t]}
                </button>
              ))}
            </div>
          </div>

          {isChoice && (
            <div className="cp-sec">
              <div className="cp-lbl">选项来源</div>
              <div className="cp-seg">
                <button className={source === 'manual' ? 'on' : ''} onClick={() => setSource('manual')}>
                  手动填写
                </button>
                <button className={source === 'range' ? 'on' : ''} onClick={() => setSource('range')}>
                  引用区域
                </button>
              </div>
              {source === 'range' && (
                <input className="cp-input" aria-label="选项区域" placeholder="如 Sheet2!A2:A20，或包里别的表格 [供应商.xlsx]Sheet1!A2:A30" value={ref} onChange={(e) => setRef(e.target.value)} />
              )}
              {source === 'manual' && (
                <>
                  <div className="cp-opts">
                    {options.map((o, i) => (
                      <div key={i} className="cp-optwrap">
                        <div className="cp-opt">
                          <button
                            className="cp-dot"
                            aria-label="选项颜色"
                            title="点色块换颜色"
                            style={{ background: o.color ?? OPTION_PALETTE[i % OPTION_PALETTE.length] }}
                            onClick={() => setPaletteFor(paletteFor === i ? null : i)}
                          />
                          <input
                            aria-label="选项"
                            value={o.v}
                            onChange={(e) => setOptions(options.map((x, k) => (k === i ? { ...x, v: e.target.value } : x)))}
                          />
                          <span className="cp-count" title="本列里用了几次">
                            {counts.get(o.v.trim()) ?? 0}
                          </span>
                          <button className="cp-del" aria-label="删除选项" onClick={() => setOptions(options.filter((_, k) => k !== i))}>
                            ×
                          </button>
                        </div>
                        {paletteFor === i && (
                          <div className="cp-palette">
                            {OPTION_PALETTE.map((c) => (
                              <button
                                key={c}
                                aria-label={`颜色 ${c}`}
                                style={{ background: c }}
                                className={(o.color ?? OPTION_PALETTE[i % OPTION_PALETTE.length]) === c ? 'on' : ''}
                                onClick={() => {
                                  setOptions(options.map((x, k) => (k === i ? { ...x, color: c } : x)));
                                  setPaletteFor(null);
                                }}
                              />
                            ))}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                  <div className="cp-row">
                    <button className="cp-link" onClick={() => setOptions([...options, { v: '' }])}>
                      + 添加选项
                    </button>
                    <button className="cp-link" onClick={fillFromColumn}>
                      从本列已有值补全
                    </button>
                  </div>
                </>
              )}
            </div>
          )}

          {isNum && (
            <div className="cp-sec">
              <div className="cp-row">
                <label htmlFor="cp-dec">小数位数</label>
                <input id="cp-dec" className="cp-input sm" inputMode="numeric" placeholder="不限" value={decimals} onChange={(e) => setDecimals(e.target.value)} />
                {type !== 'percent' && (
                  <label className="cp-check">
                    <input type="checkbox" checked={thousands} onChange={(e) => setThousands(e.target.checked)} />
                    千位分隔符
                  </label>
                )}
              </div>
              <div className="cp-row">
                <span>取值范围</span>
                <input className="cp-input sm" aria-label="最小值" placeholder="最小" value={min} onChange={(e) => setMin(e.target.value)} />
                <span>至</span>
                <input className="cp-input sm" aria-label="最大值" placeholder="不限" value={max} onChange={(e) => setMax(e.target.value)} />
              </div>
              <label className="cp-check">
                <input type="checkbox" checked={integer} onChange={(e) => setInteger(e.target.checked)} />
                只允许整数
              </label>
              <div className="cp-note">
                示例：<b>{numSample}</b>
              </div>
            </div>
          )}

          {type === 'date' && (
            <div className="cp-sec">
              <div className="cp-lbl">显示格式</div>
              <div className="cp-radios">
                {DATE_FORMATS.map((d) => (
                  <label key={d.fmt}>
                    <input type="radio" name="cp-df" checked={dateFmt === d.fmt} onChange={() => setDateFmt(d.fmt)} />
                    {d.label}
                  </label>
                ))}
              </div>
            </div>
          )}

          {type === 'check' && (
            <div className="cp-sec cp-note">
              单元格里写入「是」显示为勾选，「否」显示为空框；点单元格里的方框或按空格切换。本列已有的其它值保持原样，标为不合规。
            </div>
          )}

          {(type === 'link' || type === 'email' || type === 'phone') && (
            <div className="cp-sec">
              <div className="cp-lbl">校验内容</div>
              <div className="cp-seg">
                {(['link', 'email', 'phone'] as const).map((t) => (
                  <button key={t} className={type === t ? 'on' : ''} onClick={() => changeType(t)}>
                    {t === 'link' ? '网址' : t === 'email' ? '邮箱' : '电话'}
                  </button>
                ))}
              </div>
            </div>
          )}

          {(type === 'autonum' || type === 'created') && (
            <div className="cp-sec cp-note">{type === 'autonum' ? '用表单新建记录时自动填「本列最大值 + 1」，表单里只读。' : '用表单新建记录时自动填当前时间，表单里只读。'}</div>
          )}

          {type === 'text' && !required && !def && <div className="cp-sec cp-note">普通输入，不做限制。下面仍可设为必填或加默认值 / 说明。</div>}

          {type !== 'autonum' && type !== 'created' && (
            <>
              <div className="cp-sec cp-two">
                <label className="cp-switch">
                  <span>
                    <b>必填</b>
                    <small>表单里不填不能提交，表格里空着标红</small>
                  </span>
                  <button role="switch" aria-checked={required} aria-label="必填" className={required ? 'on' : ''} onClick={() => setRequired(!required)}>
                    <i />
                  </button>
                </label>
                <div className="cp-field">
                  <label htmlFor="cp-def">
                    <b>默认值</b>
                    <small>新建记录时自动填入</small>
                  </label>
                  <input id="cp-def" className="cp-input" placeholder={type === 'date' ? '如 {{today}}、{{today+7}}' : '可不填'} value={def} onChange={(e) => setDef(e.target.value)} />
                  {(isChoice || type === 'check') && (
                    <div className="cp-quick">
                      {(type === 'check' ? ['是', '否'] : options.map((o) => o.v).filter(Boolean)).map((v) => (
                        <button key={v} className={def === v ? 'on' : ''} onClick={() => setDef(def === v ? '' : v)}>
                          {v}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>
              {(isChoice || isNum || type === 'date' || type === 'check' || type === 'link' || type === 'email' || type === 'phone') && (
                <div className="cp-sec">
                  <div className="cp-lbl">输入不合规时</div>
                  <div className="cp-radios row">
                    {INVALID.map((o) => (
                      <label key={o.v}>
                        <input type="radio" name="cp-inv" checked={onInvalid === o.v} onChange={() => setOnInvalid(o.v)} />
                        {o.label}
                      </label>
                    ))}
                  </div>
                </div>
              )}
              <div className="cp-sec">
                <label className="cp-lbl" htmlFor="cp-note">
                  说明（表单里显示）
                </label>
                <input id="cp-note" className="cp-input" placeholder="可不填" value={note} onChange={(e) => setNote(e.target.value)} />
              </div>
            </>
          )}

          {error && (
            <div className="cp-error" role="alert">
              {error}
            </div>
          )}

          <div className="cp-info">
            列属性保存在工作包里，不改动 xlsx 文件本身；用 Excel 单独打开时仍是普通表格。多选在单元格里是用逗号连接的文字（如「合格证,材质证明」）。
          </div>
        </div>

        <div className="cp-foot">
          {existing && (
            <button className="btn ghost" onClick={() => onApply(null)}>
              清除列属性
            </button>
          )}
          <span className="grow" />
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" onClick={apply}>
            应用
          </button>
        </div>
      </div>
    </div>
  );
}

/** 条件格式：给选中的区域新建规则（突出显示、前 N 项、平均值、色阶、数据条、公式），列出并删除工作表里的规则。 */
import { useEffect, useState } from 'react';
import { host } from '@wz/editor-shared';
import { Icon } from '@wz/editor-shared/Icon';
import { addCfRule, CF_BAR_COLORS, CF_LOOKS, CF_OPERATORS, CF_SCALES, describeRule, rangeRef, rangesText } from '../core/cfEdit';
import type { CfRule, Range, SheetData } from '../core/workbook';

type Kind = 'cellIs' | 'text' | 'duplicate' | 'unique' | 'top' | 'bottom' | 'above' | 'below' | 'scale' | 'bar' | 'expression';

const KINDS: { kind: Kind; label: string }[] = [
  { kind: 'cellIs', label: '单元格的值（大于、小于、介于…）' },
  { kind: 'text', label: '文本包含' },
  { kind: 'duplicate', label: '重复值' },
  { kind: 'unique', label: '唯一值' },
  { kind: 'top', label: '值最大的前 N 项' },
  { kind: 'bottom', label: '值最小的后 N 项' },
  { kind: 'above', label: '高于平均值' },
  { kind: 'below', label: '低于平均值' },
  { kind: 'scale', label: '色阶' },
  { kind: 'bar', label: '数据条' },
  { kind: 'expression', label: '用公式确定（如 =$A1>100）' },
];

type NewRule = Parameters<typeof addCfRule>[2];

/** 输入的值 → 公式文字：数字照写，“=”开头是公式，其它当文字加引号 */
function valueFormula(v: string): string {
  const t = v.trim();
  if (t.startsWith('=')) return t.slice(1);
  if (t !== '' && Number.isFinite(Number(t))) return String(Number(t));
  return `"${t.replace(/"/g, '""')}"`;
}

interface Props {
  sheet: SheetData;
  /** 选中的区域（新规则的作用范围） */
  range: Range;
  onAdd: (rule: NewRule) => void;
  onRemove: (rule: CfRule) => void;
  /** 清除：range 为 null 时整个工作表 */
  onClear: (range: Range | null) => void;
  onClose: () => void;
}

export function CfDialog(p: Props) {
  const [kind, setKind] = useState<Kind>('cellIs');
  const [op, setOp] = useState('greaterThan');
  const [v1, setV1] = useState('');
  const [v2, setV2] = useState('');
  const [text, setText] = useState('');
  const [n, setN] = useState('10');
  const [percent, setPercent] = useState(false);
  const [formula, setFormula] = useState('');
  const [look, setLook] = useState(0);
  const [scale, setScale] = useState(0);
  const [bar, setBar] = useState(0);

  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);

  const needsLook = !['scale', 'bar'].includes(kind);

  const build = (): NewRule | string => {
    const ranges = [p.range];
    const base = { ranges, formulas: [] as string[] };
    const l = CF_LOOKS[look];
    switch (kind) {
      case 'cellIs':
        if (!v1.trim()) return '请输入比较的值';
        if ((op === 'between' || op === 'notBetween') && !v2.trim()) return '请输入范围的另一端';
        return { ...base, type: 'cellIs', operator: op, formulas: op === 'between' || op === 'notBetween' ? [valueFormula(v1), valueFormula(v2)] : [valueFormula(v1)], look: l };
      case 'text':
        if (!text) return '请输入要查找的文字';
        return { ...base, type: 'containsText', text, look: l };
      case 'duplicate':
        return { ...base, type: 'duplicateValues', look: l };
      case 'unique':
        return { ...base, type: 'uniqueValues', look: l };
      case 'top':
      case 'bottom': {
        const rank = Math.floor(Number(n));
        if (!(rank >= 1 && rank <= (percent ? 100 : 1000))) return percent ? '百分比请输入 1–100' : '项数请输入 1–1000';
        return { ...base, type: 'top10', rank, percent: percent || undefined, bottom: kind === 'bottom' || undefined, look: l };
      }
      case 'above':
      case 'below':
        return { ...base, type: 'aboveAverage', above: kind === 'above', look: l };
      case 'scale': {
        const colors = CF_SCALES[scale].colors;
        const cfvo = colors.length === 3 ? [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }];
        return { ...base, type: 'colorScale', colors, cfvo };
      }
      case 'bar':
        return { ...base, type: 'dataBar', colors: [CF_BAR_COLORS[bar].color], cfvo: [{ type: 'min' }, { type: 'max' }] };
      case 'expression': {
        const f = formula.trim().replace(/^=/, '');
        if (!f) return '请输入公式';
        return { ...base, type: 'expression', formulas: [f], look: l };
      }
    }
  };

  const add = () => {
    const r = build();
    if (typeof r === 'string') return host().toast(r);
    p.onAdd(r);
    host().toast('已添加条件格式');
  };

  const rules = p.sheet.cf ?? [];
  const here = rangeRef(p.range);
  const sw = (c: string) => <span className="sheet-cf-sw" style={{ background: c }} />;

  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-cf" role="dialog" aria-label="条件格式" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
        <div className="sheet-detail-head">
          <strong>条件格式</strong>
          <span className="sub mono">应用于 {here}</span>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>

        <div className="sheet-sort-row">
          <span className="sub">规则</span>
          <select className="ed-select" value={kind} onChange={(e) => setKind(e.target.value as Kind)}>
            {KINDS.map((k) => (
              <option key={k.kind} value={k.kind}>
                {k.label}
              </option>
            ))}
          </select>
        </div>

        {kind === 'cellIs' && (
          <div className="sheet-sort-row">
            <span className="sub">条件</span>
            <select className="ed-select" value={op} onChange={(e) => setOp(e.target.value)}>
              {CF_OPERATORS.map((o) => (
                <option key={o.op} value={o.op}>
                  {o.label}
                </option>
              ))}
            </select>
            <input className="field" placeholder="值或 =单元格" value={v1} onChange={(e) => setV1(e.target.value)} />
            {(op === 'between' || op === 'notBetween') && (
              <>
                <span className="sub">和</span>
                <input className="field" placeholder="值或 =单元格" value={v2} onChange={(e) => setV2(e.target.value)} />
              </>
            )}
          </div>
        )}
        {kind === 'text' && (
          <div className="sheet-sort-row">
            <span className="sub">文字</span>
            <input className="field" placeholder="包含这段文字（不分大小写）" value={text} onChange={(e) => setText(e.target.value)} />
          </div>
        )}
        {(kind === 'top' || kind === 'bottom') && (
          <div className="sheet-sort-row">
            <span className="sub">{kind === 'top' ? '前' : '后'}</span>
            <input className="field" style={{ width: 80 }} type="number" min={1} value={n} onChange={(e) => setN(e.target.value)} />
            <label className="sheet-cf-check">
              <input type="checkbox" checked={percent} onChange={(e) => setPercent(e.target.checked)} />
              按百分比
            </label>
          </div>
        )}
        {kind === 'expression' && (
          <div className="sheet-sort-row">
            <span className="sub">公式</span>
            <input className="field mono" placeholder={`相对引用以 ${here.split(':')[0]} 为准，如 =$A1>100`} value={formula} onChange={(e) => setFormula(e.target.value)} />
          </div>
        )}

        {needsLook && (
          <div className="sheet-sort-row">
            <span className="sub">设置为</span>
            <select className="ed-select" value={look} onChange={(e) => setLook(Number(e.target.value))}>
              {CF_LOOKS.map((l, i) => (
                <option key={l.label} value={i}>
                  {l.label}
                </option>
              ))}
            </select>
            <span className="sheet-cf-demo" style={{ background: CF_LOOKS[look].fill, color: CF_LOOKS[look].color, fontWeight: CF_LOOKS[look].bold ? 700 : undefined }}>
              AaBb 示例
            </span>
          </div>
        )}
        {kind === 'scale' && (
          <div className="sheet-sort-row">
            <span className="sub">颜色</span>
            <select className="ed-select" value={scale} onChange={(e) => setScale(Number(e.target.value))}>
              {CF_SCALES.map((s, i) => (
                <option key={s.label} value={i}>
                  {s.label}
                </option>
              ))}
            </select>
            <span className="sheet-cf-swatches">{CF_SCALES[scale].colors.map((c) => sw(c))}</span>
          </div>
        )}
        {kind === 'bar' && (
          <div className="sheet-sort-row">
            <span className="sub">颜色</span>
            <select className="ed-select" value={bar} onChange={(e) => setBar(Number(e.target.value))}>
              {CF_BAR_COLORS.map((c, i) => (
                <option key={c.label} value={i}>
                  {c.label}
                </option>
              ))}
            </select>
            <span className="sheet-cf-swatches">{sw(CF_BAR_COLORS[bar].color)}</span>
          </div>
        )}

        <div className="sheet-detail-foot">
          <button className="btn primary" onClick={add}>
            添加规则
          </button>
        </div>

        <div className="sheet-cf-list">
          <div className="sheet-cf-list-head">
            <strong>本工作表的规则（{rules.length}）</strong>
            <span className="grow" />
            <button className="btn" disabled={!rules.length} title="清除与当前选区有重叠的规则" onClick={() => p.onClear(p.range)}>
              清除所选区域的规则
            </button>
            <button className="btn" disabled={!rules.length} onClick={() => p.onClear(null)}>
              清除整个工作表的规则
            </button>
          </div>
          {rules.length === 0 && <div className="sub">还没有条件格式。</div>}
          {rules.map((r, i) => (
            <div key={`${r.idx ?? 'n'}-${i}`} className="sheet-cf-rule">
              <span className="grow ellipsis" title={`${describeRule(r)}　${rangesText(r)}`}>
                {describeRule(r)}
              </span>
              <span className="sub mono">{rangesText(r)}</span>
              <button className="icon-btn" aria-label="删除这条规则" title="删除这条规则" onClick={() => p.onRemove(r)}>
                <Icon name="trash" size={12} />
              </button>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/** 选择性粘贴对话框：粘贴内容（全部 / 公式 / 值 / 格式）、运算、跳过空单元格、转置。 */
import { useEffect, useState } from 'react';
import { host } from '@wz/editor-shared';
import { Icon } from '@wz/editor-shared/Icon';
import type { PasteMode, PasteOp, PasteOpts } from '../core/ops';

export interface PasteSpecialChoice {
  mode: PasteMode;
  transpose: boolean;
  opts: PasteOpts;
}

const WHAT: [PasteMode, string][] = [
  ['all', '全部'],
  ['formulas', '公式（不含格式）'],
  ['values', '数值（不含公式和格式）'],
  ['formats', '格式'],
];
const OPS: [PasteOp | '', string][] = [
  ['', '无'],
  ['add', '加'],
  ['sub', '减'],
  ['mul', '乘'],
  ['div', '除'],
];

export function PasteSpecialDialog(p: { onOk: (c: PasteSpecialChoice) => void; onClose: () => void }) {
  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  const [mode, setMode] = useState<PasteMode>('all');
  const [op, setOp] = useState<PasteOp | ''>('');
  const [skipBlanks, setSkipBlanks] = useState(false);
  const [transpose, setTranspose] = useState(false);
  // 运算和格式、转置不能一起用：运算只改数值
  const opOn = op !== '';
  const ok = () => p.onOk({ mode: opOn ? 'values' : mode, transpose: opOn ? false : transpose, opts: { op: opOn ? op : undefined, skipBlanks } });
  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div
        className="sheet-detail sheet-sort"
        role="dialog"
        aria-label="选择性粘贴"
        onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => (e.key === 'Escape' ? p.onClose() : e.key === 'Enter' && ok())}
      >
        <div className="sheet-detail-head">
          <strong>选择性粘贴</strong>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        <div className="sub" style={{ padding: '4px 0' }}>
          粘贴
        </div>
        <div className="sheet-data-cols">
          {WHAT.map(([m, label]) => (
            <label key={m}>
              <input type="radio" name="ps-what" checked={mode === m} disabled={opOn} onChange={() => setMode(m)} />
              {label}
            </label>
          ))}
        </div>
        <div className="sub" style={{ padding: '4px 0' }}>
          运算（只对数字生效，目标是公式的格子不变）
        </div>
        <div className="sheet-data-cols">
          {OPS.map(([o, label]) => (
            <label key={o || 'none'}>
              <input type="radio" name="ps-op" checked={op === o} onChange={() => setOp(o)} />
              {label}
            </label>
          ))}
        </div>
        <div className="sheet-data-cols">
          <label>
            <input type="checkbox" checked={skipBlanks} onChange={(e) => setSkipBlanks(e.target.checked)} />
            跳过空单元格
          </label>
          <label>
            <input type="checkbox" checked={transpose && !opOn} disabled={opOn} onChange={(e) => setTranspose(e.target.checked)} />
            转置
          </label>
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

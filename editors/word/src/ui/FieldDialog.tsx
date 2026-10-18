/**
 * 编辑域对话框（同 Word 的“域”对话框的简化版）：左边选域名，右边按域名给出常用选项，下面是域代码（可以直接改）和结果预览。
 * 选项只是改域代码的快捷方式：代码是唯一的依据，改代码后选项跟着变。
 */
import { useMemo, useState } from 'react';
import { COMPUTABLE, FIELD_INFO, evalField, formatInstr, parseInstr, withSwitch, type FieldEnv, type Instr } from '../core/fields';
import { Dialog } from './WordDialogs';

const DATE_PICS = ['yyyy/M/d', 'yyyy-MM-dd', 'yyyy年M月d日', 'yyyy年M月d日 dddd', 'yyyy年M月', 'M月d日', 'yyyy/M/d H:mm', 'yyyy年M月d日 H时m分'];
const TIME_PICS = ['H:mm', 'HH:mm:ss', 'h:mm am/pm', 'H时m分', 'H时m分s秒'];
const NUM_PICS = ['0', '0.00', '#,##0', '#,##0.00', '0%'];
const SEQ_STYLES: { v: string; label: string }[] = [
  { v: 'ARABIC', label: '1, 2, 3' },
  { v: 'ROMAN', label: 'I, II, III' },
  { v: 'roman', label: 'i, ii, iii' },
  { v: 'ALPHABETIC', label: 'A, B, C' },
  { v: 'alphabetic', label: 'a, b, c' },
  { v: 'CHINESENUM3', label: '一, 二, 三' },
  { v: 'CHINESENUM2', label: '壹, 贰, 叁' },
];
const CASES: { v: string; label: string }[] = [
  { v: 'Upper', label: '大写' },
  { v: 'Lower', label: '小写' },
  { v: 'FirstCap', label: '首字母大写' },
  { v: 'Caps', label: '每个单词首字母大写' },
];

/** 选了一个域名时的默认代码 */
function defaultCode(name: string, bookmarks: string[]): string {
  switch (name) {
    case 'DATE':
      return 'DATE \\@ "yyyy年M月d日" \\* MERGEFORMAT';
    case 'TIME':
      return 'TIME \\@ "H:mm" \\* MERGEFORMAT';
    case 'REF':
      return `REF ${bookmarks[0] ?? '书签名'} \\h \\* MERGEFORMAT`;
    case 'SEQ':
      return 'SEQ 表 \\* ARABIC \\* MERGEFORMAT';
    case '=':
      return '= 1 + 1 \\* MERGEFORMAT';
    case 'QUOTE':
      return 'QUOTE "文字" \\* MERGEFORMAT';
    default:
      return `${name} \\* MERGEFORMAT`;
  }
}

const starOf = (i: Instr, vals: string[]) => i.sw.find((s) => s.key === '*' && vals.some((v) => v.toLowerCase() === (s.value ?? '').toLowerCase()))?.value ?? '';

export interface FieldDialogProps {
  instr: string;
  env: FieldEnv;
  bookmarks: string[];
  onOk: (instr: string) => void;
  onClose: () => void;
}

export function FieldDialog({ instr, env, bookmarks, onOk, onClose }: FieldDialogProps) {
  const [code, setCode] = useState(instr);
  const parsed = useMemo(() => parseInstr(code), [code]);
  const name = parsed.name.toUpperCase();
  const preview = useMemo(() => evalField(code, env), [code, env]);
  const set = (i: Instr) => setCode(formatInstr(i));
  const sw = (key: string) => parsed.sw.find((s) => s.key === key);
  const arg0 = parsed.args[0] ?? '';
  const withArg = (v: string) => set({ ...parsed, args: [v, ...parsed.args.slice(1)] });
  const keep = parsed.sw.some((s) => s.key === '*' && s.value?.toUpperCase() === 'MERGEFORMAT');
  const toggleKeep = (on: boolean) => set(on ? { ...parsed, sw: [...parsed.sw, { key: '*', value: 'MERGEFORMAT' }] } : { ...parsed, sw: parsed.sw.filter((s) => !(s.key === '*' && s.value?.toUpperCase() === 'MERGEFORMAT')) });
  const known = FIELD_INFO.some((f) => f.name === name);

  const pickName = (n: string) => setCode(n === name ? code : defaultCode(n, bookmarks));

  const caseSelect = (
    <label className="label">
      文字格式
      <select className="field" value={starOf(parsed, CASES.map((c) => c.v))} onChange={(e) => set(withSwitch(parsed, '*', e.target.value || null))}>
        <option value="">（无）</option>
        {CASES.map((c) => (
          <option key={c.v} value={c.v}>
            {c.label}
          </option>
        ))}
      </select>
    </label>
  );

  const options = (() => {
    switch (name) {
      case 'DATE':
      case 'TIME': {
        const pics = name === 'DATE' ? DATE_PICS : TIME_PICS;
        const cur = sw('@')?.value ?? '';
        return (
          <>
            <div className="label">
              日期格式
              <div className="wd-fld-pics" role="listbox" aria-label="日期格式">
                {pics.map((p) => (
                  <button key={p} type="button" role="option" aria-selected={p === cur} className={p === cur ? 'on' : undefined} onClick={() => set(withSwitch(parsed, '@', p))}>
                    {evalField(`${name} \\@ "${p}"`, env)}
                    <span>{p}</span>
                  </button>
                ))}
              </div>
            </div>
            <label className="label">
              自定义格式
              <input className="field" value={cur} placeholder="yyyy年M月d日" onChange={(e) => set(withSwitch(parsed, '@', e.target.value || null))} />
            </label>
          </>
        );
      }
      case 'FILENAME':
        return (
          <>
            <label className="wd-fld-check">
              <input type="checkbox" checked={!!sw('p')} onChange={(e) => set(withSwitch(parsed, 'p', e.target.checked ? true : null))} /> 添加路径
            </label>
            {caseSelect}
          </>
        );
      case 'REF':
        return (
          <>
            <label className="label">
              书签
              <select className="field" value={arg0} onChange={(e) => withArg(e.target.value)}>
                {!bookmarks.length && !arg0 && <option value="">（文档里没有书签）</option>}
                {arg0 && !bookmarks.includes(arg0) && <option value={arg0}>{arg0}（已不存在）</option>}
                {bookmarks.map((b) => (
                  <option key={b}>{b}</option>
                ))}
              </select>
            </label>
            {caseSelect}
          </>
        );
      case 'SEQ':
        return (
          <>
            <label className="label">
              题注标识
              <input className="field" list="wd-seq-ids" value={arg0} onChange={(e) => withArg(e.target.value.replace(/\s+/g, ''))} />
              <datalist id="wd-seq-ids">
                <option value="图" />
                <option value="表" />
                <option value="公式" />
              </datalist>
            </label>
            <label className="label">
              编号格式
              <select className="field" value={starOf(parsed, SEQ_STYLES.map((s) => s.v)) || 'ARABIC'} onChange={(e) => set(withSwitch(parsed, '*', e.target.value))}>
                {SEQ_STYLES.map((s) => (
                  <option key={s.v} value={s.v}>
                    {s.label}
                  </option>
                ))}
              </select>
            </label>
            <label className="label">
              重新编号，从这个数开始（可不填）
              <input className="field" inputMode="numeric" value={sw('r')?.value ?? ''} onChange={(e) => set(withSwitch(parsed, 'r', e.target.value.replace(/\D/g, '') || null))} />
            </label>
          </>
        );
      case '=':
        return (
          <>
            <label className="label">
              公式
              <input className="field" value={arg0} placeholder="例如 (120 + 80) * 1.13，可用书签名、SUM、ROUND 等" onChange={(e) => set({ ...parsed, args: e.target.value.trim() ? [e.target.value] : [] })} />
            </label>
            <label className="label">
              数字格式
              <select className="field" value={sw('#')?.value ?? ''} onChange={(e) => set(withSwitch(parsed, '#', e.target.value || null))}>
                <option value="">（常规）</option>
                {NUM_PICS.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
            </label>
          </>
        );
      case 'QUOTE':
        return (
          <>
            <label className="label">
              文字
              <input className="field" value={arg0} onChange={(e) => withArg(e.target.value)} />
            </label>
            {caseSelect}
          </>
        );
      case 'NUMPAGES':
        return <div className="sub">没有其它选项</div>;
      default:
        return <div className="sub">这个域没有可选的选项，可以在下面直接修改域代码。</div>;
    }
  })();

  const submit = () => {
    if (!code.trim()) return;
    onOk(code.trim());
  };

  return (
    <Dialog
      title="编辑域"
      width={600}
      onClose={onClose}
      foot={
        <>
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" disabled={!code.trim()} onClick={submit}>
            确定
          </button>
        </>
      }
    >
      <div className="wd-fld">
        <div className="wd-fld-names" role="listbox" aria-label="域名">
          {FIELD_INFO.map((f) => (
            <button key={f.name} type="button" role="option" aria-selected={f.name === name} className={f.name === name ? 'on' : undefined} title={f.desc} onClick={() => pickName(f.name)}>
              <span>{f.label}</span>
              <span className="mono">{f.name}</span>
            </button>
          ))}
          {!known && (
            <button type="button" role="option" aria-selected className="on">
              <span>其它</span>
              <span className="mono">{name || '（空）'}</span>
            </button>
          )}
        </div>
        <div className="wd-fld-opts">
          {options}
          <label className="wd-fld-check">
            <input type="checkbox" checked={keep} onChange={(e) => toggleKeep(e.target.checked)} /> 更新时保留原格式
          </label>
        </div>
      </div>
      <label className="label" style={{ marginTop: 12 }}>
        域代码
        <input className="field mono" value={code} spellCheck={false} onChange={(e) => setCode(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && submit()} />
      </label>
      <div className="wd-fld-preview">
        结果预览：
        {preview === null ? <span className="sub">{COMPUTABLE.has(name) ? '算不出结果（请检查域代码）' : '这个域编辑器暂时算不出结果，确定后保留原来的结果'}</span> : <strong>{preview || '（空）'}</strong>}
      </div>
    </Dialog>
  );
}

/**
 * 表单录入面板（表格右侧）：按列属性生成控件，必填 / 范围 / 选项校验，「提交」写入下一个空行（一步撤销），
 * 「提交并继续新建」（Ctrl+Enter）清空后接着录；标了「沿用」的字段保留上一条的值；也可以改当前行。
 * 只在工作包里的表格出现。表单定义存在工作包里（forms/<id>.json），没保存过表单时用默认表单（所有已设类型的列）。
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import { COLUMN_TYPE_LABELS, OPTION_PALETTE, host, joinMulti, splitMulti, type FormDef } from '@wz/editor-shared';
import { optionList } from '../core/colmeta';
import { initialValues, nextAutoNumber, resolveForm, rowValues, validateValues, type FormTarget, type FormValues, type ResolvedField } from '../core/forms';
import type { SheetData, Workbook } from '../core/workbook';
import './columns.css';

/** 提交结果：写入的行 / 校验脚本不通过（errors 是要给用户看的原因）/ 失败（受保护等，提示已由调用方给出） */
export type FormSubmitResult = { row: number } | { errors: string[] } | null;

export interface FormFillProps {
  book: Workbook;
  sheet: SheetData;
  forms: FormDef[];
  /** 打开时选中的表单 id（空 = 默认表单） */
  formId: string;
  user: string | null;
  /** 活动单元格所在的行（0 起），“编辑当前行”用 */
  activeRow: number;
  /** 提交：写进表格并返回写入的行；失败（受保护等）返回 null */
  onSubmit: (fields: ResolvedField[], values: FormValues, target: FormTarget, form: FormDef | null) => FormSubmitResult | Promise<FormSubmitResult>;
  onDesign: (formId: string) => void;
  /** 生成离线填写页（加进工作包）/ 导入离线填写结果；只有保存过的表单才有 */
  onExportHtml: (form: FormDef) => void;
  onImportResult: (form: FormDef) => void;
  onClose: () => void;
}

const LONG_TEXT = /备注|说明|描述|原因|意见/;

export function FormFill({ book, sheet, forms, formId: initialFormId, user, activeRow, onSubmit, onDesign, onExportHtml, onImportResult, onClose }: FormFillProps) {
  const [formId, setFormId] = useState(initialFormId);
  const form = forms.find((f) => f.id === formId) ?? null;
  const fields = useMemo(() => resolveForm(sheet.columns, form), [sheet.columns, form]);
  const shown = fields.filter((f) => !f.hidden);
  const header = sheet.columns?.header ?? 1;
  const [mode, setMode] = useState<'new' | 'edit'>('new');
  const [values, setValues] = useState<FormValues>(() => initialValues(fields, user, new Date()));
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [done, setDone] = useState(0);
  const [checking, setChecking] = useState(false);
  /** 校验脚本不通过的原因（表单校验之外的、跨表的规则） */
  const [scriptErrs, setScriptErrs] = useState<string[]>([]);
  const [keepAll, setKeepAll] = useState(false);
  const [keeps, setKeeps] = useState<Record<string, boolean>>({});
  const root = useRef<HTMLDivElement>(null);
  const lastValues = useRef<FormValues | undefined>(undefined);

  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);

  const isKeep = (f: ResolvedField) => keepAll || (keeps[f.meta.id] ?? f.keep);

  /** 切换表单 / 模式后重新取初值 */
  const reset = (nextFields: ResolvedField[], nextMode: 'new' | 'edit', keepPrev = false) => {
    setErrors({});
    if (nextMode === 'edit') setValues(rowValues(sheet, nextFields, activeRow));
    else setValues(initialValues(nextFields.map((f) => ({ ...f, keep: keepAll || (keeps[f.meta.id] ?? f.keep) })), user, new Date(), keepPrev ? lastValues.current : undefined));
  };

  const change = (id: string, text: string) => {
    setValues((v) => ({ ...v, [id]: text }));
    if (errors[id]) setErrors((e) => ({ ...e, [id]: '' }));
  };

  const editable = mode === 'new' || activeRow >= header;
  const missing = shown.filter((f) => !f.auto && f.meta.required && !(values[f.meta.id] ?? '').trim()).length;

  const submit = async (again: boolean) => {
    if (!editable || checking) return;
    const errs = validateValues(fields, values);
    if (Object.keys(errs).length) {
      setErrors(errs);
      const first = fields.find((f) => errs[f.meta.id] && !f.hidden);
      root.current?.querySelector<HTMLElement>(`[data-field="${first?.meta.id}"] input, [data-field="${first?.meta.id}"] textarea, [data-field="${first?.meta.id}"] button`)?.focus();
      return;
    }
    const target: FormTarget = mode === 'new' ? { kind: 'new' } : { kind: 'edit', row: activeRow };
    setChecking(true);
    let res: FormSubmitResult;
    try {
      res = await onSubmit(fields, values, target, form);
    } finally {
      setChecking(false);
    }
    if (!res) return;
    if ('errors' in res) return setScriptErrs(res.errors);
    setScriptErrs([]);
    if (mode === 'new') {
      lastValues.current = values;
      setDone((n) => n + 1);
      // “提交并继续”或表单设为提交后继续新建：留在面板里接着录；否则关闭
      if (again || form?.after === 'continue') reset(fields, 'new', true);
      else onClose();
    } else {
      host().toast(`已更新第 ${res.row + 1} 行`);
      if (!again) onClose();
    }
  };

  const control = (f: ResolvedField) => {
    const m = f.meta;
    const id = m.id;
    const value = values[id] ?? '';
    const err = errors[id];
    const inputId = `ff-${id}`;
    if (f.auto || m.type === 'autonum' || m.type === 'created') {
      const text = mode === 'edit' ? value : m.type === 'autonum' || f.auto === 'autonum' ? `${nextAutoNumber(sheet, m.col)}（提交时自动填）` : '提交时自动填';
      return <div className="ff-auto">{text}</div>;
    }
    switch (m.type) {
      case 'select': {
        const opts = optionList(book, sheet, m);
        return (
          <div className="ff-pills" role="radiogroup" aria-label={f.label}>
            {opts.map((o, i) => (
              <button
                key={o.v}
                role="radio"
                aria-checked={value === o.v}
                className={`ff-pill${value === o.v ? ' on' : ''}`}
                onClick={() => change(id, value === o.v ? '' : o.v)}
                style={value === o.v ? { borderColor: o.color ?? OPTION_PALETTE[i % OPTION_PALETTE.length] } : undefined}
              >
                {o.v}
              </button>
            ))}
            {opts.length === 0 && <span className="sub">（列属性里还没有选项）</span>}
          </div>
        );
      }
      case 'multi': {
        const opts = optionList(book, sheet, m);
        const on = splitMulti(value);
        return (
          <div className="ff-pills" role="group" aria-label={f.label}>
            {opts.map((o) => {
              const sel = on.includes(o.v);
              return (
                <button key={o.v} aria-pressed={sel} className={`ff-pill${sel ? ' on' : ''}`} onClick={() => change(id, joinMulti(sel ? on.filter((x) => x !== o.v) : [...on, o.v]))}>
                  {sel ? '✓ ' : ''}
                  {o.v}
                </button>
              );
            })}
            {opts.length === 0 && <span className="sub">（列属性里还没有选项）</span>}
          </div>
        );
      }
      case 'check':
        return (
          <button className="ff-check" role="checkbox" aria-checked={value === '是'} aria-label={f.label} onClick={() => change(id, value === '是' ? '否' : '是')}>
            <span className={`ff-box${value === '是' ? ' on' : ''}`}>{value === '是' ? '✓' : ''}</span>
            {value === '是' ? '是' : '否'}
          </button>
        );
      case 'date':
        return <input id={inputId} type="date" className={`ff-input${err ? ' bad' : ''}`} value={/^\d{4}-\d{2}-\d{2}$/.test(value) ? value : ''} onChange={(e) => change(id, e.target.value)} />;
      case 'number':
      case 'money':
      case 'percent':
        return (
          <div className={`ff-num${err ? ' bad' : ''}`}>
            {m.type === 'money' && <span>¥</span>}
            <input id={inputId} inputMode="decimal" value={value} onChange={(e) => change(id, e.target.value)} />
            {m.type === 'percent' && <span>%</span>}
          </div>
        );
      default:
        return LONG_TEXT.test(m.title) ? (
          <textarea id={inputId} rows={3} className={`ff-input${err ? ' bad' : ''}`} placeholder="可不填" value={value} onChange={(e) => change(id, e.target.value)} />
        ) : (
          <input id={inputId} className={`ff-input${err ? ' bad' : ''}`} value={value} onChange={(e) => change(id, e.target.value)} />
        );
    }
  };

  return (
    <div
      ref={root}
      className="ff"
      role="dialog"
      aria-label="表单录入"
      onKeyDown={(e) => {
        if (e.key === 'Enter' && e.ctrlKey) {
          e.preventDefault();
          void submit(true);
        } else if (e.key === 'Escape') {
          e.stopPropagation();
          onClose();
        }
      }}
    >
      <div className="ff-head">
        <div className="ff-tabs" role="tablist">
          <button role="tab" aria-selected={mode === 'new'} className={mode === 'new' ? 'on' : ''} onClick={() => (setMode('new'), reset(fields, 'new', true))}>
            新建记录
          </button>
          <button role="tab" aria-selected={mode === 'edit'} className={mode === 'edit' ? 'on' : ''} onClick={() => (setMode('edit'), reset(fields, 'edit'))}>
            编辑当前行
          </button>
        </div>
        <span className="grow" />
        {done > 0 && <span className="ff-done">本次已录入 {done} 条</span>}
        <button className="cp-x" aria-label="关闭" onClick={onClose}>
          ×
        </button>
      </div>
      <div className="ff-sub">
        <select
          className="ed-select"
          aria-label="表单"
          value={formId}
          onChange={(e) => {
            setFormId(e.target.value);
            const next = resolveForm(sheet.columns, forms.find((f) => f.id === e.target.value) ?? null);
            reset(next, mode);
          }}
        >
          <option value="">默认表单（所有已设类型的列）</option>
          {forms.map((f) => (
            <option key={f.id} value={f.id}>
              {f.title}
            </option>
          ))}
        </select>
        <button className="cp-link" onClick={() => onDesign(formId)}>
          {form ? '设计这个表单…' : '设计表单…'}
        </button>
      </div>
      {form && (
        <div className="ff-sub ff-share">
          <span className="sub">发给别人填：</span>
          <button className="cp-link" title="生成一个不依赖本程序的 .html 填写页，加进工作包；在工作包页导出后发给别人" onClick={() => onExportHtml(form)}>
            生成离线填写页
          </button>
          <button className="cp-link" title="导入离线填写页导出的 .json（逐条按列属性校验）" onClick={() => onImportResult(form)}>
            导入填写结果…
          </button>
        </div>
      )}

      <div className="ff-body">
        {shown.length === 0 && <div className="sg-empty">还没有可录入的列。先在列头右键“列属性…”或用“数据 → 转为智能列”给列设置类型。</div>}
        {mode === 'edit' && !editable && <div className="ff-warn">活动单元格在表头或表头以上，请先选中一个数据行。</div>}
        {shown.map((f) => (
          <div key={f.meta.id} className="ff-field" data-field={f.meta.id}>
            <div className="ff-lbl">
              <label htmlFor={`ff-${f.meta.id}`}>
                {f.label}
                {f.meta.required && !f.auto && <span className="req"> *</span>}
              </label>
              <span className="ff-type">{COLUMN_TYPE_LABELS[f.meta.type]}</span>
              {f.meta.default && mode === 'new' && !f.auto && <span className="ff-tag">默认</span>}
              <span className="grow" />
              {mode === 'new' && !f.auto && (
                <button className={`ff-keep${isKeep(f) ? ' on' : ''}`} title="提交后保留这一项的值" aria-pressed={isKeep(f)} onClick={() => setKeeps({ ...keeps, [f.meta.id]: !isKeep(f) })}>
                  沿用
                </button>
              )}
            </div>
            {control(f)}
            {(f.hint || f.meta.note) && !errors[f.meta.id] && <span className="ff-hint">{f.hint || f.meta.note}</span>}
            {errors[f.meta.id] && <span className="ff-err">{errors[f.meta.id]}</span>}
          </div>
        ))}
      </div>

      <div className="ff-foot">
        {scriptErrs.length > 0 && (
          <div className="ff-warn" role="alert">
            校验没有通过，没有写入：
            <ul style={{ margin: '4px 0 0', paddingLeft: 18 }}>
              {scriptErrs.map((m, i) => (
                <li key={i}>{m}</li>
              ))}
            </ul>
          </div>
        )}
        {mode === 'new' && (
          <label className="cp-check">
            <input type="checkbox" checked={keepAll} onChange={(e) => setKeepAll(e.target.checked)} />
            提交后保留上一条的值（标「沿用」的字段始终保留）
          </label>
        )}
        <div className="ff-actions">
          <span className="sub">{missing > 0 ? `还有 ${missing} 个必填项未填` : 'Tab 下一项 · Ctrl+Enter 提交并继续'}</span>
          <span className="grow" />
          <button className="btn" disabled={!editable || checking} onClick={() => void submit(false)}>
            {checking ? '正在校验…' : mode === 'new' ? '提交' : '保存修改'}
          </button>
          {mode === 'new' && (
            <button className="btn primary" disabled={!editable || checking} onClick={() => void submit(true)}>
              提交并继续新建
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

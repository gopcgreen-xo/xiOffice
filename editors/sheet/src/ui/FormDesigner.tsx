/**
 * 表单设计器：从本表的列里勾选要放进表单的字段，改标签 / 说明、排序、隐藏、自动值（当前时间 / 当前用户 / 自动编号）、
 * 「沿用」，右边实时预览。没有列属性的列加入表单时自动设为文本列。保存到工作包（forms/<id>.json）。
 */
import { useEffect, useMemo, useState } from 'react';
import { COLUMN_TYPE_LABELS, host, type AutoValue, type BundleItem, type ColumnMeta, type FormDef, type FormField } from '@wz/editor-shared';
import { cellText } from '../core/colmeta';
import { textMetaFor } from '../core/forms';
import type { SheetData } from '../core/workbook';
import './columns.css';

interface Row {
  col: number;
  title: string;
  meta?: ColumnMeta;
  on: boolean;
  label: string;
  hint: string;
  hidden: boolean;
  auto: AutoValue | '';
  keep: boolean;
}

interface Props {
  sheet: SheetData;
  /** 表头所在行（1 起） */
  header: number;
  /** 目标条目 id（写进表单定义） */
  itemId: string;
  /** 正在编辑的表单；null = 新建 */
  form: FormDef | null;
  newId: string;
  items: BundleItem[];
  onSave: (form: FormDef, newColumns: ColumnMeta[]) => void;
  onDelete?: () => void;
  onClose: () => void;
}

export function FormDesigner({ sheet, header, itemId, form, newId, items, onSave, onDelete, onClose }: Props) {
  const [title, setTitle] = useState(form?.title ?? '新建表单');
  const [after, setAfter] = useState<'continue' | 'close'>(form?.after ?? 'continue');
  const [onSubmit, setOnSubmit] = useState(form?.onSubmit ?? '');
  const [onValidate, setOnValidate] = useState(form?.onValidate ?? '');
  const [rows, setRows] = useState<Row[]>(() => {
    const byCol = new Map<number, ColumnMeta>((sheet.columns?.columns ?? []).filter((c) => !c.unbound).map((c) => [c.col, c]));
    const all: Row[] = [];
    for (let c = 0; c < sheet.cols; c++) {
      const t = cellText(sheet, header - 1, c).trim();
      const meta = byCol.get(c);
      if (!t && !meta) continue;
      const f = form?.fields.find((x) => x.column === meta?.id);
      all.push({
        col: c,
        title: t || meta?.title || `列 ${c + 1}`,
        ...(meta && { meta }),
        // 新建表单：默认放入所有已设类型的列
        on: form ? !!f : !!meta,
        label: f?.label ?? '',
        hint: f?.hint ?? '',
        hidden: !!f?.hidden,
        auto: f?.auto ?? '',
        keep: !!f?.keep,
      });
    }
    if (!form) return all;
    // 已有表单：按表单里的顺序在前，其它列在后
    const order = new Map(form.fields.map((f, i) => [f.column, i]));
    return all.sort((a, b) => {
      const ia = a.meta && order.has(a.meta.id) ? order.get(a.meta.id)! : 1e6 + a.col;
      const ib = b.meta && order.has(b.meta.id) ? order.get(b.meta.id)! : 1e6 + b.col;
      return ia - ib;
    });
  });

  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    document.addEventListener('keydown', onKey, true);
    return () => document.removeEventListener('keydown', onKey, true);
  }, [onClose]);

  const set = (i: number, patch: Partial<Row>) => setRows((rs) => rs.map((r, k) => (k === i ? { ...r, ...patch } : r)));
  const move = (i: number, d: number) =>
    setRows((rs) => {
      const j = i + d;
      if (j < 0 || j >= rs.length) return rs;
      const next = [...rs];
      [next[i], next[j]] = [next[j], next[i]];
      return next;
    });

  const used = rows.filter((r) => r.on);
  const scripts = items.filter((i) => i.kind === 'script');
  const previewRows = useMemo(() => used.filter((r) => !r.hidden), [used]);

  const save = () => {
    if (!used.length) return host().toast('至少选一个字段放进表单');
    if (!title.trim()) return host().toast('请给表单起个名字');
    const existing: { id: string }[] = [...(sheet.columns?.columns ?? [])];
    const created: ColumnMeta[] = [];
    const fields: FormField[] = used.map((r) => {
      let meta = r.meta;
      if (!meta) {
        meta = textMetaFor(r.col, r.title, [...existing, ...created]);
        created.push(meta);
      }
      const f: FormField = { column: meta.id };
      if (r.label.trim() && r.label.trim() !== r.title) f.label = r.label.trim();
      if (r.hint.trim()) f.hint = r.hint.trim();
      if (r.hidden) f.hidden = true;
      if (r.auto) f.auto = r.auto;
      if (r.keep) f.keep = true;
      return f;
    });
    const out: FormDef = { ...(form ?? {}), id: form?.id ?? newId, title: title.trim(), target: { item: itemId, sheet: sheet.name }, fields, after };
    if (onSubmit) out.onSubmit = onSubmit;
    else delete out.onSubmit;
    if (onValidate) out.onValidate = onValidate;
    else delete out.onValidate;
    onSave(out, created);
  };

  return (
    <div className="sg-mask" onMouseDown={onClose}>
      <div className="fd" role="dialog" aria-label="表单设计器" onMouseDown={(e) => e.stopPropagation()}>
        <div className="sg-head">
          <div>
            <div className="sg-title">{form ? '设计表单' : '新建表单'}</div>
            <div className="sg-sub">从「{sheet.name}」的列里选字段。字段的类型、选项、必填、默认值来自列属性；这里只定义表单上怎么显示。</div>
          </div>
          <button className="cp-x" aria-label="关闭" onClick={onClose}>
            ×
          </button>
        </div>
        <div className="fd-body">
          <div className="fd-left">
            <label className="fd-title">
              表单名称
              <input className="cp-input" value={title} onChange={(e) => setTitle(e.target.value)} />
            </label>
            <div className="fd-list">
              {rows.map((r, i) => (
                <div key={r.col} className={`fd-row${r.on ? '' : ' off'}`}>
                  <div className="fd-main">
                    <input type="checkbox" aria-label={`把「${r.title}」放进表单`} checked={r.on} onChange={(e) => set(i, { on: e.target.checked })} />
                    <span className="fd-name">{r.title}</span>
                    <span className="ff-type">{r.meta ? COLUMN_TYPE_LABELS[r.meta.type] : '文本（将设为文本列）'}</span>
                    <span className="grow" />
                    <button className="cp-del" aria-label="上移" disabled={i === 0} onClick={() => move(i, -1)}>
                      ↑
                    </button>
                    <button className="cp-del" aria-label="下移" disabled={i === rows.length - 1} onClick={() => move(i, 1)}>
                      ↓
                    </button>
                  </div>
                  {r.on && (
                    <div className="fd-opts">
                      <input className="cp-input sm" aria-label="标签" placeholder={`标签（默认「${r.title}」）`} value={r.label} onChange={(e) => set(i, { label: e.target.value })} />
                      <input className="cp-input sm" aria-label="说明" placeholder="说明（可不填）" value={r.hint} onChange={(e) => set(i, { hint: e.target.value })} />
                      <label className="cp-check">
                        <input type="checkbox" checked={r.hidden} onChange={(e) => set(i, { hidden: e.target.checked })} />
                        不显示
                      </label>
                      <label className="cp-check">
                        <input type="checkbox" checked={r.keep} onChange={(e) => set(i, { keep: e.target.checked })} />
                        沿用
                      </label>
                      <select className="ed-select" aria-label="自动值" value={r.auto} onChange={(e) => set(i, { auto: e.target.value as Row['auto'] })}>
                        <option value="">手动填写</option>
                        <option value="now">自动：当前时间</option>
                        <option value="user">自动：当前用户</option>
                        <option value="autonum">自动：编号 + 1</option>
                      </select>
                    </div>
                  )}
                </div>
              ))}
            </div>
            <div className="fd-extra">
              <label>
                提交后
                <select className="ed-select" value={after} onChange={(e) => setAfter(e.target.value as 'continue' | 'close')}>
                  <option value="continue">留在表单里继续新建</option>
                  <option value="close">关闭表单</option>
                </select>
              </label>
              {scripts.length > 0 && (
                <label title="提交前运行：脚本返回错误信息就不写入（只读，不能改包里的内容）">
                  提交前校验脚本
                  <select className="ed-select" value={onValidate} onChange={(e) => setOnValidate(e.target.value)}>
                    <option value="">不校验</option>
                    {scripts.map((s) => (
                      <option key={s.id} value={s.id}>
                        {s.title}
                      </option>
                    ))}
                  </select>
                </label>
              )}
              {scripts.length > 0 && (
                <label>
                  提交后运行脚本
                  <select className="ed-select" value={onSubmit} onChange={(e) => setOnSubmit(e.target.value)}>
                    <option value="">不运行</option>
                    {scripts.map((s) => (
                      <option key={s.id} value={s.id}>
                        {s.title}
                      </option>
                    ))}
                  </select>
                </label>
              )}
            </div>
          </div>
          <div className="fd-right" aria-label="预览">
            <div className="fd-prev-title">预览 · {title || '表单'}</div>
            {previewRows.length === 0 && <div className="sg-empty">还没有字段</div>}
            {previewRows.map((r) => (
              <div key={r.col} className="ff-field">
                <div className="ff-lbl">
                  <span>
                    {r.label.trim() || r.title}
                    {r.meta?.required && <span className="req"> *</span>}
                  </span>
                  <span className="ff-type">{r.meta ? COLUMN_TYPE_LABELS[r.meta.type] : '文本'}</span>
                  {r.auto && <span className="ff-tag">自动</span>}
                  {r.keep && <span className="ff-tag">沿用</span>}
                </div>
                <div className="ff-auto">{r.auto ? '提交时自动填' : r.meta?.default ? `默认：${r.meta.default}` : ' '}</div>
                {(r.hint || r.meta?.note) && <span className="ff-hint">{r.hint || r.meta?.note}</span>}
              </div>
            ))}
          </div>
        </div>
        <div className="sg-foot">
          {onDelete && (
            <button
              className="btn ghost"
              onClick={async () => {
                if ((await host().confirm(`删除表单「${form?.title}」吗？`))) onDelete();
              }}
            >
              删除表单
            </button>
          )}
          <span className="grow" />
          <button className="btn" onClick={onClose}>
            取消
          </button>
          <button className="btn primary" onClick={save}>
            保存表单
          </button>
        </div>
      </div>
    </div>
  );
}

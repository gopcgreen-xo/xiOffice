import { useEffect, useRef, useState } from 'react';
import { Dialog } from './WordDialogs';
import { PT_PER_CM } from '../core/model';
import { imageElement, loadImages, registerImage } from '../core/images';
import { drawPicture, pictureBounds, pictureFormatError, type PictureFormat, type PictureWrap } from '../core/picture';

const WRAPS: [PictureWrap, string][] = [['inline', '嵌入型'], ['square', '四周型'], ['topAndBottom', '上下型'], ['front', '浮于文字上方'], ['behind', '衬于文字下方']];
const round = (v: number) => Math.round(v * 10000) / 10000;

export function PictureDialog({ value, onApply, onClose }: { value: PictureFormat; onApply: (v: PictureFormat) => void; onClose: () => void }) {
  const [form, setForm] = useState<PictureFormat>(() => structuredClone(value));
  const [locked, setLocked] = useState(true);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const preview = useRef<HTMLCanvasElement>(null);
  const mounted = useRef(true);
  useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
  useEffect(() => {
    let active = true;
    void loadImages([form.image.id]).then(() => {
      const canvas = preview.current;
      if (!active || !canvas) return;
      const ctx = canvas.getContext('2d');
      const source = imageElement(form.image.id);
      if (!ctx) return;
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      if (!source) return;
      const bounds = pictureBounds(form.image);
      const scale = Math.min(1.5, (canvas.width - 20) / bounds.w, (canvas.height - 20) / bounds.h);
      if (!(scale > 0 && Number.isFinite(scale))) return;
      ctx.save(); ctx.translate(canvas.width / 2, canvas.height / 2); ctx.scale(scale, scale);
      drawPicture(ctx, source, form.image, -form.image.w / 2, -form.image.h / 2);
      ctx.restore();
    });
    return () => { active = false; };
  }, [form.image]);
  const size = (key: 'w' | 'h', cm: number) => setForm((f) => {
    const other = key === 'w' ? 'h' : 'w';
    const n = cm * PT_PER_CM;
    return { ...f, image: { ...f.image, [key]: n, ...(locked && f.image[key] > 0 ? { [other]: f.image[other] * n / f.image[key] } : {}) } };
  });
  const crop = form.image.crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const setCrop = (key: keyof typeof crop, percent: number) => setForm((f) => {
    const old = f.image.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const next = { ...old, [key]: percent / 100 };
    const w = f.image.w * (1 - next.l - next.r) / (1 - old.l - old.r);
    const h = f.image.h * (1 - next.t - next.b) / (1 - old.t - old.b);
    if (!(w > 0 && h > 0)) { setError('裁剪后必须保留可见的图片区域'); return f; }
    setError('');
    return { ...f, image: { ...f.image, crop: next, w, h } };
  });
  const replace = async (file?: File) => {
    if (!file) return;
    if (!/^image\/(png|jpeg|gif|bmp|webp)$/.test(file.type)) return setError('请选择 PNG、JPEG、GIF、BMP 或 WebP 图片');
    setBusy(true); setError('');
    try {
      const id = registerImage(new Uint8Array(await file.arrayBuffer()), file.type);
      await loadImages([id]);
      if (!imageElement(id)) throw new Error('无法读取这张图片');
      if (mounted.current) setForm((f) => ({ ...f, image: { ...f.image, id } }));
    } catch (e) { if (mounted.current) setError(String(e instanceof Error ? e.message : e)); }
    finally { if (mounted.current) setBusy(false); }
  };
  const numeric = (label: string, v: number, change: (n: number) => void, unit = '厘米') => (
    <label className="label">{label}（{unit}）<input className="field" type="number" step="0.01" value={round(v)} onChange={(e) => change(Number(e.target.value))} /></label>
  );
  const invalid = pictureFormatError(form);
  return (
    <Dialog title="设置图片格式" width={600} onClose={onClose} foot={<><button className="btn" onClick={onClose}>取消</button><button className="btn primary" disabled={busy || !!invalid} onClick={() => onApply(form)}>确定</button></>}>
      <div style={{ display: 'grid', gap: 12, maxHeight: '70vh', overflowY: 'auto' }}>
        <canvas ref={preview} width={540} height={150} aria-label="图片预览" style={{ width: '100%', height: 150, objectFit: 'contain', background: '#fff', border: '1px solid var(--border)' }} />
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
          {numeric('宽度', form.image.w / PT_PER_CM, (n) => size('w', n))}
          {numeric('高度', form.image.h / PT_PER_CM, (n) => size('h', n))}
        </div>
        <label><input type="checkbox" checked={locked} onChange={(e) => setLocked(e.target.checked)} /> 锁定纵横比</label>
        <div style={{ display: 'flex', gap: 8, alignItems: 'end', flexWrap: 'wrap' }}>
          {numeric('旋转', form.image.rot ?? 0, (rot) => setForm({ ...form, image: { ...form.image, rot } }), '度')}
          <button className="btn" onClick={() => setForm({ ...form, image: { ...form.image, rot: ((form.image.rot ?? 0) + 90) % 360 } })}>向右旋转 90°</button>
          <button className="btn" aria-pressed={!!form.image.flipH} onClick={() => setForm({ ...form, image: { ...form.image, flipH: !form.image.flipH } })}>水平翻转</button>
          <button className="btn" aria-pressed={!!form.image.flipV} onClick={() => setForm({ ...form, image: { ...form.image, flipV: !form.image.flipV } })}>垂直翻转</button>
        </div>
        <fieldset><legend>裁剪</legend><div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
          {([['l', '左侧'], ['r', '右侧'], ['t', '上方'], ['b', '下方']] as const).map(([k, label]) => <div key={k}>{numeric(label, crop[k] * 100, (n) => setCrop(k, n), '%')}</div>)}
        </div></fieldset>
        <label className="label">文字环绕<select className="field" value={form.wrap} onChange={(e) => setForm({ ...form, wrap: e.target.value as PictureWrap })}>
          {WRAPS.map(([v, label]) => <option key={v} value={v}>{label}</option>)}
          {['tight', 'through'].includes(value.wrap) && <option value={value.wrap}>{value.wrap === 'tight' ? '紧密型（原文档）' : '穿越型（原文档）'}</option>}
        </select></label>
        {form.wrap !== 'inline' && <fieldset><legend>位置与文字距离</legend><div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
          {(['posH', 'posV'] as const).map((axis) => <div key={axis} style={{ display: 'grid', gap: 6 }}>
            <label className="label">{axis === 'posH' ? '水平参照' : '垂直参照'}<select className="field" value={form[axis].rel} onChange={(e) => setForm({ ...form, [axis]: { ...form[axis], rel: e.target.value, align: undefined } })}>
              <option value="page">页面</option><option value="margin">页边距</option><option value={axis === 'posH' ? 'column' : 'paragraph'}>{axis === 'posH' ? '栏' : '段落'}</option>
              {!['page', 'margin', axis === 'posH' ? 'column' : 'paragraph'].includes(form[axis].rel) && <option value={form[axis].rel}>原文档参照</option>}
            </select></label>
            {numeric(axis === 'posH' ? '水平偏移' : '垂直偏移', form[axis].offset / PT_PER_CM, (n) => setForm({ ...form, [axis]: { ...form[axis], offset: n * PT_PER_CM, align: undefined } }))}
            {form[axis].align && <small>当前按 {form[axis].align} 对齐；修改偏移后改用绝对位置。</small>}
          </div>)}
          {([['l', '左距'], ['r', '右距'], ['t', '上距'], ['b', '下距']] as const).map(([k, label]) => <div key={k}>{numeric(label, form.dist[k] / PT_PER_CM, (n) => setForm({ ...form, dist: { ...form.dist, [k]: n * PT_PER_CM } }))}</div>)}
        </div></fieldset>}
        <label className="btn" style={{ alignSelf: 'start', cursor: 'pointer' }}>更改图片…<input type="file" hidden disabled={busy} accept="image/png,image/jpeg,image/gif,image/bmp,image/webp" onChange={(e) => { void replace(e.target.files?.[0]); e.target.value = ''; }} /></label>
        {(error || invalid) && <div role="alert" style={{ color: 'var(--danger, #d33)' }}>{error || invalid}</div>}
      </div>
    </Dialog>
  );
}

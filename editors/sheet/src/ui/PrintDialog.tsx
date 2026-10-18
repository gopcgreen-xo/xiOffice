/**
 * 打印 / 导出 PDF：左边页面设置（范围、方向、纸张、缩放、标题行、网格线），右边分页预览。
 * 设置会写进工作表的页面设置（保存后 Excel 里打印也一样）。
 */
import { useEffect, useMemo, useState } from 'react';
import { host } from '@wz/editor-shared';
import { layoutPrint, PAPERS, renderPage } from '../core/print';
import { DEFAULT_MARGINS, type CellStyle, type PageMargins, type PageSetup, type Range, type SheetData } from '../core/workbook';
import { colName } from '@wz/editor-shared/xlsx';
import { Icon } from '@wz/editor-shared/Icon';

type Fit = 'none' | 'page' | 'width' | 'height' | 'custom';

interface Props {
  sheet: SheetData;
  styles: CellStyle[];
  defSize: number;
  defFont: string;
  sel: Range;
  /** 选区是多格 */
  multi: boolean;
  onSavePage: (page: PageSetup) => void;
  onPrint: (urls: string[], w: number, h: number) => void;
  onPdf: (urls: string[], w: number, h: number) => void;
  onClose: () => void;
  /** 页眉页脚里的文件名 */
  fileName?: string;
}

/** 常用的页眉页脚（Excel 代码） */
const HF_PRESETS: [string, string][] = [
  ['（无）', ''],
  ['第 1 页', '&C第 &P 页'],
  ['第 1 页，共 ? 页', '&C第 &P 页，共 &N 页'],
  ['表名', '&C&A'],
  ['文件名', '&C&F'],
  ['日期（右）', '&R&D'],
  ['文件名 + 页码', '&L&F&R第 &P 页，共 &N 页'],
];
const CM_PER_INCH = 2.54;
const cm = (inch: number) => Math.round(inch * CM_PER_INCH * 100) / 100;

const rangeText = (g: Range) => `${colName(g.c0)}${g.r0 + 1}:${colName(g.c1)}${g.r1 + 1}`;

export function PrintDialog(p: Props) {
  const [page, setPage] = useState<PageSetup>(() => ({ paper: 9, ...p.sheet.page }));
  const [scope, setScope] = useState<'sheet' | 'selection' | 'area'>(p.sheet.page?.area ? 'area' : p.multi ? 'selection' : 'sheet');
  const [titles, setTitles] = useState(page.titleRows ? `${page.titleRows[0] + 1}:${page.titleRows[1] + 1}` : '');
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    host().pushPopup();
    return () => host().popPopup();
  }, []);

  const fit: Fit = page.fitW !== undefined || page.fitH !== undefined ? (page.fitW === 1 && page.fitH === 1 ? 'page' : page.fitW === 1 ? 'width' : 'height') : page.scale && page.scale !== 100 ? 'custom' : 'none';
  const setFit = (f: Fit) => {
    const { fitW: _w, fitH: _h, scale: _s, ...rest } = page;
    if (f === 'page') setPage({ ...rest, fitW: 1, fitH: 1 });
    else if (f === 'width') setPage({ ...rest, fitW: 1, fitH: 0 });
    else if (f === 'height') setPage({ ...rest, fitW: 0, fitH: 1 });
    else if (f === 'custom') setPage({ ...rest, scale: page.scale ?? 90 });
    else setPage(rest);
  };
  const area = scope === 'selection' ? p.sel : scope === 'area' ? page.area : undefined;
  const effective: PageSetup = { ...page, area: scope === 'sheet' ? undefined : page.area };
  const layout = useMemo(() => layoutPrint(p.sheet, effective, area), [p.sheet, JSON.stringify(effective), area?.r0, area?.c0, area?.r1, area?.c1]); // eslint-disable-line react-hooks/exhaustive-deps
  const opts = { styles: p.styles, defSize: p.defSize, defFont: p.defFont, gridLines: page.gridLines, sheetName: p.sheet.name, fileName: p.fileName };
  const margins = page.margins ?? DEFAULT_MARGINS;
  const setMargin = (k: keyof PageMargins, v: string) => {
    const n = Number(v);
    if (!Number.isFinite(n) || n < 0 || n > 20) return;
    setPage({ ...page, margins: { ...margins, [k]: Math.round((n / CM_PER_INCH) * 1000) / 1000 } });
  };
  const hfField = (key: 'header' | 'footer', label: string) => {
    const v = page[key] ?? '';
    const preset = HF_PRESETS.some(([, c]) => c === v) ? v : 'custom';
    return (
      <>
        <label>
          <span className="sub">{label}</span>
          <select
            className="ed-select"
            value={preset}
            onChange={(e) => {
              const { [key]: _x, ...rest } = page;
              const c = e.target.value === 'custom' ? v || '&C' : e.target.value;
              setPage(c ? { ...rest, [key]: c } : rest);
            }}
          >
            {HF_PRESETS.map(([l, c]) => (
              <option key={l} value={c}>
                {l}
              </option>
            ))}
            <option value="custom">自定义…</option>
          </select>
        </label>
        {preset === 'custom' && (
          <label>
            <span className="sub" />
            <input
              title="&L 左 &C 中 &R 右；&P 页码 &N 总页数 &A 表名 &F 文件名 &D 日期 &T 时间"
              value={v}
              onChange={(e) => {
                const { [key]: _x, ...rest } = page;
                setPage(e.target.value ? { ...rest, [key]: e.target.value } : rest);
              }}
            />
          </label>
        )}
      </>
    );
  };
  const previews = useMemo(
    () => layout.pages.slice(0, 12).map((pg) => renderPage(p.sheet, layout, pg, { ...opts, pxPerPt: 0.5 }).toDataURL('image/png')),
    [layout, page.gridLines, page.header, page.footer], // eslint-disable-line react-hooks/exhaustive-deps
  );

  const commitTitles = (v: string) => {
    setTitles(v);
    const m = v.trim().replace(/\$/g, '').match(/^(\d+)(?::(\d+))?$/);
    const { titleRows: _t, ...rest } = page;
    if (!v.trim()) return setPage(rest);
    if (m) {
      const a = Number(m[1]) - 1;
      const b = Number(m[2] ?? m[1]) - 1;
      setPage({ ...rest, titleRows: [Math.min(a, b), Math.max(a, b)] });
    }
  };
  const save = () => {
    const { area: _a, ...rest } = page;
    const next: PageSetup = scope === 'area' && page.area ? page : rest;
    if (JSON.stringify(next) !== JSON.stringify(p.sheet.page ?? { paper: 9 })) p.onSavePage(next);
  };
  const render = async (fn: (urls: string[], w: number, h: number) => void) => {
    setBusy(true);
    save();
    // 先让“正在准备”显示出来
    await new Promise((r) => setTimeout(r, 30));
    const urls = layout.pages.map((pg) => renderPage(p.sheet, layout, pg, { ...opts, pxPerPt: 2 }).toDataURL('image/png'));
    setBusy(false);
    fn(urls, layout.w, layout.h);
  };

  return (
    <div className="sheet-detail-mask" onMouseDown={p.onClose}>
      <div className="sheet-detail sheet-print" role="dialog" aria-label="打印" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.key === 'Escape' && p.onClose()}>
        <div className="sheet-detail-head">
          <strong>打印</strong>
          <span className="sub">
            共 {layout.pages.length} 页{layout.scale < 1 ? ` · 缩放 ${Math.round(layout.scale * 100)}%` : ''}
          </span>
          <span className="grow" />
          <button className="icon-btn" aria-label="关闭" onClick={p.onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        <div className="sheet-print-body">
          <div className="sheet-print-form">
            <label>
              <span className="sub">打印范围</span>
              <select className="ed-select" value={scope} onChange={(e) => setScope(e.target.value as typeof scope)}>
                <option value="sheet">整个工作表</option>
                <option value="selection" disabled={!p.multi}>
                  选定区域{p.multi ? `（${rangeText(p.sel)}）` : ''}
                </option>
                <option value="area" disabled={!page.area}>
                  打印区域{page.area ? `（${rangeText(page.area)}）` : ''}
                </option>
              </select>
            </label>
            <div className="sheet-print-btns">
              <button className="btn" disabled={!p.multi} onClick={() => (setPage({ ...page, area: p.sel }), setScope('area'))}>
                选区设为打印区域
              </button>
              <button className="btn" disabled={!page.area} onClick={() => (setPage({ ...page, area: undefined }), setScope('sheet'))}>
                清除
              </button>
            </div>
            <label>
              <span className="sub">方向</span>
              <select className="ed-select" value={page.landscape ? 'l' : 'p'} onChange={(e) => setPage({ ...page, landscape: e.target.value === 'l' })}>
                <option value="p">纵向</option>
                <option value="l">横向</option>
              </select>
            </label>
            <label>
              <span className="sub">纸张</span>
              <select className="ed-select" value={page.paper ?? 9} onChange={(e) => setPage({ ...page, paper: Number(e.target.value) })}>
                {Object.entries(PAPERS).map(([id, x]) => (
                  <option key={id} value={id}>
                    {x.name}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span className="sub">缩放</span>
              <select className="ed-select" value={fit} onChange={(e) => setFit(e.target.value as Fit)}>
                <option value="none">无缩放</option>
                <option value="page">将工作表调整为一页</option>
                <option value="width">将所有列调整为一页</option>
                <option value="height">将所有行调整为一页</option>
                <option value="custom">自定义缩放</option>
              </select>
            </label>
            {fit === 'custom' && (
              <label>
                <span className="sub">缩放比例 %</span>
                <input type="number" min={10} max={400} value={page.scale ?? 100} onChange={(e) => setPage({ ...page, scale: Math.max(10, Math.min(400, Number(e.target.value) || 100)) })} />
              </label>
            )}
            <label>
              <span className="sub">顶端标题行</span>
              <input placeholder="如 1:2，每页都打印" value={titles} onChange={(e) => commitTitles(e.target.value)} />
            </label>
            <label className="sheet-print-check">
              <input type="checkbox" checked={!!page.gridLines} onChange={(e) => setPage({ ...page, gridLines: e.target.checked || undefined })} />
              打印网格线
            </label>
            <div className="sheet-print-margins">
              <span className="sub">页边距（厘米）</span>
              {(
                [
                  ['top', '上'],
                  ['bottom', '下'],
                  ['left', '左'],
                  ['right', '右'],
                ] as [keyof PageMargins, string][]
              ).map(([k, l]) => (
                <label key={k}>
                  <span className="sub">{l}</span>
                  <input type="number" step={0.1} min={0} value={cm(margins[k])} onChange={(e) => setMargin(k, e.target.value)} />
                </label>
              ))}
            </div>
            <label className="sheet-print-check">
              <input type="checkbox" checked={!!page.centerH} onChange={(e) => setPage({ ...page, centerH: e.target.checked || undefined })} />
              水平居中
              <input type="checkbox" checked={!!page.centerV} onChange={(e) => setPage({ ...page, centerV: e.target.checked || undefined })} />
              垂直居中
            </label>
            {hfField('header', '页眉')}
            {hfField('footer', '页脚')}
            {!!(page.rowBreaks?.length || page.colBreaks?.length) && (
              <div className="sheet-print-btns">
                <span className="sub grow">
                  有 {(page.rowBreaks?.length ?? 0) + (page.colBreaks?.length ?? 0)} 个手动分页符
                </span>
                <button
                  className="btn"
                  onClick={() => {
                    const { rowBreaks: _r, colBreaks: _c, ...rest } = page;
                    setPage(rest);
                  }}
                >
                  全部删除
                </button>
              </div>
            )}
          </div>
          <div className="sheet-print-preview">
            {previews.map((u, i) => (
              <figure key={i}>
                <img src={u} alt={`第 ${i + 1} 页`} />
                <figcaption className="sub">第 {i + 1} 页</figcaption>
              </figure>
            ))}
            {layout.pages.length > previews.length && <div className="sub">… 还有 {layout.pages.length - previews.length} 页</div>}
          </div>
        </div>
        <div className="sheet-detail-foot">
          <span className="sub grow">{busy ? '正在准备…' : '页面设置会随表格一起保存'}</span>
          <button className="btn" onClick={p.onClose}>
            取消
          </button>
          <button className="btn" disabled={busy} onClick={() => void render(p.onPdf)}>
            导出 PDF
          </button>
          <button className="btn primary" disabled={busy} onClick={() => void render(p.onPrint)}>
            <Icon name="printer" size={14} />
            打印
          </button>
        </div>
      </div>
    </div>
  );
}

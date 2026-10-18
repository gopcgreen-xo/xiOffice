/**
 * 工作表里的图表（只读）：柱形、条形、折线、面积、饼图，按当前单元格的值画（引用取不到时用文件里的缓存值）。
 * 其它类型显示标题和“图表（保存时保留）”。
 */
import type { ChartSpec } from '../core/drawing';
import { refCells } from '../core/charts';
import type { Workbook } from '../core/workbook';

const COLORS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];

export function SheetChart({ spec, w, h, book }: { spec: ChartSpec; w: number; h: number; book: Workbook }) {
  const series = spec.series.map((s) => {
    const vals = refCells(book, s.valRef);
    const cats = refCells(book, s.catRef);
    return { name: s.name, vals: vals ? vals.map((x) => x.v ?? 0) : s.vals, cats: cats ? cats.map((x) => x.t) : s.cats };
  });
  const cats = series[0]?.cats.length ? series[0].cats : (series[0]?.vals.map((_, i) => String(i + 1)) ?? []);
  const top = spec.title ? 28 : 10;
  const legendH = series.length > 1 || spec.type === 'pie' ? 22 : 0;
  const pad = { l: 44, r: 12, t: top, b: 24 + legendH };
  const pw = Math.max(10, w - pad.l - pad.r);
  const ph = Math.max(10, h - pad.t - pad.b);
  const all = series.flatMap((s) => s.vals);
  const max = Math.max(0, ...all);
  const min = Math.min(0, ...all);
  const span = max - min || 1;
  const Y = (v: number) => pad.t + ph - ((v - min) / span) * ph;
  const fmt = (v: number) => (Math.abs(v) >= 10000 ? `${(v / 10000).toFixed(1)}万` : String(Math.round(v * 100) / 100));
  const body = () => {
    if (spec.type === 'other' || !series.length) {
      return (
        <text x={w / 2} y={h / 2} textAnchor="middle" fontSize="12" fill="#888">
          图表（保存时保留）
        </text>
      );
    }
    if (spec.type === 'pie') {
      const vals = series[0].vals.map((v) => Math.max(0, v));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = w / 2;
      const cy = pad.t + ph / 2;
      const rad = Math.max(5, Math.min(pw, ph) / 2 - 4);
      let a0 = -Math.PI / 2;
      return (
        <>
          {vals.map((v, i) => {
            const a1 = a0 + (v / total) * Math.PI * 2;
            const large = a1 - a0 > Math.PI ? 1 : 0;
            const d = `M${cx},${cy} L${cx + rad * Math.cos(a0)},${cy + rad * Math.sin(a0)} A${rad},${rad} 0 ${large} 1 ${cx + rad * Math.cos(a1)},${cy + rad * Math.sin(a1)} Z`;
            a0 = a1;
            return <path key={i} d={d} fill={COLORS[i % COLORS.length]} stroke="#fff" strokeWidth="1" />;
          })}
          {cats.slice(0, 8).map((c, i) => (
            <g key={c + i} transform={`translate(${8 + (i % 4) * ((w - 16) / 4)}, ${h - 14 - (i >= 4 ? 0 : 0)})`}>
              <rect width="8" height="8" y="-7" fill={COLORS[i % COLORS.length]} />
              <text x="11" fontSize="10" fill="#444">
                {c}
              </text>
            </g>
          ))}
        </>
      );
    }
    const n = cats.length || 1;
    const grid = [0, 0.25, 0.5, 0.75, 1].map((k) => min + span * k);
    const axis = (
      <>
        {grid.map((v) => (
          <g key={v}>
            <line x1={pad.l} x2={pad.l + pw} y1={Y(v)} y2={Y(v)} stroke="#e3e3e3" />
            <text x={pad.l - 4} y={Y(v) + 3} textAnchor="end" fontSize="9" fill="#777">
              {fmt(v)}
            </text>
          </g>
        ))}
        {cats.map((c, i) =>
          n <= 12 || i % Math.ceil(n / 12) === 0 ? (
            <text key={i} x={pad.l + (pw / n) * (i + 0.5)} y={pad.t + ph + 13} textAnchor="middle" fontSize="9" fill="#555">
              {c.length > 6 ? `${c.slice(0, 6)}…` : c}
            </text>
          ) : null,
        )}
      </>
    );
    if (spec.type === 'col' || spec.type === 'bar') {
      const gw = pw / n;
      const bw = (gw * 0.7) / series.length;
      return (
        <>
          {axis}
          {series.map((s, si) =>
            s.vals.map((v, i) => (
              <rect key={`${si}-${i}`} x={pad.l + gw * i + gw * 0.15 + bw * si} y={Math.min(Y(v), Y(0))} width={Math.max(1, bw - 1)} height={Math.abs(Y(v) - Y(0))} fill={COLORS[si % COLORS.length]} />
            )),
          )}
        </>
      );
    }
    return (
      <>
        {axis}
        {series.map((s, si) => {
          const pts = s.vals.map((v, i) => `${pad.l + (pw / n) * (i + 0.5)},${Y(v)}`).join(' ');
          return spec.type === 'area' ? (
            <polygon key={si} points={`${pad.l + (pw / n) * 0.5},${Y(0)} ${pts} ${pad.l + (pw / n) * (s.vals.length - 0.5)},${Y(0)}`} fill={COLORS[si % COLORS.length]} fillOpacity="0.5" />
          ) : (
            <polyline key={si} points={pts} fill="none" stroke={COLORS[si % COLORS.length]} strokeWidth="2" />
          );
        })}
      </>
    );
  };
  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} className="sheet-chart">
      <rect x="0.5" y="0.5" width={w - 1} height={h - 1} fill="#fff" stroke="#d0d0d0" />
      {spec.title && (
        <text x={w / 2} y="19" textAnchor="middle" fontSize="13" fontWeight="600" fill="#333">
          {spec.title}
        </text>
      )}
      {body()}
      {series.length > 1 && spec.type !== 'pie' && (
        <g>
          {series.map((s, i) => (
            <g key={i} transform={`translate(${pad.l + i * 90}, ${h - 8})`}>
              <rect width="8" height="8" y="-7" fill={COLORS[i % COLORS.length]} />
              <text x="11" fontSize="10" fill="#444">
                {s.name || `系列${i + 1}`}
              </text>
            </g>
          ))}
        </g>
      )}
    </svg>
  );
}

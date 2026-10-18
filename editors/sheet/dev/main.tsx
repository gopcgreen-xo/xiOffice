/** 开发页：不启动 Tauri，在浏览器里直接跑表格编辑器（`npm run dev:sheet`）；文件系统是内存里的假实现。
 *  不带参数时打开一份现做的示例（含行 / 列分组、动态数组）；?file=<本机完整路径> 打开真实文件。 */
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { addFilesToBundle, itemPath, loadBundle, newBundleFile, setEditorHost, withBundles, type EditorHost } from '@wz/editor-shared';
import '../../../src/styles/themes.css';
import '../../../src/styles/app.css';
import { Editor } from '../src';
import { groupRange, toggleGroup } from '../src/core/outline';
import { setInput, Tx } from '../src/core/ops';
import { saveXlsx } from '../src/core/save';
import { Engine } from '../src/core/formula/eval';
import { blankWorkbook } from '../src/core/workbook';

const noop = () => undefined;
const file = new URLSearchParams(location.search).get('file');

/** 示例：三个季度各 3 个月明细（行分组，汇总行在下），月份列分组，F 列起一个 UNIQUE */
function sample(): Uint8Array {
  const b = blankWorkbook();
  const tx = new Tx(b, '示例');
  const set = (r: number, c: number, v: string | number) => setInput(tx, 0, r, c, String(v));
  ['项目', '1 月', '2 月', '3 月', '合计'].forEach((t, c) => set(0, c, t));
  let r = 1;
  for (const q of ['一', '二', '三']) {
    for (let m = 1; m <= 3; m++, r++) {
      set(r, 0, `${q}季度-${m}`);
      for (let c = 1; c <= 3; c++) set(r, c, r * c * 10);
      set(r, 4, `=SUM(B${r + 1}:D${r + 1})`);
    }
    set(r, 0, `${q}季度合计`);
    for (let c = 1; c <= 4; c++) set(r, c, `=SUM(${String.fromCharCode(65 + c)}${r - 2}:${String.fromCharCode(65 + c)}${r})`);
    r++;
  }
  ['苹果', '香蕉', '苹果', '梨'].forEach((t, i) => set(i, 6, t));
  set(0, 8, '=UNIQUE(G1:G4)');
  tx.commit();
  new Engine(b).analyze();
  for (const [a, z] of [[1, 3], [5, 7], [9, 11]]) groupRange(new Tx(b, 'g'), 0, 'row', a, z);
  groupRange(new Tx(b, 'g'), 0, 'col', 1, 3);
  toggleGroup(new Tx(b, 'g'), 0, 'row', 8);
  return saveXlsx(b);
}

/** 工作包模式（?pack=1：空白列属性，?pack=2：预置列属性）：内存里建一个包，里面放一张采购明细表 */
const packMode = new URLSearchParams(location.search).get('pack');
const mem = new Map<string, Uint8Array>();

function procurement(): Uint8Array {
  const b = blankWorkbook();
  const tx = new Tx(b, '示例');
  const rows = [
    ['物资名称', '规格型号', '单位', '数量', '单价', '到货日期', '供应商', '已验收', '随货资料', '备注'],
    ['法兰', 'DN100 PN16', '个', '24', '86', '2026-10-12', '供应商甲', '是', '合格证,材质证明', ''],
    ['截止阀', 'J41H-16C DN50', '台', '6', '1250', '2026-10-15', '供应商乙', '否', '合格证', '加急'],
    ['不锈钢管', '316L Φ57×3.5', '米', '120', '145.5', '2026-10-18', '供应商丙', '否', '材质证明', ''],
    ['工业电缆', 'CEFR 3×2.5', '米', '800', '12.8', '2026-10-20', '供应商丁', '是', '合格证,材质证明', '分两批到货'],
    ['密封垫片', '石棉橡胶 DN100', '片', '50', '6.5', '2026-10-12', '供应商甲', '是', '合格证', ''],
    ['室内灯具', 'CFD3-1 24V', '套', '18', '320', '2026-10-25', '供应商戊', '否', '', ''],
    ['蝶阀', 'D71X-10 DN150', '台', '4', '980', '2026-10-28', '供应商乙', '否', '合格证,检验报告', '附合格证'],
    ['不锈钢螺栓', 'M16×60 A4-70', '套', '300', '3.2', '2026-10-12', '供应商戊', '是', '合格证', ''],
    ['镀锌钢管', 'DN80 SCH40', '米', '60', '88', '2026-10-30', '供应商丙', '否', '材质证明', ''],
    ['防火涂料', '膨胀型 25kg', '桶', '12', '560', '2026-11-02', '供应商甲', '否', '检验报告', ''],
    ['压力表', 'Y-100 0-1.6MPa', '块', '10', '75', '待定', '供应商乙', '是', '合格证', ''],
    ['工业插座', 'CZK 16A IP56', '个', '40', '48', '2026-10-25', '供应商丁', '否', '', ''],
  ];
  rows.forEach((row, r) => row.forEach((v, c) => setInput(tx, 0, r, c, v)));
  tx.commit();
  return saveXlsx(b);
}

const host = {
  readBytes: async (p: string) => (mem.has(p) ? mem.get(p)! : file ? new Uint8Array(await (await fetch(`/@fs/${p.split('\\').join('/')}`)).arrayBuffer()) : sample()),
  writeBytes: async (p: string, d: Uint8Array) => void mem.set(p, d),
  writeText: async () => undefined,
  readText: async (p: string) => ({ content: new TextDecoder().decode(mem.get(p) ?? new Uint8Array()), encoding: 'UTF-8' }),
  listDir: async () => [],
  readXls: async () => [],
  openWithSystem: async () => undefined,
  openUrl: async () => undefined,
  workFolders: () => ['C:/dev'],
  toast: (m: string) => console.log('[toast]', m),
  askText: async (_t: string, v?: string) => v ?? '',
  confirm: async () => true,
  guard: async <T,>(p: Promise<T>) => p,
  pushPopup: noop,
  popPopup: noop,
  updateTab: noop,
  printImages: async () => undefined,
  log: noop,
  useWorkFolders: () => ['C:/dev'],
  useFsEvent: () => ({ tick: 0, dirs: [] }),
  useUserName: () => '开发者',
  pickFiles: async () => [],
} as unknown as EditorHost;
setEditorHost(packMode ? withBundles(host) : host);

let tabPath = file ?? 'C:/dev/示例.xlsx';
if (packMode) {
  mem.set('C:/dev/采购明细.xlsx', procurement());
  const pack = 'C:/dev/采购.wbp';
  await newBundleFile(pack, '采购');
  const [item] = await addFilesToBundle(pack, ['C:/dev/采购明细.xlsx']);
  tabPath = itemPath(pack, item);
  if (packMode === '2') {
    const b = await loadBundle(pack);
    const opt = (...v: string[]) => v.map((x) => ({ v: x }));
    b.setColumns(item.id, {
      v: 1,
      sheets: {
        Sheet1: {
          header: 1,
          columns: [
            { id: 'c1', title: '物资名称', col: 0, type: 'text', required: true },
            { id: 'c2', title: '单位', col: 2, type: 'select', options: opt('个', '台', '米', '片', '套', '桶', '块'), default: '个' },
            { id: 'c3', title: '数量', col: 3, type: 'number', integer: true, min: 0, required: true },
            { id: 'c4', title: '单价', col: 4, type: 'money', min: 0 },
            { id: 'c5', title: '到货日期', col: 5, type: 'date', default: '{{today+7}}' },
            { id: 'c6', title: '供应商', col: 6, type: 'select', options: opt('供应商甲', '供应商乙', '供应商丙', '供应商丁', '供应商戊') },
            { id: 'c7', title: '已验收', col: 7, type: 'check', default: '否' },
            { id: 'c8', title: '随货资料', col: 8, type: 'multi', options: [{ v: '合格证', color: '#10b981' }, { v: '材质证明', color: '#3f5ef0' }, { v: '检验报告', color: '#f59e0b' }], required: true },
          ],
        },
      },
    });
    b.saveForm({ id: 'f1', title: '到货登记', target: { item: item.id, sheet: 'Sheet1' }, fields: [{ column: 'c1' }, { column: 'c2', keep: true }, { column: 'c3' }, { column: 'c5' }, { column: 'c6', keep: true }, { column: 'c8' }, { column: 'c7' }], after: 'continue' });
  }
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <Editor tab={{ id: 'dev', title: file ? file.split(/[\\/]/).pop()! : '示例.xlsx', path: tabPath }} />
  </StrictMode>,
);

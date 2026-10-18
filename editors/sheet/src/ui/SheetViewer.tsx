import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import { AUTOSAVE_MS, type EditorTab, basename, clearDraft, dirname, draftTime, extOf, host, loadDraft, pickSave, saveDraft } from '@wz/editor-shared';
import { COLUMN_TYPE_GLYPHS, COLUMN_TYPE_LABELS, columnAt, formRequestKey, onFormRequest, takeFormRequest, isCheckOn, optionColor, splitMulti, type BundleContext, type ColumnMeta, type FormDef, type SelectOption } from '@wz/editor-shared';
import { flushSync } from 'react-dom';
import { colName } from '@wz/editor-shared/xlsx';
import { bindColumns, cellText as colCellText, checkCell, metaAt, missingRequired, optionList } from '../core/colmeta';
import { applyInferences, fillDefaults, rowIsBlank, setColumn, setColumnInput, setHeaderRow } from '../core/columns';
import { buildFormHtml } from '../core/formHtml';
import { importFormResult, parseFormResult, resolveForm, submitForm, type FormTarget, type FormValues, type ResolvedField } from '../core/forms';
import { ColumnPanel } from './ColumnPanel';
import { ColumnPicker } from './ColumnPicker';
import { FormDesigner } from './FormDesigner';
import { FormFill } from './FormFill';
import { SmartColumnsDialog } from './SmartColumnsDialog';
import './columns.css';
import { quoteSheet } from '../core/formula/refs';
import { parseInWorker } from '../core/parse';
import { fromXls } from '../core/xls';
import type { CellBorder } from '../core/color';
import { blankWorkbook, parseRef, styleAt, type Range, type SCell, type SheetData, type Workbook } from '../core/workbook';
import { formatCode, defaultFontName, defaultFontSize } from '../core/styles';
import {
  applyBorders,
  applyEntry,
  applyStyle,
  blockFromText,
  clearRange,
  copyBlock,
  eachCell,
  growRows,
  insertCopied,
  insertCut,
  shiftRange,
  insertDelete,
  cellShiftBlocked,
  violatesProtection,
  lostByMerge,
  mergeAcross,
  mergeCells,
  pasteBlock,
  setColWidth,
  setFreeze,
  setInput,
  setRowHeight,
  Tx,
  type Block,
  type Entry,
  type PasteMode,
  type PasteOpts,
} from '../core/ops';
import { fillSeries } from '../core/fill';
import { autoRowPt, stepDecimals, type BorderKind } from '../core/format';
import { FormatBar, type MergeKind } from './FormatBar';
import { SheetTabs } from './SheetTabs';
import { FilterMenu } from './FilterMenu';
import { SortDialog } from './SortDialog';
import { CfDialog } from './CfDialog';
import { overTable, tableStyleAt } from '../core/tableStyle';
import { addCfRule, clearCfRules, removeCfRule } from '../core/cfEdit';
import { DuplicatesDialog, SplitDialog } from './DataDialog';
import { PasteSpecialDialog } from './PasteSpecialDialog';
import { FormatCellsDialog, type FormatCellsResult } from './FormatCellsDialog';
import { PivotDialog } from './PivotDialog';
import { NamesDialog } from './NamesDialog';
import { NoteDialog } from './NoteDialog';
import { ChartDialog } from './ChartDialog';
import { addChart, deleteDrawing, moveDrawing, updateChart } from '../core/charts';
import type { Anchor } from '../core/drawing';
import { deleteNote, isThreaded, setNote } from '../core/notes';
import { ProtectDialog, type ProtectKind } from './ProtectDialog';
import { protectSheetAttrs, setSheetProtection, setStructureLock, structureLocked, verifyPassword } from '../core/protect';
import { deleteName, nameTarget, saveName, type DefinedName } from '../core/names';
import { createPivot, defaultConfig, fieldNames, refreshPivot } from '../core/pivot';
import { groupRange, levelAt, maxLevel, showLevel, summaryAfter, toggleGroup, collapsedAt, ungroupRange, type Axis as OutlineAxis } from '../core/outline';
import { removeDuplicates, splitWidth, textToColumns } from '../core/data';
import { FindReplace } from './FindReplace';
import { PrintDialog } from './PrintDialog';
import { editText, replaceAll as replaceAllCells, replaceCell, type FindOptions, type Hit as FindHit } from '../core/find';
import { imagesToPdf, isMobile, loadZoom, saveZoom, scriptErrors, usePinchZoom, zoomIn } from '@wz/editor-shared';
import { SheetMobileBar } from './SheetMobile';
import { FloatViewButton } from '@wz/editor-shared/FloatViewButton';
import { autoSumRange, callAt, FUNCTION_INFO, suggest, typingName } from '../core/assist';
import { evalCf } from '../core/cf';
import { SheetChart } from './SheetChart';
import { checkValue, listItems, validationAt } from '../core/validate';
import { parseInput } from '../core/input';
import { addSheet, copySheet, deleteSheet, moveSheet, renameSheet, setSheetHidden, setTabColor, sheetNameError } from '../core/sheets';
import { columnExtras, filterEnv, columnValues, currentRegion, guessHeader, hiddenByOthers, refilter, setAutoFilter, setColumnFilter, sortRange, type SortKey } from '../core/sortFilter';
import type { ColumnFilter } from '../core/workbook';
import type { StyleChange } from '../core/styles';
import { markSaved, saveCsv, saveXlsx } from '../core/save';
import { blockFromHtml, decodeCsvBytes } from '../core/clip';
import { Engine } from '../core/formula/eval';
import { canInsertRef, cycleAbsolute, refSpans, refStr } from '../core/formula/refs';
import {
  buildAxis,
  borderEdges,
  buildMergeIndex,
  CELL_PAD,
  cellPadding,
  cellKey,
  cellBorder,
  cellColors,
  cellName,
  collect,
  contains,
  edgeAt,
  expandToMerges,
  findAll,
  fmtStat,
  gridExtent,
  indexAt,
  moveBy,
  rangeName,
  rangeOf,
  sizeOf,
  snap,
  statsOf,
  toHtml,
  toMarkdown,
  toTsv,
  type Axis,
  type Pos,
} from '../core/grid';
import { Icon } from '@wz/editor-shared/Icon';
import '@wz/editor-shared/controls.css';
import './sheet.css';

interface Sel {
  anchor: Pos;
  /** 活动单元格（合并单元格时为左上角） */
  active: Pos;
  /** Shift 扩选的另一端 */
  end: Pos;
  range: Range;
  /** Ctrl+点击加上的其它区域（不连续选区）；活动区域是 range */
  extra?: Range[];
}

/** 撤销记录里保存的界面状态：在哪张表、选区 */
interface UiState {
  si: number;
  sel: Sel;
}

/** 拖动分隔线时列宽 / 行高的下限（像素），避免拖到 0 后无法再拖回来 */
const MIN_SIZE = 4;
const ZOOMS = [0.7, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2];
const MAX_UNDO = 100;
/** 编辑公式时各个引用的框颜色（同 Excel 的顺序） */
const REF_COLORS = ['#4472c4', '#c0504d', '#9b59b6', '#2e9e5b', '#e67e22', '#16a3b8'];
const cellSel = (p: Pos, range?: Range): Sel => ({ anchor: p, active: p, end: p, range: range ?? { r0: p.r, c0: p.c, r1: p.r, c1: p.c } });

/** 用 copy 事件写剪贴板：可同时写纯文本（TSV）和 HTML 表格 */
function writeClipboard(plain: string, html?: string) {
  const onCopy = (e: ClipboardEvent) => {
    e.clipboardData?.setData('text/plain', plain);
    if (html) e.clipboardData?.setData('text/html', html);
    e.preventDefault();
  };
  document.addEventListener('copy', onCopy);
  const ok = document.execCommand('copy');
  document.removeEventListener('copy', onCopy);
  if (!ok) void navigator.clipboard.writeText(plain);
}

/**
 * 表格内部剪贴板（各个表格标签共用）：复制时记下带格式 / 公式的内容和写进系统剪贴板的文字，
 * 粘贴时系统剪贴板的文字还是它，就用带格式的内容；否则当作外部文字。
 */
/** 分组栏里一级分组占的宽度（像素） */
const OG = 13;

/** whole：复制的是整行 / 整列（“插入复制的单元格”插入整行 / 整列） */
let clip: { block: Block; tsv: string; book: Workbook; si: number; range: Range; cut: boolean; whole?: 'rows' | 'cols' } | null = null;

export function SheetViewer({ tab }: { tab: EditorTab }) {
  const path = tab.path;
  const book = useRef<Workbook | null>(null);
  const engine = useRef<Engine | null>(null);
  /** 这个文件所在的工作包（不在工作包里为 null）：列属性、表单只在工作包里 */
  const bundleRef = useRef<BundleContext | null>(null);
  const [loaded, setLoaded] = useState<Workbook | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [reload, setReload] = useState(0);
  const [si, setSi] = useState(0);
  /** AI 助手读的当前选区（SheetView 每次渲染写入，A1 引用） */
  const aiSel = useRef('');
  const liveSi = useRef(0);
  liveSi.current = si;
  const closeAssistant = useRef<(() => void) | null>(null);
  useEffect(() => () => closeAssistant.current?.(), []);
  const [loading, setLoading] = useState(true);
  /** 读取超过 SKELETON_DELAY 才显示骨架屏，小文件不闪 */
  const [slow, setSlow] = useState(false);
  /** 骨架屏出现过：表格淡入接替它 */
  const [fadeIn, setFadeIn] = useState(false);
  const slowRef = useRef(false);
  /** 模型改动计数：每次编辑 / 撤销加 1，界面据此重绘 */
  const [version, setVersion] = useState(0);
  const undoStack = useRef<(Entry & { id: number })[]>([]);
  const redoStack = useRef<(Entry & { id: number })[]>([]);
  const nextId = useRef(1);
  /** 保存时撤销栈顶的记录 id：撤销 / 重做回到这里就不算未保存 */
  const [savedId, setSavedId] = useState(0);
  /** 文件签名（大小 + 修改时间），用来判断是否被外部程序改过 */
  const sig = useRef('');
  const [external, setExternal] = useState(false);
  const skipLoad = useRef<string | null>(null);
  const loadPath = useRef(path);
  /** 查找替换面板（换表时保留）与换表后要选中的格子 */
  const [findMode, setFindMode] = useState<'find' | 'replace' | null>(null);
  const pendingGoto = useRef<FindHit | null>(null);
  /** 自己保存后一小段时间内的文件变动通知不算外部修改 */
  const ignoreFsUntil = useRef(0);
  const folders = host().useWorkFolders();
  const fsEvent = host().useFsEvent();

  const savedIdRef = useRef(0);
  savedIdRef.current = savedId;
  const topId = undoStack.current[undoStack.current.length - 1]?.id ?? 0;
  const dirty = !!loaded && topId !== savedId;
  useEffect(() => {
    host().updateTab(tab.id, { dirty });
  }, [dirty, tab.id]);

  useEffect(() => {
    if (!loading) return;
    const t = window.setTimeout(() => {
      slowRef.current = true;
      setSlow(true);
    }, SKELETON_DELAY);
    return () => window.clearTimeout(t);
  }, [loading]);

  const fileSig = useCallback(async () => {
    if (!path) return '';
    const list = await host().listDir(dirname(path)).catch(() => []);
    const e = list.find((x) => x.path.toLowerCase() === path.toLowerCase() || x.name === basename(path));
    return e ? `${e.size}|${e.modified}` : '';
  }, [path]);

  useEffect(() => {
    if (loadPath.current !== path) {
      loadPath.current = path;
      if (loaded) { skipLoad.current = null; return; }
    }
    if (path && skipLoad.current === path) {
      skipLoad.current = null;
      return;
    }
    let cancelled = false;
    setLoading(true);
    void (async () => {
      try {
        let next: Workbook;
        if (!path) next = blankWorkbook();
        else {
          const name = basename(path).replace(/\.[^.]+$/, '');
          const ext = extOf(path);
          if (ext === 'csv') {
            const dec = decodeCsvBytes(await host().readBytes(path));
            next = await parseInWorker({ kind: 'csv', text: dec.text, name });
            next.csv = { delim: next.csv?.delim ?? ',', encoding: dec.encoding, bom: dec.bom };
          } else if (ext === 'xls') next = fromXls(await host().readXls(path));
          else {
            const bytes = await host().readBytes(path);
            next = await parseInWorker({ kind: 'xlsx', data: bytes.slice() });
            next.source = bytes;
          }
        }
        // 上次没保存就退出了：有自动保存的草稿时提示恢复
        let recovered = false;
        const draft = path ? await loadDraft(path) : null;
        if (draft && !cancelled) {
          const ok = (await host().confirm(`「${basename(path!)}」有 ${draftTime(draft.time)} 自动保存的修改，上次没有保存就关闭了。\n\n恢复这些修改吗？选择“取消”将打开文件里保存的内容。`));
          if (ok) {
            try {
              const kind = next.kind;
              const csv = next.csv;
              next =
                typeof draft.data === 'string'
                  ? await parseInWorker({ kind: 'csv', text: draft.data, name: basename(path!).replace(/\.[^.]+$/, '') })
                  : await parseInWorker({ kind: 'xlsx', data: draft.data.slice() });
              if (kind === 'xlsx' && typeof draft.data !== 'string') next.source = draft.data;
              else next.kind = kind;
              if (csv) next.csv = csv;
              recovered = true;
            } catch {
              host().toast('草稿已损坏，打开文件里保存的内容');
            }
          }
          if (!recovered) await clearDraft(path!);
        }
        const s = await fileSig();
        if (cancelled) return;
        // 工作包里的表格：把包里存的列属性套到各工作表（表头文字对不上的列按表头重新找，找不到的标为未绑定）
        const ctx = path ? (host().bundleOf?.(path) ?? null) : null;
        bundleRef.current = ctx;
        if (ctx) {
          next.externalValues = (ref) => ctx.readRange(ref);
          const file = ctx.columns();
          const moved: string[] = [];
          const unbound: string[] = [];
          for (const sh of next.sheets) {
            const sc = file.sheets[sh.name];
            if (!sc) continue;
            const r = bindColumns(sh, sc);
            sh.columns = r.sc;
            moved.push(...r.moved);
            unbound.push(...r.unbound);
          }
          if (unbound.length) host().toast(`列「${unbound.join('、')}」的表头在表格里找不到了，暂不套用列属性（可在列属性里重新设置）`, 6000);
          else if (moved.length) host().toast(`列「${moved.join('、')}」的位置变了，已按表头文字重新对应`, 5000);
        }
        book.current = next;
        engine.current = new Engine(next);
        // 空闲时分析一遍公式，标出不能计算的
        const eng = engine.current;
        window.setTimeout(() => {
          if (eng.analyze() || eng.unsupported.size) setVersion((v) => v + 1);
        }, 0);
        sig.current = s;
        undoStack.current = [];
        redoStack.current = [];
        // 恢复的草稿算未保存的修改
        setSavedId(recovered ? -1 : 0);
        setExternal(false);
        setLoaded(next);
        setVersion((v) => v + 1);
        setError(null);
        setSi((i) => (i < next.sheets.length && !next.sheets[i].hidden ? i : next.sheets.findIndex((x) => !x.hidden)));
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      }
      if (cancelled) return;
      // 同一次渲染里切换，不会先出现空白帧
      setFadeIn(slowRef.current);
      slowRef.current = false;
      setSlow(false);
      setLoading(false);
    })();
    return () => {
      cancelled = true;
    };
  }, [path, reload, fileSig]);

  // 外部程序改了文件：没有未保存的修改就直接重新读取，有修改时提示
  const dirtyRef = useRef(dirty);
  dirtyRef.current = dirty;
  useEffect(() => {
    if (!path || !fsEvent.tick || Date.now() < ignoreFsUntil.current || !fsEvent.dirs.some((d) => d.toLowerCase() === dirname(path).toLowerCase())) return;
    void fileSig().then((s) => {
      if (!s || !sig.current || s === sig.current) return;
      if (dirtyRef.current) setExternal(true);
      else setReload((n) => n + 1);
    });
  }, [fsEvent, path, fileSig]);

  // 有未保存的修改时定时存草稿；关闭标签时删除草稿（正常关闭不需要恢复）
  useEffect(() => {
    if (!path || !dirty) return;
    const t = window.setInterval(() => {
      const b = book.current;
      if (!b) return;
      try {
        const data = b.kind === 'csv' ? saveCsv(b.sheets[0], b.csv?.delim ?? ',') : saveXlsx(b);
        void saveDraft(path, data);
      } catch {
        /* 草稿只是保险，失败不影响编辑 */
      }
    }, AUTOSAVE_MS);
    return () => window.clearInterval(t);
  }, [path, dirty]);
  useEffect(() => () => void (path && clearDraft(path)), [path]);

  const openExternal = () => path && void host().guard(host().openWithSystem(path));

  // ---------- 编辑 / 撤销 ----------
  const bump = () => setVersion((v) => v + 1);
  /** AI 助手（只有 AI 版的工作台提供）：AI 用命令行同一套 op 改表格，每个 op 的撤销记录进本编辑器的撤销栈 */
  const ai = host().ai;
  const openAssistant = ai?.openAssistant
    ? () => {
        closeAssistant.current?.();
        closeAssistant.current = ai.openAssistant!({
          kind: 'sheet',
          title: tab.title,
          api: {
            book: () => book.current,
            engine: () => engine.current,
            activeSheet: () => book.current?.sheets[liveSi.current]?.name ?? '',
            selectionRef: () => aiSel.current,
            record: (entry) => {
              undoStack.current.push({ ...(entry as Entry), id: nextId.current++ });
              if (undoStack.current.length > MAX_UNDO) undoStack.current.shift();
              redoStack.current = [];
              bump();
            },
          },
        });
      }
    : undefined;
  const recalc = (changed: { si: number; r: number; c: number }[], opts: { rebuild?: boolean; all?: boolean }) => {
    const eng = engine.current;
    if (!eng) return;
    eng.recalc(changed, opts);
    if (eng.circular) host().toast(`有 ${eng.circular} 个公式存在循环引用（显示为 #CIRC!）`);
  };
  /** 撤销 / 重做后：恢复的格子及引用它们的公式重算 */
  const recalcEntry = (e: Entry) => recalc(e.cells, { rebuild: true, all: !!(e.sheets.length || e.book) });
  const run = useCallback(<T,>(label: string, fn: (tx: Tx) => T, ui?: UiState): T | undefined => {
    const b = book.current;
    if (!b) return undefined;
    const wasDirty = b.sheets.map((s) => s.dirty);
    const tx = new Tx(b, label, ui);
    const out = fn(tx);
    const e = tx.commit();
    if (!e) return out;
    if (violatesProtection(b, e)) {
      // 受保护的工作表：撤回这次改动
      applyEntry(b, e, -1);
      b.sheets.forEach((s, i) => (s.dirty = wasDirty[i] ?? s.dirty));
      bump();
      host().toast('工作表受保护：只能在未锁定的单元格里输入。要修改其它内容，请先在 Excel 里撤销工作表保护');
      return undefined;
    }
    const formulas = e.cells.some((x) => x.before?.f || x.after?.f || x.before?.fraw);
    recalc(tx.changed, { rebuild: formulas, all: tx.structural });
    undoStack.current.push({ ...e, id: nextId.current++ });
    if (undoStack.current.length > MAX_UNDO) undoStack.current.shift();
    redoStack.current = [];
    bump();
    return out;
  }, []);
  const undo = useCallback((): UiState | undefined => {
    const b = book.current;
    const e = undoStack.current.pop();
    if (!b || !e) return undefined;
    applyEntry(b, e, -1);
    recalcEntry(e);
    redoStack.current.push(e);
    bump();
    return e.ui as UiState | undefined;
  }, []);
  const redo = useCallback((): UiState | undefined => {
    const b = book.current;
    const e = redoStack.current.pop();
    if (!b || !e) return undefined;
    applyEntry(b, e, 1);
    recalcEntry(e);
    undoStack.current.push(e);
    bump();
    return e.ui as UiState | undefined;
  }, []);

  const inWorkFolder = !!path && folders.some((f) => {
    const root = f.replace(/[\\/]+$/, '').toLowerCase();
    const p = path.toLowerCase();
    return p.startsWith(`${root}\\`) || p.startsWith(`${root}/`);
  });

  // 不在工作文件夹里但是用户自己打开的文件（资源管理器右键、常用文件）也能原地保存：问宿主
  const [openedWritable, setOpenedWritable] = useState(false);
  useEffect(() => {
    setOpenedWritable(false);
    if (!path || inWorkFolder) return;
    let cancelled = false;
    void host().canWrite?.(path).then((ok) => !cancelled && setOpenedWritable(ok), () => undefined);
    return () => {
      cancelled = true;
    };
  }, [path, inWorkFolder]);
  const inPlace = inWorkFolder || openedWritable;

  // ---------- 保存 ----------
  const save = useCallback(
    async (asNew: boolean): Promise<boolean> => {
      const b = book.current;
      const s = host();
      if (!b) return false;
      let target = path;
      let kind = b.kind;
      if (kind === 'xls' || kind === 'new' || !inPlace) asNew = true;
      if (asNew) {
        const dir = path && inWorkFolder ? dirname(path) : undefined;
        const stem = path ? basename(path).replace(/\.[^.]+$/, '') + (kind === 'xls' || !inWorkFolder ? '' : '_副本') : '新建工作簿';
        const why = kind === 'xls' ? '.xls 不能直接保存，另存为 .xlsx' : !path ? '保存新建的工作簿' : !inPlace ? '此文件不能原地保存，另存到' : '另存为';
        const exts = kind === 'csv' ? ['csv', 'xlsx'] : ['xlsx', 'csv'];
        const picked = await pickSave({ title: why, name: `${stem}.${exts[0]}`, dir, label: '表格', exts });
        if (!picked) return false;
        target = picked;
        kind = /\.csv$/i.test(picked) ? 'csv' : 'xlsx';
      }
      if (!target) return false;
      const savedRevision = undoStack.current[undoStack.current.length - 1]?.id ?? 0;
      if (kind === 'csv') {
        const sheet = b.sheets[si] ?? b.sheets[0];
        const lossy = b.sheets.length > 1 || sheet.cells.some((row) => row?.some((c) => c?.f || c?.s)) || sheet.merges.length > 0;
        if (lossy && !(await host().confirm('CSV 只能保存当前工作表的值（公式、格式、合并单元格和其它工作表不会保存）。\n\n继续保存为 CSV 吗？'))) return false;
        const csv = b.csv ?? { delim: ',', encoding: 'UTF-8', bom: true };
        const text = (csv.bom && /utf-?8/i.test(csv.encoding) ? '﻿' : '') + saveCsv(sheet, csv.delim);
        if ((await s.guard(host().writeText(target, text, csv.encoding).then(() => true))) === undefined) return false;
        b.csv = csv;
        b.kind = 'csv';
        if ((undoStack.current[undoStack.current.length - 1]?.id ?? 0) === savedRevision) {
          for (const x of b.sheets) x.dirty = false;
        }
      } else {
        // 工作包里的表格：列属性随这次保存写进包（不写进 xlsx）。只改了列属性时 xlsx 不动，只写包
        const ctx = target === path ? bundleRef.current : null;
        let columnsOnly = false;
        if (ctx) {
          for (const sh of b.sheets) ctx.setColumns(sh.name, sh.columns ?? null);
          for (const name of Object.keys(ctx.columns().sheets)) if (!b.sheets.some((sh) => sh.name === name)) ctx.setColumns(name, null);
          const stack = undoStack.current;
          const at = savedIdRef.current === 0 ? -1 : stack.findIndex((x) => x.id === savedIdRef.current);
          columnsOnly =
            b.kind === 'xlsx' &&
            !!b.source &&
            !b.namesDirty &&
            !b.sheets.some((sh) => sh.dirty) &&
            (savedIdRef.current === 0 || at >= 0) &&
            stack.slice(at + 1).length > 0 &&
            stack.slice(at + 1).every((x) => !!x.cols?.length && !x.cells.length && !x.sheets.length && !x.metas.length && !x.book);
        }
        if (columnsOnly && ctx) {
          if ((await s.guard(ctx.flush().then(() => true))) === undefined) return false;
        } else {
          let bytes: Uint8Array;
          try {
            bytes = saveXlsx(b);
          } catch (e) {
            s.toast(`保存失败：${e instanceof Error ? e.message : String(e)}`);
            return false;
          }
          if ((await s.guard(host().writeBytes(target, bytes).then(() => true))) === undefined) return false;
          if ((undoStack.current[undoStack.current.length - 1]?.id ?? 0) === savedRevision) markSaved(b, bytes);
        }
      }
      setSavedId(savedRevision);
      setExternal(false);
      void clearDraft(target);
      if (path && path !== target) void clearDraft(path);
      if (target !== path) {
        // 路径变了不用重新读取：模型就是刚保存的内容
        skipLoad.current = target;
        bundleRef.current = s.bundleOf?.(target) ?? null;
        s.updateTab(tab.id, { path: target, title: basename(target) });
      }
      ignoreFsUntil.current = Date.now() + 2500;
      sig.current = await fileSig();
      const unsaved = (undoStack.current[undoStack.current.length - 1]?.id ?? 0) !== savedRevision;
      s.updateTab(tab.id, { dirty: unsaved });
      s.toast(`已保存 ${basename(target)}`);
      bump();
      return !unsaved;
    },
    [path, inWorkFolder, inPlace, si, tab.id, fileSig],
  );

  const saveRef = useRef(save);
  saveRef.current = save;
  // 关闭有未保存修改的标签时，工作台通过它保存（新工作簿弹“另存为”）
  useEffect(() => host().registerSave?.(tab.id, () => saveRef.current(false)), [tab.id]);

  /** 导出的 PDF 存到表格所在的工作文件夹（不在工作文件夹里时存到第一个工作文件夹） */
  const exportPdf = async (bytes: Uint8Array) => {
    const s = host();
    const stem = path ? basename(path).replace(/\.[^.]+$/, '') : '新建工作簿';
    const target = await pickSave({ title: '导出 PDF', name: `${stem}.pdf`, dir: path && inWorkFolder ? dirname(path) : undefined, label: 'PDF 文档', exts: ['pdf'] });
    if (!target) return;
    if ((await s.guard(host().writeBytes(target, bytes).then(() => true))) === undefined) return;
    s.toast(`已导出 ${basename(target)}`);
  };

  const notes = useMemo(() => {
    const out: string[] = [];
    if (!loaded) return out;
    if (path && !inPlace) out.push('此文件不在工作文件夹内，保存时会另存到工作文件夹');
    if (loaded.kind === 'xls') out.push('旧版 .xls 文件保存时会另存为 .xlsx');
    return out;
  }, [loaded, path, inPlace]);

  if (error) {
    return (
      <div className="sheet-state">
        <div className="sheet-state-card">
          <strong>无法在工作台内显示这个表格</strong>
          <span className="sub">{error}</span>
          <div style={{ display: 'flex', gap: 8, justifyContent: 'center' }}>
            <button className="btn" onClick={() => setReload((n) => n + 1)}>
              重试
            </button>
            {path && (
              <button className="btn sheet-excel-btn" onClick={openExternal}>
                <ExcelIcon />
                用 Excel 打开
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }
  const b = book.current;
  // 首次读取：先是与表格同色的空白底，超过 SKELETON_DELAY 才出骨架屏
  if (!loaded || !b) return slow ? <SheetSkeleton name={path ? basename(path) : '新建工作簿'} /> : <div className="sheet-blank" />;
  const sheet = b.sheets[si] ?? b.sheets[0];
  return (
    <SheetView
      key={si}
      bundle={bundleRef.current}
      fileName={tab.path ? basename(tab.path) : undefined}
      book={b}
      engine={engine.current}
      sheet={sheet}
      sheetIndex={si}
      version={version}
      fadeIn={fadeIn}
      dirty={dirty}
      canUndo={undoStack.current.length > 0}
      canRedo={redoStack.current.length > 0}
      notes={notes}
      external={external}
      hasPath={!!path}
      tabId={tab.id}
      path={path}
      onAssistant={openAssistant}
      aiSel={aiSel}
      onSheet={(i) => {
        setFadeIn(false);
        setSi(i);
      }}
      onReload={async () => {
        if (dirty && !(await host().confirm('重新读取会放弃未保存的修改，确定吗？'))) return;
        setReload((n) => n + 1);
      }}
      onKeepMine={() => setExternal(false)}
      onExternal={openExternal}
      run={run}
      findMode={findMode}
      setFindMode={setFindMode}
      pendingGoto={pendingGoto}
      exportPdf={exportPdf}
      undo={undo}
      redo={redo}
      save={save}
    />
  );
}

const SKELETON_DELAY = 220;

/** Excel 图标：带网格的页面 + 左侧深绿方块里的白色 X（颜色固定为 Excel 绿，不随主题变化） */
function ExcelIcon({ size = 16 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" aria-hidden="true">
      <path d="M5 1.5h6.5L14.5 4.5v9a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1v-11a1 1 0 0 1 1-1z" fill="#fff" stroke="#217346" />
      <path d="M9.5 6h3.5M9.5 8h3.5M9.5 10h3.5M9.5 12h3.5" stroke="#33a06f" strokeWidth="1.1" />
      <rect x="0.5" y="4" width="8" height="8" rx="1" fill="#107c41" />
      <path d="M2.8 6l3.4 4M6.2 6l-3.4 4" stroke="#fff" strokeWidth="1.3" strokeLinecap="round" />
    </svg>
  );
}

/** 读取中：与表格同样的空白网格（行号列标 + 网格线）+ 顶部进度条 + 文件名；超过 3 秒显示已用时间 */
function SheetSkeleton({ name }: { name: string }) {
  const [sec, setSec] = useState(0);
  useEffect(() => {
    const t = window.setInterval(() => setSec((n) => n + 1), 1000);
    return () => window.clearInterval(t);
  }, []);
  return (
    <div className="sheet sheet-loading" aria-label="正在读取表格" aria-busy="true">
      <div className="sheet-loading-bar" />
      <div className="sheet-loading-grid" />
      <div className="sheet-loading-card">
        <span className="sheet-loading-spin" />
        <div className="sheet-loading-text">
          <span className="ellipsis">正在读取「{name}」</span>
          <span className="sub">{sec >= 3 ? `已用 ${sec} 秒，文件较大或在网络文件夹上时会慢一些` : '请稍候…'}</span>
        </div>
      </div>
    </div>
  );
}

interface ViewProps {
  /** 所在的工作包；null = 不在工作包里（不显示列属性、表单界面） */
  bundle: BundleContext | null;
  book: Workbook;
  engine: Engine | null;
  sheet: SheetData;
  sheetIndex: number;
  version: number;
  /** 骨架屏之后首次出现：淡入 */
  fadeIn: boolean;
  /** 正在重新读取（刷新按钮转动） */
  dirty: boolean;
  canUndo: boolean;
  canRedo: boolean;
  notes: string[];
  /** 文件被外部修改（且有未保存的修改） */
  external: boolean;
  hasPath: boolean;
  /** 标签页和文件路径（工具条的“悬浮显示”按钮用） */
  tabId: string;
  path?: string | null;
  /** 打开 AI 助手（只有 AI 版有） */
  onAssistant?: () => void;
  /** 当前选区写到这里，给 AI 助手读 */
  aiSel: { current: string };
  /** 文件名（页眉页脚的 &F） */
  fileName?: string;
  onSheet: (i: number) => void;
  onReload: () => void;
  onKeepMine: () => void;
  onExternal: () => void;
  run: <T>(label: string, fn: (tx: Tx) => T, ui?: UiState) => T | undefined;
  findMode: 'find' | 'replace' | null;
  setFindMode: (m: 'find' | 'replace' | null) => void;
  pendingGoto: React.MutableRefObject<FindHit | null>;
  exportPdf: (bytes: Uint8Array) => Promise<void>;
  undo: () => UiState | undefined;
  redo: () => UiState | undefined;
  save: (asNew: boolean) => Promise<unknown>;
}

interface Editing {
  r: number;
  c: number;
  text: string;
  /** 进入编辑时的原文（没改就不写入） */
  orig: string;
  /** enter = 直接打字进入（方向键提交并移动）；edit = F2 / 双击进入（方向键移动光标） */
  mode: 'enter' | 'edit';
  /** 在编辑栏里编辑 */
  bar?: boolean;
}

function SheetView(props: ViewProps) {
  const { book, sheet, sheetIndex, version, fadeIn, onSheet, onReload, onExternal, run } = props;
  const si = sheetIndex;
  // 只看挂载时的值：刷新不会重放淡入
  const [enter] = useState(fadeIn);
  const scroller = useRef<HTMLDivElement>(null);
  const view = useRef<HTMLDivElement>(null);
  const kb = useRef<HTMLTextAreaElement>(null);
  const barInput = useRef<HTMLTextAreaElement>(null);
  // 安卓手机版（docs/ANDROID.md §6.3）：不显示功能区 / 编辑栏，底部工具条；隐藏输入框不拿焦点（拿到就弹键盘），双指缩放
  const mobile = isMobile();
  const gridContent = useRef<HTMLDivElement>(null);
  /** 手机：编辑模式（底部编辑栏、格式等）；选区模式 = 点格子扩展选区（手指拖动用来滚动） */
  const [mEdit, setMEdit] = useState(false);
  const [mRange, setMRange] = useState(false);
  const searchInput = useRef<HTMLInputElement>(null);
  const [vp, setVp] = useState({ w: 800, h: 500 });
  const [scroll, setScroll] = useState({ x: 0, y: 0 });
  // 按上次用的缩放打开（悬浮文档窗口里再开同一个文件时缩放才和主窗口一致）
  const [zoom, setZoom] = useState(() => loadZoom<number>('sheet.zoom', 1, zoomIn(ZOOMS[0], ZOOMS[ZOOMS.length - 1])));
  useEffect(() => saveZoom('sheet.zoom', zoom), [zoom]);
  const [sel, setSel] = useState<Sel>(() => cellSel({ r: 0, c: 0 }));
  props.aiSel.current = `${colName(sel.range.c0)}${sel.range.r0 + 1}:${colName(sel.range.c1)}${sel.range.r1 + 1}`;
  const [addr, setAddr] = useState<string | null>(null);
  const [query, setQuery] = useState('');
  const [hitIdx, setHitIdx] = useState(-1);
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  /** 冻结下拉菜单的位置（按钮下方） */
  const [freezeMenu, setFreezeMenu] = useState<{ x: number; y: number } | null>(null);
  const [detail, setDetail] = useState<Pos | null>(null);
  const [copied, setCopied] = useState<{ n: number; text: string } | null>(null);
  const [editing, setEditing] = useState<Editing | null>(null);
  const [noteHidden, setNoteHidden] = useState(false);
  /** 复制 / 剪切后的虚线框（Esc 或编辑后消失） */
  const [marquee, setMarquee] = useState<{ range: Range; cut: boolean } | null>(null);
  /** 拖动填充柄时的目标区域 */
  const [fillTo, setFillTo] = useState<Range | null>(null);
  /** 打开的筛选下拉（列号与位置）、排序对话框、排序和筛选菜单 */
  const [filterMenu, setFilterMenu] = useState<{ c: number; x: number; y: number } | null>(null);
  const [sortDialog, setSortDialog] = useState<{ range: Range; header: boolean } | null>(null);
  const [pasteDlg, setPasteDlg] = useState(false);
  /** 设置单元格格式对话框（Ctrl+1） */
  const [fmtDlg, setFmtDlg] = useState(false);
  const [pivotDlg, setPivotDlg] = useState<Range | null>(null);
  const [namesDlg, setNamesDlg] = useState<{ draft?: { name?: string; f?: string } } | null>(null);
  const userName = host().useUserName() ?? '作者';
  const [chartDlg, setChartDlg] = useState<{ range?: Range; index?: number } | null>(null);
  /** 选中的图片 / 图表（drawings 序号）和拖动中的预览位置（未缩放的像素，相对整张表） */
  const [selDrawing, setSelDrawing] = useState<number | null>(null);
  const [drawGhost, setDrawGhost] = useState<{ x: number; y: number; w: number; h: number } | null>(null);
  const ghostRef = useRef<{ x: number; y: number; w: number; h: number } | null>(null);
  const dragDrawing = useRef<{ i: number; dir: string; sx: number; sy: number; rect: { x: number; y: number; w: number; h: number } } | null>(null);
  const [noteDlg, setNoteDlg] = useState<{ r: number; c: number } | null>(null);
  const [protectDlg, setProtectDlg] = useState<ProtectKind | null>(null);
  const [nameList, setNameList] = useState<{ x: number; y: number } | null>(null);
  const [dataDialog, setDataDialog] = useState<{ kind: 'dup' | 'split'; range: Range; header: boolean } | null>(null);
  const [sortMenu, setSortMenu] = useState<{ x: number; y: number } | null>(null);
  const [cfOpen, setCfOpen] = useState(false);
  const [printOpen, setPrintOpen] = useState(false);
  const [saveMenu, setSaveMenu] = useState(false);
  const [barTab, setBarTab] = useState<'home' | 'page'>('home');
  const saveMenuRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!saveMenu) return;
    // 捕获阶段监听：表格区域会拦截鼠标事件，冒泡阶段收不到
    const close = (e: MouseEvent) => {
      if (!saveMenuRef.current?.contains(e.target as Node)) setSaveMenu(false);
    };
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && setSaveMenu(false);
    document.addEventListener('mousedown', close, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', close, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [saveMenu]);
  /** 数据验证的下拉列表是否打开 */
  const [listOpen, setListOpen] = useState(false);
  // ---- 工作包里的列属性 / 表单（bundle 为 null 时都不出现）----
  const bundle = props.bundle;
  const [colPanel, setColPanel] = useState<{ c: number; x: number; y: number } | null>(null);
  const [smartDlg, setSmartDlg] = useState(false);
  const [picker, setPicker] = useState<{ r: number; c: number } | null>(null);
  const [formPanel, setFormPanel] = useState<{ formId: string } | null>(null);
  const [designer, setDesigner] = useState<{ form: FormDef | null } | null>(null);
  const [menuCol, setMenuCol] = useState(0);
  const [, setFormsTick] = useState(0);
  useEffect(() => bundle?.subscribe(() => setFormsTick((t) => t + 1)), [bundle]);
  /** 编辑框里的光标位置（函数提示用）与提示列表里选中的项；hintOff = 按 Esc 关掉提示时的文字 */
  const [caretAt, setCaretAt] = useState(0);
  const [hintIdx, setHintIdx] = useState(0);
  const [hintOff, setHintOff] = useState<string | null>(null);
  /** 格式刷：记下的格式块；once 刷一次后结束，sticky 连续刷到 Esc */
  const [painter, setPainter] = useState<{ block: Block; mode: 'once' | 'sticky' } | null>(null);

  const extBase = useMemo(() => gridExtent(sheet), [sheet, version]); // eslint-disable-line react-hooks/exhaustive-deps
  // 选区附近总留出空白行列，可以一直往下 / 往右输入
  const ext = useMemo(
    () => ({
      rows: Math.min(1048576, Math.max(extBase.rows, Math.ceil((Math.max(sel.end.r, sel.active.r) + 30) / 50) * 50)),
      cols: Math.min(16384, Math.max(extBase.cols, Math.ceil((Math.max(sel.end.c, sel.active.c) + 8) / 8) * 8)),
    }),
    [extBase, sel.end.r, sel.end.c, sel.active.r, sel.active.c],
  );
  /** 拖动分隔线时的实时尺寸（未缩放的像素），松开后写入模型 */
  const [fit, setFit] = useState<Record<number, number>>({});
  const [rowFit, setRowFit] = useState<Record<number, number>>({});
  const cols = useMemo(() => buildAxis(ext.cols, { ...sheet.colWidths, ...fit }, sheet.defaultColWidth, zoom), [ext.cols, sheet, version, zoom, fit]); // eslint-disable-line react-hooks/exhaustive-deps
  const rows = useMemo(() => buildAxis(ext.rows, { ...sheet.rowHeights, ...rowFit }, sheet.defaultRowHeight, zoom), [ext.rows, sheet, version, zoom, rowFit]); // eslint-disable-line react-hooks/exhaustive-deps
  const merges = useMemo(() => buildMergeIndex(sheet.merges), [sheet, version]); // eslint-disable-line react-hooks/exhaustive-deps
  const freeze = sheet.freeze;
  const fr = Math.min(freeze.rows, ext.rows);
  const fc = Math.min(freeze.cols, ext.cols);
  const hw0 = Math.round((String(ext.rows).length * 7.5 + 20) * zoom);
  const hh0 = Math.round(24 * zoom);
  // 分组栏（行号左边 / 列标上边）：有分组时才占位，不随缩放变化
  const olRow = maxLevel(sheet, 'row');
  const olCol = maxLevel(sheet, 'col');
  const gw = (olRow ? (olRow + 1) * OG : 0) + (olCol ? OG : 0);
  const gh = (olCol ? (olCol + 1) * OG : 0) + (olRow ? OG : 0);
  const hw = hw0 + gw;
  const hh = hh0 + gh;
  const fw = cols.pos[fc];
  const fh = rows.pos[fr];
  const zoneW = Math.max(0, vp.w - hw - fw);
  const zoneH = Math.max(0, vp.h - hh - fh);
  const hiddenRow = useCallback((r: number) => sizeOf(rows, r) === 0, [rows]);
  const hiddenCol = useCallback((c: number) => sizeOf(cols, c) === 0, [cols]);
  const moveCtx = { sheet, rows: ext.rows, cols: ext.cols, merges, hiddenRow, hiddenCol };
  const defFont = useMemo(() => (book.styleBook ? defaultFontSize(book.styleBook) : 11), [book.styleBook]);
  /** 默认字体用界面字体显示，只有设置了别的字体的格子才换字体 */
  const defFontName = useMemo(() => (book.styleBook ? defaultFontName(book.styleBook) : undefined), [book.styleBook]);
  const ui = (): UiState => ({ si, sel });
  /** 活动单元格的数据验证 */
  const activeDv = validationAt(sheet, sel.active.r, sel.active.c);
  useEffect(() => setListOpen(false), [sel.active.r, sel.active.c, si]);
  useEffect(() => setPicker(null), [sel.active.r, sel.active.c, si]);
  /** 活动单元格所在列的列属性，和能弹出选择器的类型 */
  const activeMeta = bundle ? metaAt(sheet, sel.active.r, sel.active.c) : undefined;
  const pickable = !!activeMeta && (activeMeta.type === 'select' || activeMeta.type === 'multi' || activeMeta.type === 'date');
  const pickerMeta = picker && bundle ? metaAt(sheet, picker.r, picker.c) : undefined;
  /** 本次渲染里各列的选项（引用区域的选项只读一遍） */
  const optCache = new Map<ColumnMeta, SelectOption[]>();
  const optsOf = (cm: ColumnMeta): SelectOption[] => {
    let o = optCache.get(cm);
    if (!o) optCache.set(cm, (o = optionList(book, sheet, cm)));
    return o;
  };
  const missing = useMemo(() => (bundle && sheet.columns ? missingRequired(sheet) : []), [bundle, sheet, version]); // eslint-disable-line react-hooks/exhaustive-deps
  /** 选中的全部区域（不连续选区时不止一块） */
  const allRanges: Range[] = [...(sel.extra ?? []), sel.range];

  useLayoutEffect(() => {
    const el = scroller.current;
    if (!el) return;
    // 隐藏的标签页量到 0：保留上次尺寸，切回来时不会先画一帧空表
    const measure = () => {
      if (el.clientWidth && el.clientHeight) setVp((v) => (v.w === el.clientWidth && v.h === el.clientHeight ? v : { w: el.clientWidth, h: el.clientHeight }));
    };
    // 首次在绘制前同步量好，避免先按默认尺寸画一帧再跳变
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  useEffect(() => {
    if (!mobile) kb.current?.focus({ preventScroll: true });
  }, []);

  // 滚动事件本身每帧最多一次：同步渲染，单元格与滚动条同一帧到位，不抖动
  const onScroll = () => {
    const el = scroller.current;
    if (el) flushSync(() => setScroll({ x: el.scrollLeft, y: el.scrollTop }));
  };

  usePinchZoom(scroller, gridContent, zoom, mobile ? { min: 0.4, max: 3, onZoom: setZoom } : null);

  // Ctrl+滚轮缩放（需要非 passive 监听才能阻止页面缩放）
  useEffect(() => {
    const el = scroller.current;
    if (!el) return;
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      setZoom((z) => stepZoom(z, e.deltaY < 0 ? 1 : -1));
    };
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, []);

  /** 让单元格出现在可视区（冻结区里的不用滚） */
  const reveal = (p: Pos) => {
    const el = scroller.current;
    if (!el) return;
    if (p.c >= fc) {
      const x0 = cols.pos[p.c] - fw;
      const x1 = cols.pos[p.c + 1] - fw;
      if (x0 < el.scrollLeft) el.scrollLeft = x0;
      else if (x1 > el.scrollLeft + zoneW) el.scrollLeft = Math.min(x0, x1 - zoneW);
    }
    if (p.r >= fr) {
      const y0 = rows.pos[p.r] - fh;
      const y1 = rows.pos[p.r + 1] - fh;
      if (y0 < el.scrollTop) el.scrollTop = y0;
      else if (y1 > el.scrollTop + zoneH) el.scrollTop = Math.min(y0, y1 - zoneH);
    }
  };

  const selectCell = (p: Pos, scrollTo = true) => {
    const q = snap(p, merges);
    setSel(cellSel(q, expandToMerges({ r0: q.r, c0: q.c, r1: q.r, c1: q.c }, sheet.merges)));
    if (scrollTo) reveal(q);
  };
  const extendTo = (anchor: Pos, active: Pos, end: Pos, scrollTo = true, extra?: Range[]) => {
    setSel({ anchor, active, end, range: expandToMerges(rangeOf(anchor, end), sheet.merges), ...(extra?.length && { extra }) });
    if (scrollTo) reveal(end);
  };
  const selectRange = (g: Range) => {
    const a = { r: g.r0, c: g.c0 };
    setSel({ anchor: a, active: a, end: { r: g.r1, c: g.c1 }, range: expandToMerges(g, sheet.merges) });
  };
  const dataRange = (): Range => ({ r0: 0, c0: 0, r1: Math.max(0, sheet.rows - 1), c1: Math.max(0, sheet.cols - 1) });
  const selectAll = () => setSel({ anchor: { r: 0, c: 0 }, active: sel.active, end: { r: ext.rows - 1, c: ext.cols - 1 }, range: dataRange() });
  const focusGrid = () => !mobile && kb.current?.focus({ preventScroll: true });
  const restore = (u: UiState | undefined) => {
    if (!u) return;
    if (u.si !== si) onSheet(u.si);
    else setSel(u.sel);
  };

  // ---------- 复制 ----------
  const clipRange = (g: Range): Range => ({ r0: g.r0, c0: g.c0, r1: Math.min(g.r1, Math.max(sheet.rows - 1, g.r0)), c1: Math.min(g.c1, Math.max(sheet.cols - 1, g.c0)) });
  const flash = (n: number, text: string) => {
    setCopied({ n, text });
    window.setTimeout(() => setCopied((c) => (c?.n === n ? null : c)), 1600);
  };
  /** 选区 → 剪贴板内容（纯文本 TSV、HTML 表格） */
  const clipboardOf = (mode: 'table' | 'raw' | 'markdown', g: Range) => {
    const range = clipRange(g);
    const m = collect(sheet, book.styles, range, { raw: mode === 'raw', hiddenRow, hiddenCol });
    const w = m[0]?.length ?? 0;
    const plain = mode === 'markdown' ? toMarkdown(m) : toTsv(m);
    const html = mode === 'table' && w * m.length > 1 ? toHtml(m) : undefined;
    const label = w * m.length === 1 ? `已复制 ${rangeName(range)}` : `已复制 ${mode === 'markdown' ? 'Markdown 表格' : `${m.length} 行 × ${w} 列`}`;
    return { range, plain, html, label };
  };
  const remember = (g: Range, tsv: string, cut: boolean) => {
    const rows = g.c0 === 0 && g.c1 >= ext.cols - 1;
    const cols = g.r0 === 0 && g.r1 >= ext.rows - 1;
    const whole = rows !== cols ? (rows ? 'rows' : 'cols') : undefined;
    clip = { block: copyBlock(book, si, g), tsv, book, si, range: g, cut, ...(whole && { whole }) };
    setMarquee({ range: g, cut });
  };
  const copySelection = (mode: 'table' | 'raw' | 'markdown' = 'table', g = sel.range, cut = false) => {
    const out = clipboardOf(mode, g);
    writeClipboard(out.plain, out.html);
    if (mode === 'table') remember(g, out.plain, cut);
    flash(Date.now(), cut ? out.label.replace('已复制', '已剪切') : out.label);
  };
  const copyText = (text: string, label: string) => {
    writeClipboard(text);
    flash(Date.now(), `已复制${label}`);
  };

  // ---------- 粘贴 ----------
  const doPaste = (block: Block, mode: PasteMode = 'all', transpose = false, fromClip = false, opts: PasteOpts = {}) => {
    const g = sel.range;
    const bw = transpose ? block.rows : block.cols;
    const bh = transpose ? block.cols : block.rows;
    // 选区是块的整数倍时铺满（同 Excel），否则从左上角粘贴一份
    const tile = !transpose && !opts.op && (g.r1 - g.r0 + 1) % bh === 0 && (g.c1 - g.c0 + 1) % bw === 0 && (g.r1 - g.r0 + 1 > bh || g.c1 - g.c0 + 1 > bw);
    const cut = fromClip && clip?.cut ? clip : null;
    const target = run(
      cut ? '移动' : '粘贴',
      (tx) => {
        let out: Range = { r0: g.r0, c0: g.c0, r1: g.r0 + bh - 1, c1: g.c0 + bw - 1 };
        if (cut && cut.book === book) clearRange(tx, cut.si, cut.range, 'all');
        if (tile) {
          for (let r = g.r0; r <= g.r1; r += bh) for (let c = g.c0; c <= g.c1; c += bw) pasteBlock(tx, si, r, c, block, mode, false, opts);
          out = g;
        } else out = pasteBlock(tx, si, g.r0, g.c0, block, mode, transpose, opts);
        return out;
      },
      ui(),
    );
    if (cut) {
      clip = null;
      setMarquee(null);
    }
    if (target) selectRange(target);
  };
  const pasteText = (text: string, html?: string, mode: PasteMode = 'all', transpose = false, opts: PasteOpts = {}) => {
    const norm = (t: string) => t.replace(/\r\n?/g, '\n').replace(/\n$/, '');
    if (clip && norm(clip.tsv) === norm(text)) return doPaste(clip.block, mode, transpose, true, opts);
    const block = (html && blockFromHtml(html)) || blockFromText(text);
    if (!block.rows || (block.rows === 1 && block.cols === 1 && !block.cells[0][0] && !text)) return;
    doPaste(block, 'all', transpose);
  };
  /** 右键菜单里的粘贴：读系统剪贴板（读不到时用表格内部剪贴板） */
  const pasteFromMenu = async (mode: PasteMode = 'all', transpose = false, opts: PasteOpts = {}) => {
    setMenu(null);
    let text: string | null = null;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      text = null;
    }
    if (text === null) {
      if (clip) return doPaste(clip.block, mode, transpose, true, opts);
      host().toast('无法读取剪贴板，请按 Ctrl+V 粘贴');
      return;
    }
    pasteText(text, undefined, mode, transpose, opts);
    focusGrid();
  };

  // ---------- 编辑 ----------
  /** 编辑时显示的原文（同 Excel：公式带 =，日期按短日期，百分比带 %） */
  const rawText = (p: Pos): string => editText(book, sheet.cells[p.r]?.[p.c]);
  const startEdit = (mode: 'enter' | 'edit', text?: string, bar = false) => {
    const p = sel.active;
    const orig = rawText(p);
    setEditing({ r: p.r, c: p.c, text: text ?? orig, orig: text === undefined ? orig : `\u0000${orig}`, mode, bar });
    setMarquee(null);
    reveal(p);
  };
  const commit = async (move: [number, number] | null, all = false) => {
    const ed = editing;
    if (!ed) return;
    // 列属性（工作包）：不合规时按列的“输入不合规时”处理（拒绝 / 警告 / 仅提示）
    const cm = bundle && ed.text !== ed.orig && !ed.text.startsWith('=') ? metaAt(sheet, ed.r, ed.c) : undefined;
    const colMsg = cm ? checkCell(book, sheet, cm, ed.text) : null;
    if (cm && colMsg) {
      const label = cm.title || colName(ed.c);
      if (!cm.onInvalid || cm.onInvalid === 'stop') {
        host().toast(`「${ed.text}」不能填入${label}：${colMsg}`);
        return;
      }
      if (cm.onInvalid === 'warn' && !(await host().confirm(`${label}：${colMsg}\n\n仍然要保存「${ed.text}」吗？`))) return;
      if (cm.onInvalid === 'info') host().toast(`${label}：${colMsg}`);
    }
    // 数据验证：不符合时按“出错警告”的样式处理（停止：不让输入；警告：问是否继续；信息：提示后照常输入）
    const dv = ed.text !== ed.orig && !ed.text.startsWith('=') ? validationAt(sheet, ed.r, ed.c) : undefined;
    if (dv?.showError) {
      const cell = sheet.cells[ed.r]?.[ed.c];
      const p = parseInput(ed.text, book.styleBook ? formatCode(book.styleBook, cell?.s) : 'General', book.date1904);
      const msg = checkValue(dv, p.kind === 'value' ? p.v : null, book, si, ed.r, ed.c, props.engine);
      if (msg) {
        const text = dv.error || msg;
        if (!dv.errorStyle || dv.errorStyle === 'stop') {
          host().toast(`${dv.errorTitle || '输入值无效'}：${text}`);
          return;
        }
        if (dv.errorStyle === 'warning' && !(await host().confirm(`${dv.errorTitle || '警告'}\n\n${text}\n\n是否继续？`))) return;
        if (dv.errorStyle === 'information') host().toast(text);
      }
    }
    setEditing(null);
    const g = sel.range;
    if (ed.text !== ed.orig || all) {
      run(
        '输入',
        (tx) => {
          if (all) eachCell(sheet, g, (r, c) => setInput(tx, si, r, c, ed.text));
          else writeCell(tx, ed.r, ed.c, ed.text);
        },
        ui(),
      );
    }
    if (move) moveActive(move[0], move[1]);
    focusGrid();
  };
  /**
   * 写一个格子：在列属性的格式下输入（数字套这一列的格式，选项文字不被当成数字）；
   * 这一行是刚开始录入时，给其它有默认值、还空着的列填上默认值。
   */
  const writeCell = (tx: Tx, r: number, c: number, text: string) => {
    const m = bundle ? metaAt(sheet, r, c) : undefined;
    const startsRow = !!bundle && !!sheet.columns && r >= sheet.columns.header && !!text.trim() && rowIsBlank(sheet, r, c);
    if (m) setColumnInput(tx, si, r, c, m, text);
    else setInput(tx, si, r, c, text);
    if (startsRow) fillDefaults(tx, si, r, c, userName);
  };
  const pickValue = (r: number, c: number, text: string) => {
    setPicker(null);
    run('输入', (tx) => writeCell(tx, r, c, text), ui());
    focusGrid();
  };
  const toggleCheck = (r: number, c: number) => {
    run('切换复选框', (tx) => writeCell(tx, r, c, isCheckOn(colCellText(sheet, r, c)) ? '否' : '是'), ui());
  };
  /** 表头所在行（1 起）：已有列属性用它的，否则取光标所在数据区的第一行 */
  const headerRow = (): number => sheet.columns?.header ?? currentRegion(sheet, sel.active.r, sel.active.c).r0 + 1;
  const openColumnPanel = (c: number, at: { x: number; y: number }) => {
    if (editing) commit(null);
    setPicker(null);
    setColPanel({ c, ...at });
  };
  const applyColumn = (c: number, meta: ColumnMeta | null) => {
    const header = headerRow();
    setColPanel(null);
    run(
      meta ? '设置列属性' : '清除列属性',
      (tx) => {
        if (!sheet.columns) setHeaderRow(tx, si, header);
        setColumn(tx, si, c, meta);
      },
      ui(),
    );
    focusGrid();
  };
  const addOption = (meta: ColumnMeta, text: string) => {
    run('新建选项', (tx) => setColumn(tx, si, meta.col, { ...meta, options: [...(meta.options ?? []), { v: text }] }), ui());
  };
  const formsOf = () => (bundle ? bundle.forms().filter((f) => f.target.sheet === sheet.name) : []);
  const openForm = () => {
    if (editing) commit(null);
    setFormPanel({ formId: formsOf()[0]?.id ?? '' });
  };
  /** 生成离线填写页（自包含 .html）加进工作包：在工作包页导出后发给别人填 */
  const exportFormHtml = async (form: FormDef) => {
    if (!bundle) return;
    const fields = resolveForm(sheet.columns, form);
    const html = buildFormHtml({
      packTitle: bundle.packTitle,
      packId: bundle.packId,
      formId: form.id,
      title: form.title,
      fields,
      optionsOf: (f) => optionList(book, sheet, f.meta),
    });
    const item = await host().guard(bundle.addFile(`${form.title}-离线填写.html`, new TextEncoder().encode(html)));
    if (item) host().toast(`已生成「${item.name}」并加进工作包。到工作包页导出它，发给别人用浏览器填写；填完让对方发回导出的 .json，再点“导入填写结果…”`, 9000);
  };
  const importFormFile = async (form: FormDef) => {
    if (!bundle) return;
    const picked = await host().pickFiles?.('选择填写结果文件（离线填写页导出的 .json）', ['json']);
    if (!picked?.length) return;
    const file = await host().guard(host().readText(picked[0]!));
    if (!file) return;
    let result;
    try {
      result = parseFormResult(file.content);
    } catch (e) {
      return host().toast(e instanceof Error ? e.message : String(e));
    }
    if (result.form !== form.id) return host().toast('这份填写结果是别的表单收集的，不能导入到这个表单');
    if (result.pack && result.pack !== bundle.packId && !(await host().confirm('这份填写结果来自另一个工作包，仍然导入吗？'))) return;
    const fields = resolveForm(sheet.columns, form);
    let out = { imported: 0, skipped: [] as { index: number; reason: string }[] };
    run('导入填写结果', (tx) => void (out = importFormResult(tx, si, fields, result, userName)), ui());
    const bad = out.skipped.length ? `；${out.skipped.length} 条有问题没导入：${out.skipped.slice(0, 3).map((x) => `第 ${x.index} 条（${x.reason}）`).join('、')}${out.skipped.length > 3 ? '…' : ''}` : '';
    host().toast(`已导入 ${out.imported} 条${bad}`, 9000);
  };
  // 工作包页点了“表单录入”：表格打开（或已经打开）后弹出表单面板
  useEffect(() => {
    if (!bundle) return;
    const take = () => {
      const id = takeFormRequest(bundle.pack, bundle.itemId);
      if (id !== undefined) setFormPanel({ formId: bundle.forms().some((f) => f.id === id && f.target.sheet === sheet.name) ? id : '' });
    };
    take();
    const mine = formRequestKey(bundle.pack, bundle.itemId);
    return onFormRequest((key) => key === mine && take());
  }, [bundle, sheet.name]);
  /** 提交前校验脚本（表单设的 onValidate）：脚本只读，返回错误列表；不通过就不写入。脚本要先经过信任确认 */
  const validateByScript = async (fields: ResolvedField[], values: FormValues, target: FormTarget, form: FormDef | null): Promise<string[]> => {
    const scriptId = form?.onValidate;
    const source = scriptId ? bundle?.readText(scriptId) : undefined;
    if (!bundle || !scriptId || source === undefined || !host().runScript) return [];
    const byTitle: Record<string, string> = {};
    for (const f of fields) if (!f.hidden || values[f.meta.id]) byTitle[f.meta.title] = values[f.meta.id] ?? '';
    try {
      const r = host().runScript!(
        bundle.pack,
        scriptId,
        source,
        { log: (level, text) => level !== 'log' && host().toast(text, 4000) },
        { mode: 'validate', input: { form: form!.id, sheet: sheet.name, mode: target.kind, ...(target.kind === 'edit' ? { row: target.row + 1 } : {}), values: byTitle } },
      );
      return scriptErrors((await r.done).value);
    } catch (e) {
      return [`校验脚本没有运行成功：${e instanceof Error ? e.message : String(e)}`];
    }
  };
  const submitFromForm = async (fields: ResolvedField[], values: FormValues, target: FormTarget, form: FormDef | null): Promise<{ row: number } | { errors: string[] } | null> => {
    const errors = await validateByScript(fields, values, target, form);
    if (errors.length) return { errors };
    const res = run('表单录入', (tx) => submitForm(tx, si, fields, values, target, userName), ui());
    if (!res || !bundle) return null;
    bundle.logSubmission(form?.id ?? 'default', { t: new Date().toISOString(), user: userName, row: res.row + 1, mode: target.kind, values: res.written });
    // 滚到刚写入的行，让人看到录进去了
    const first = fields.find((f) => !f.hidden)?.meta.col ?? 0;
    selectCell({ r: res.row, c: first });
    // 有校验脚本的表单：马上存进包，下一条录入时校验脚本读到的就是最新数据（唯一性、对应关系等）
    if (form?.onValidate) void props.save(false);
    // 表单设了“提交后运行脚本”：运行（要先经过信任确认；脚本不能改正在打开的这张表，见 docs/WBP_FORMAT.md）
    const scriptId = form?.onSubmit;
    const source = scriptId ? bundle.readText(scriptId) : undefined;
    if (scriptId && source !== undefined && host().runScript) {
      const r = host().runScript!(bundle.pack, scriptId, source, { log: (level, text) => level !== 'log' && host().toast(text, 4000) });
      void r.done.then(
        (out) => host().toast(`脚本已运行${out.value !== undefined ? `：${out.value}` : ''}`),
        (e: unknown) => host().toast(`脚本没有运行成功：${e instanceof Error ? e.message : String(e)}`, 6000),
      );
    }
    return res;
  };
  const cancelEdit = () => {
    setEditing(null);
    focusGrid();
  };
  /** Enter / Tab 移动：选区多格时在选区里循环（同 Excel），否则正常移动 */
  const moveActive = (dr: number, dc: number) => {
    const g = sel.range;
    if (!rangeSel) {
      selectCell(moveBy(moveCtx, sel.active, dr, dc));
      return;
    }
    let { r, c } = sel.active;
    const m = merges.get(cellKey(r, c));
    if (dr) {
      r = dr > 0 ? (m ? m.r1 : r) + 1 : r - 1;
      if (r > g.r1) {
        r = g.r0;
        c = c + 1 > g.c1 ? g.c0 : c + 1;
      } else if (r < g.r0) {
        r = g.r1;
        c = c - 1 < g.c0 ? g.c1 : c - 1;
      }
    } else {
      c = dc > 0 ? (m ? m.c1 : c) + 1 : c - 1;
      if (c > g.c1) {
        c = g.c0;
        r = r + 1 > g.r1 ? g.r0 : r + 1;
      } else if (c < g.c0) {
        c = g.c1;
        r = r - 1 < g.r0 ? g.r1 : r - 1;
      }
    }
    const q = snap({ r, c }, merges);
    setSel({ ...sel, active: q });
    reveal(q);
  };

  // ---------- 行列 ----------
  const fullRows = sel.range.c0 === 0 && sel.range.c1 >= ext.cols - 1;
  const fullCols = sel.range.r0 === 0 && sel.range.r1 >= ext.rows - 1;
  const nRows = sel.range.r1 - sel.range.r0 + 1;
  const nCols = sel.range.c1 - sel.range.c0 + 1;
  const insertRows = (below = false) => run('插入行', (tx) => insertDelete(tx, si, 'row', below ? sel.range.r1 + 1 : sel.range.r0, nRows), ui());
  const insertCols = (right = false) => run('插入列', (tx) => insertDelete(tx, si, 'col', right ? sel.range.c1 + 1 : sel.range.c0, nCols), ui());
  const deleteRows = () => {
    run('删除行', (tx) => insertDelete(tx, si, 'row', sel.range.r0, -nRows), ui());
    selectCell({ r: sel.range.r0, c: sel.active.c });
  };
  const deleteCols = () => {
    run('删除列', (tx) => insertDelete(tx, si, 'col', sel.range.c0, -nCols), ui());
    selectCell({ r: sel.active.r, c: sel.range.c0 });
  };
  /**
   * 分页符（同 Excel）：在活动单元格上方插入水平分页符、左侧插入垂直分页符；活动单元格在第 1 行时只插垂直的，在 A 列时只插水平的。
   * 删除：去掉活动单元格上方 / 左侧的分页符。
   */
  const { r: ar, c: ac } = sel.active;
  const hasBreak = !!sheet.page?.rowBreaks?.includes(ar) || !!sheet.page?.colBreaks?.includes(ac);
  const pageBreakCmd = (insert: boolean) => {
    setMenu(null);
    run(
      insert ? '插入分页符' : '删除分页符',
      (tx) => {
        tx.meta(si);
        const pg = { ...sheet.page };
        const put = (k: 'rowBreaks' | 'colBreaks', at: number) => {
          const list = (pg[k] ?? []).filter((x) => x !== at);
          if (insert && at > 0) list.push(at);
          list.sort((a, b) => a - b);
          if (list.length) pg[k] = list;
          else delete pg[k];
        };
        if (!insert || ar > 0) put('rowBreaks', ar);
        if (!insert || ac > 0) put('colBreaks', ac);
        sheet.page = pg;
      },
      ui(),
    );
  };
  /** 复制 / 剪切后、虚线框还在时可用“插入复制（剪切）的单元格” */
  const insClip = marquee && clip ? clip : null;
  /**
   * 插入复制 / 剪切的单元格（同 Excel）：整行 → 在选区上方插入同样多的行；整列 → 在左侧插入列；
   * 其它 → 在选区左上角腾出块大小的位置（下移 / 右移）。复制的粘贴进去；剪切的整块移过去，原位置删掉、引用跟着走。
   */
  const insertCopiedCmd = (dir?: 'down' | 'right') => {
    setMenu(null);
    const c = insClip;
    if (!c) return;
    const g = sel.range;
    const b = c.block;
    const shift = c.whole ?? dir ?? 'down';
    const axis = shift === 'rows' || shift === 'down' ? 'row' : 'col';
    const at = axis === 'row' ? g.r0 : g.c0;
    const n = axis === 'row' ? b.rows : b.cols;
    const band: [number, number] | undefined = shift === 'down' ? [g.c0, g.c0 + b.cols - 1] : shift === 'right' ? [g.r0, g.r0 + b.rows - 1] : undefined;
    if (c.cut) {
      if (c.book !== book || c.si !== si) return host().toast('剪切的单元格只能插入到同一张工作表里');
      const out = run('插入剪切的单元格', (tx) => insertCut(tx, si, c.range, g.r0, g.c0, shift), ui());
      if (typeof out === 'string') return host().toast(out);
      clip = null;
      setMarquee(null);
      if (out) selectRange(out);
      focusGrid();
      return;
    }
    if (band) {
      const blocked = cellShiftBlocked(sheet, axis, at, n, band);
      if (blocked) return host().toast(blocked);
    }
    const target = run('插入复制的单元格', (tx) => insertCopied(tx, si, g.r0, g.c0, b, shift), ui());
    if (!target) return;
    // 来源在同一张表里、被挤开时虚线框跟着移动
    if (c.book === book && c.si === si) {
      const moved = shiftRange(c.range, axis, at, n, band);
      if (moved) {
        c.range = moved;
        setMarquee({ range: moved, cut: false });
      }
    }
    selectRange(target);
    focusGrid();
  };
  const clear = (what: 'contents' | 'formats' | 'all') =>
    run(what === 'contents' ? '清除内容' : what === 'formats' ? '清除格式' : '全部清除', (tx) => allRanges.forEach((g) => clearRange(tx, si, g, what)), ui());
  const hideRows = (hidden: boolean) => run(hidden ? '隐藏行' : '取消隐藏行', (tx) => setRowHeight(tx, si, sel.range.r0, sel.range.r1, { hidden }), ui());
  const hideCols = (hidden: boolean) => run(hidden ? '隐藏列' : '取消隐藏列', (tx) => setColWidth(tx, si, sel.range.c0, sel.range.c1, { hidden }), ui());
  const askRowHeight = async () => {
    setMenu(null);
    const cur = sheet.rowDefs?.[sel.active.r]?.ht ?? Math.round(sheet.defaultRowHeight * 0.75 * 4) / 4;
    const v = await host().askText('行高（磅，留空为自动）', String(cur));
    focusGrid();
    if (v === null) return;
    const pt = Number(v);
    if (v.trim() === '') run('行高', (tx) => setRowHeight(tx, si, sel.range.r0, sel.range.r1, { height: null }), ui());
    else if (pt >= 0 && pt <= 409) run('行高', (tx) => setRowHeight(tx, si, sel.range.r0, sel.range.r1, pt === 0 ? { hidden: true } : { height: Math.round((pt * 4) / 3) }), ui());
    else host().toast('行高需在 0–409 磅之间');
  };
  const askColWidth = async () => {
    setMenu(null);
    const d = sheet.colDefs?.find((x) => x.min <= sel.active.c && x.max >= sel.active.c);
    const cur = d?.width ?? Math.round((sheet.defaultColWidth / 7) * 100) / 100;
    const v = await host().askText('列宽（字符数）', String(cur));
    focusGrid();
    if (v === null) return;
    const w = Number(v);
    if (w >= 0 && w <= 255) run('列宽', (tx) => setColWidth(tx, si, sel.range.c0, sel.range.c1, w === 0 ? { hidden: true } : { width: Math.round(w * 7) }), ui());
    else host().toast('列宽需在 0–255 之间');
  };

  // ---------- 格式 ----------
  const activeS = styleAt(sheet, sel.active.r, sel.active.c);
  const activeSt = (activeS !== undefined ? book.styles[activeS] : undefined) ?? {};
  const activeFmt = book.styleBook ? formatCode(book.styleBook, activeS) : 'General';
  const styleCmd = (ch: StyleChange, label: string) => {
    run(
      label,
      (tx) => {
        for (const g of allRanges) {
          applyStyle(tx, si, g, ch);
          if (ch.fontSize) growRows(tx, si, g.r0, g.r1, autoRowPt(ch.fontSize));
        }
      },
      ui(),
    );
  };
  const borderCmd = (kind: BorderKind, color?: string) => run('边框', (tx) => allRanges.forEach((g) => applyBorders(tx, si, g, kind, color)), ui());
  const mergeCmd = async (kind: MergeKind) => {
    const g = sel.range;
    if (kind === 'unmerge') return void run('取消合并', (tx) => mergeCells(tx, si, g, false), ui());
    if (g.r0 === g.r1 && g.c0 === g.c1) return;
    if (kind === 'across' && g.c0 === g.c1) return;
    const lost = kind === 'across' ? [...Array(g.r1 - g.r0 + 1)].reduce((n: number, _, i) => n + lostByMerge(sheet, { ...g, r0: g.r0 + i, r1: g.r0 + i }), 0) : lostByMerge(sheet, g);
    if (lost && !(await host().confirm('合并单元格时，仅保留左上角的值，其它值将被放弃。\n\n确定合并吗？'))) return;
    run(
      '合并单元格',
      (tx) => {
        if (kind === 'across') mergeAcross(tx, si, g);
        else mergeCells(tx, si, g, true);
        if (kind === 'center') applyStyle(tx, si, { r0: g.r0, c0: g.c0, r1: g.r0, c1: g.c0 }, { align: 'center' });
      },
      ui(),
    );
    selectRange(g);
  };
  const numFmtCmd = (code: string) => styleCmd({ numFmt: code }, '数字格式');
  /** 设置单元格格式对话框确定：改过的格式、边框、合并在一个撤销步骤里应用到选区 */
  const formatCellsCmd = async (r: FormatCellsResult) => {
    setFmtDlg(false);
    const g = sel.range;
    let merge = r.merge;
    // 单个格子不能合并；合并会丢掉左上角以外的值时先问
    if (merge && g.r0 === g.r1 && g.c0 === g.c1) merge = undefined;
    if (merge && lostByMerge(sheet, g) > 0 && !(await host().confirm('合并单元格时，仅保留左上角的值，其它值将被放弃。\n\n确定合并吗？'))) merge = undefined;
    if (!Object.keys(r.style).length && !r.border && merge === undefined) return focusGrid();
    run(
      '设置单元格格式',
      (tx) => {
        if (merge !== undefined) mergeCells(tx, si, g, merge);
        for (const x of allRanges) {
          if (Object.keys(r.style).length) applyStyle(tx, si, x, r.style);
          if (r.style.fontSize) growRows(tx, si, x.r0, x.r1, autoRowPt(r.style.fontSize));
          if (r.border) applyBorders(tx, si, x, r.border.kind, r.border.color);
        }
      },
      ui(),
    );
    focusGrid();
  };
  const customFmtCmd = async () => {
    const v = await host().askText('数字格式代码（如 0.00"吨"、yyyy-mm-dd、#,##0.00;[红色]-#,##0.00）', activeFmt);
    focusGrid();
    if (v && v.trim()) numFmtCmd(v.trim().replace(/\[红色\]/g, '[Red]'));
  };
  const decimalsCmd = (dir: 1 | -1) => {
    const v = sheet.cells[sel.active.r]?.[sel.active.c]?.v;
    const next = stepDecimals(activeFmt, dir, typeof v === 'number' ? v : undefined);
    if (next !== activeFmt) numFmtCmd(next);
  };
  // ---------- 工作表 / 排序 / 筛选 ----------
  const bookCmd = (label: string, fn: (tx: Tx) => number | boolean | void, go?: (out: number | boolean | void) => number | undefined) => {
    if (editing) commit(null);
    if (structureLocked(book)) return host().toast('工作簿结构已保护，不能新建、删除、重命名、移动或隐藏工作表（右键 → 撤销工作簿结构保护）');
    const out = run(label, fn, ui());
    const next = go?.(out);
    if (next !== undefined && next !== si) onSheet(next);
  };
  const sheetName = (i: number, name: string): string | null => {
    const err = sheetNameError(book, name, i);
    if (err) return err;
    if (name.trim() !== book.sheets[i].name) bookCmd('重命名工作表', (tx) => renameSheet(tx, i, name));
    return null;
  };
  /** 排序 / 筛选用的区域：多选时用选区，否则用活动单元格所在的数据区域（在筛选区域里时用筛选区域） */
  const dataRegion = (): { g: Range; header: boolean } | null => {
    const af = sheet.autoFilter;
    if (af && contains(af.ref, sel.active.r, sel.active.c) && !rangeSel) return { g: af.ref, header: true };
    const g = rangeSel ? sel.range : currentRegion(sheet, sel.active.r, sel.active.c);
    if (!rangeSel && !sheet.cells[sel.active.r]?.[sel.active.c] && g.r0 === g.r1 && g.c0 === g.c1) return null;
    const bold = (r: number, c: number) => !!book.styles[styleAt(sheet, r, c) ?? -1]?.bold;
    return { g, header: guessHeader(sheet, g, bold) };
  };
  const doSort = (g: Range, keys: SortKey[], header: boolean) => {
    const ok = run('排序', (tx) => sortRange(tx, si, g, keys, header), ui());
    if (ok === false) host().toast('区域里有合并单元格，不能排序');
  };
  const sortCmd = (desc: boolean) => {
    setSortMenu(null);
    const d = dataRegion();
    if (!d) return host().toast('请先选中要排序的数据');
    doSort(d.g, [{ c: sel.active.c, desc }], d.header);
  };
  const filterToggle = () => {
    setSortMenu(null);
    if (sheet.autoFilter) return void run('取消筛选', (tx) => setAutoFilter(tx, si, null), ui());
    const d = dataRegion();
    if (!d) return host().toast('请先选中要筛选的数据（第一行为标题）');
    run('筛选', (tx) => setAutoFilter(tx, si, d.g), ui());
  };
  const columnFilter = (c: number, f: ColumnFilter | null) => {
    setFilterMenu(null);
    const af = sheet.autoFilter;
    if (!af) return;
    run(
      '筛选',
      (tx) => {
        // 下面接着有新数据时筛选区域跟着扩大（同 Excel）
        const g = currentRegion(sheet, af.ref.r0, af.ref.c0);
        if (g.r1 > af.ref.r1) {
          tx.meta(si);
          sheet.autoFilter = { ...af, ref: { ...af.ref, r1: g.r1 } };
        }
        setColumnFilter(tx, si, c, f);
      },
      ui(),
    );
    focusGrid();
  };
  // ---------- 自动求和 / 函数提示 ----------
  /** Alt+=：单格时插入 =SUM(上方或左边的数字区) 进入编辑；选了一片数字时在每列下面写合计 */
  const autoSum = () => {
    if (editing) return;
    const g = sel.range;
    if (rangeSel && g.r1 > g.r0) {
      run(
        '自动求和',
        (tx) => {
          for (let c = g.c0; c <= g.c1; c++) {
            if (sheet.cells[g.r1 + 1]?.[c]) continue;
            setInput(tx, si, g.r1 + 1, c, `=SUM(${colName(c)}${g.r0 + 1}:${colName(c)}${g.r1 + 1})`);
          }
        },
        ui(),
      );
      return;
    }
    const src = autoSumRange(sheet, sel.active.r, sel.active.c);
    const text = src ? `=SUM(${rangeName(src)})` : '=SUM()';
    startEdit('edit', text);
    const at = src ? text.length : text.length - 1;
    setCaretAt(at);
    requestAnimationFrame(() => kb.current?.setSelectionRange(at, at));
  };
  const hintTyping = editing ? typingName(editing.text, caretAt) : null;
  const hints = hintTyping && hintOff !== editing?.text ? suggest(hintTyping.prefix) : [];
  const call = editing && !hints.length ? callAt(editing.text, caretAt) : null;
  /** 选中提示里的函数：替换正在输入的名字，补上左括号 */
  const acceptHint = (name: string, el: HTMLTextAreaElement | null) => {
    if (!editing || !hintTyping) return;
    const text = editing.text.slice(0, hintTyping.start) + name + '(' + editing.text.slice(caretAt);
    const at = hintTyping.start + name.length + 1;
    setEditing({ ...editing, text, mode: 'edit' });
    setCaretAt(at);
    setHintIdx(0);
    requestAnimationFrame(() => el?.setSelectionRange(at, at));
  };
  /** 提示列表打开时的按键：↑↓ 选择，Tab 确认，Esc 关闭；处理了返回 true */
  const hintKey = (e: React.KeyboardEvent<HTMLTextAreaElement>): boolean => {
    if (!hints.length) return false;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      setHintIdx((i) => (i + (e.key === 'ArrowDown' ? 1 : hints.length - 1)) % hints.length);
      return true;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      acceptHint(hints[Math.min(hintIdx, hints.length - 1)], e.currentTarget);
      return true;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      setHintOff(editing?.text ?? null);
      return true;
    }
    return false;
  };
  const trackCaret = (el: HTMLTextAreaElement) => {
    if (el.selectionStart !== caretAt) setCaretAt(el.selectionStart);
  };

  // ---------- 查找替换 / 打印 ----------
  const gotoHit2 = (h: FindHit) => {
    if (h.si !== si) {
      props.pendingGoto.current = h;
      onSheet(h.si);
    } else selectCell({ r: h.r, c: h.c });
  };
  useEffect(() => {
    const h = props.pendingGoto.current;
    if (h && h.si === si) {
      props.pendingGoto.current = null;
      selectCell({ r: h.r, c: h.c });
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const replaceOne = (h: FindHit, o: FindOptions, repl: string) => !!run('替换', (tx) => replaceCell(tx, h, o, repl), ui());
  const replaceEvery = (o: FindOptions, repl: string) => run('全部替换', (tx) => replaceAllCells(tx, o, repl), ui()) ?? 0;
  const painterCmd = (sticky: boolean) => {
    if (painter && !sticky) return setPainter(null);
    setPainter({ block: copyBlock(book, si, sel.range), mode: sticky ? 'sticky' : 'once' });
  };

  // ---------- 搜索 ----------
  const hits = useMemo(() => findAll(sheet, query), [sheet, query, version]); // eslint-disable-line react-hooks/exhaustive-deps
  /** 批注（笔记）、超链接、图片地址 */
  const noteMap = useMemo(() => new Map((sheet.notes ?? []).map((n) => [cellKey(n.r, n.c), n])), [sheet, version]); // eslint-disable-line react-hooks/exhaustive-deps
  const linkAt = (r: number, c: number) => sheet.links?.find((l) => contains(l.ref, r, c));
  const imageUrls = useMemo(() => (sheet.drawings ?? []).map((d) => (d.image ? URL.createObjectURL(new Blob([d.image.data as BlobPart], { type: d.image.mime })) : '')), [sheet.drawings]);
  useEffect(() => () => imageUrls.forEach((u) => u && URL.revokeObjectURL(u)), [imageUrls]);
  /** 打开超链接：外部网址用系统浏览器；本工作簿里的位置（Sheet2!A1）就跳过去 */
  const followLink = (r: number, c: number): boolean => {
    const l = linkAt(r, c);
    if (!l) return false;
    if (l.url) void host().openUrl(l.url).catch(() => host().toast('无法打开链接'));
    else if (l.location) {
      const m = l.location.match(/^(?:'?(.+?)'?!)?\$?([A-Za-z]+\$?\d+)/);
      const target = m?.[1] ? book.sheets.findIndex((s) => s.name === m[1]) : si;
      const g = m ? parseRef(m[2]) : null;
      if (target >= 0 && g) {
        if (target !== si) {
          props.pendingGoto.current = { si: target, r: g.r0, c: g.c0 };
          onSheet(target);
        } else selectCell({ r: g.r0, c: g.c0 });
      }
    }
    return true;
  };
  /** 条件格式的结果（值变了就重算） */
  const cfMap = useMemo(() => evalCf(book, si, props.engine), [book, si, sheet, version, props.engine]); // eslint-disable-line react-hooks/exhaustive-deps
  const hitKeys = useMemo(() => new Set(hits.map((p) => cellKey(p.r, p.c))), [hits]);
  const gotoHit = (dir: 1 | -1) => {
    if (!hits.length) return;
    const i = hitIdx < 0 ? (dir > 0 ? 0 : hits.length - 1) : (hitIdx + dir + hits.length) % hits.length;
    setHitIdx(i);
    selectCell(hits[i]);
  };
  useEffect(() => setHitIdx(-1), [query]);

  // ---------- 坐标 ----------
  /** 列 / 行的起点在 sheet-view 里的坐标（考虑冻结和滚动） */
  const xOf = (c: number) => hw + (c < fc ? cols.pos[c] : cols.pos[c] - scroll.x);
  const yOf = (r: number) => hh + (r < fr ? rows.pos[r] : rows.pos[r] - scroll.y);

  // ---------- 鼠标 ----------
  type Hit = { r: number; c: number; area: 'cell' | 'col' | 'row' | 'corner' };
  const hitTest = (clientX: number, clientY: number): Hit => {
    const rect = view.current!.getBoundingClientRect();
    const x = clientX - rect.left;
    const y = clientY - rect.top;
    const sx = x - hw;
    const sy = y - hh;
    const c = indexAt(cols, sx < fw ? Math.max(0, sx) : sx + scroll.x);
    const r = indexAt(rows, sy < fh ? Math.max(0, sy) : sy + scroll.y);
    const area = x < hw && y < hh ? 'corner' : y < hh ? 'col' : x < hw ? 'row' : 'cell';
    return { r, c, area };
  };
  const drag = useRef<{ mode: 'cell' | 'col' | 'row' | 'fill' | 'point'; anchor: Pos; active: Pos; x: number; y: number; extra?: Range[] } | null>(null);
  /** 编辑公式时用鼠标点选的引用：在文字里的位置与起点格 */
  const pointRef = useRef<{ start: number; end: number; anchor: Pos } | null>(null);
  const editorEl = () => (editing?.bar ? barInput.current : kb.current);
  /** 把点选的格子 / 区域写进公式（替换 span 那一段），光标放到它后面 */
  const insertPointRef = (span: { start: number; end: number }, a: Pos, b: Pos) => {
    if (!editing) return;
    const g = expandToMerges(rangeOf(a, b), sheet.merges);
    const addr = rangeName(g);
    const text = editing.text.slice(0, span.start) + addr + editing.text.slice(span.end);
    pointRef.current = { start: span.start, end: span.start + addr.length, anchor: a };
    setEditing({ ...editing, text });
    const el = editorEl();
    const at = span.start + addr.length;
    requestAnimationFrame(() => el?.setSelectionRange(at, at));
  };
  /** F4：光标所在的引用切换绝对 / 相对 */
  const toggleAbsolute = (el: HTMLTextAreaElement) => {
    if (!editing) return;
    const caret = el.selectionStart;
    const hit = refSpans(editing.text).find((x) => caret >= x.start && caret <= x.end);
    if (!hit) return;
    const next = refStr(cycleAbsolute(hit.ref));
    const text = editing.text.slice(0, hit.start) + next + editing.text.slice(hit.end);
    setEditing({ ...editing, text });
    const at = hit.start + next.length;
    requestAnimationFrame(() => el.setSelectionRange(at, at));
  };

  /** 鼠标在选区右下角的填充柄上 */
  const onHandle = (clientX: number, clientY: number) => {
    const rect = view.current?.getBoundingClientRect();
    if (!rect || editing) return false;
    const g = sel.range;
    const x = clientX - rect.left - xOf(g.c1 + 1);
    const y = clientY - rect.top - yOf(g.r1 + 1);
    return Math.abs(x) <= 5 && Math.abs(y) <= 5;
  };

  /** 鼠标在列标两列之间（col）/ 行号两行之间（row）的分隔线上：返回线左 / 上侧的列 / 行 */
  type Edge = { axis: 'col' | 'row'; i: number };
  const edgeTest = (clientX: number, clientY: number): Edge | null => {
    const rect = view.current!.getBoundingClientRect();
    const x = clientX - rect.left;
    const y = clientY - rect.top;
    if (y < hh && x >= hw) {
      const sx = x - hw;
      const i = edgeAt(cols, sx < fw ? sx : sx + scroll.x);
      return i < 0 ? null : { axis: 'col', i };
    }
    if (x < hw && y >= hh) {
      const sy = y - hh;
      const i = edgeAt(rows, sy < fh ? sy : sy + scroll.y);
      return i < 0 ? null : { axis: 'row', i };
    }
    return null;
  };
  /** 正在拖动的分隔线：起点鼠标坐标与起始尺寸（未缩放像素） */
  const resize = useRef<(Edge & { start: number; base: number; px: number }) | null>(null);
  const [edgeHover, setEdgeHover] = useState<'col' | 'row' | 'fill' | null>(null);
  /** 拖动时在鼠标旁提示当前宽度 / 高度 */
  const [resizeTip, setResizeTip] = useState<{ x: number; y: number; text: string } | null>(null);

  const pointAt = (h: Hit, mode: 'cell' | 'col' | 'row'): Pos =>
    mode === 'col' ? { r: ext.rows - 1, c: h.c } : mode === 'row' ? { r: h.r, c: ext.cols - 1 } : { r: h.r, c: h.c };

  const onMouseDown = (e: React.MouseEvent) => {
    if (editing && !editing.bar && (e.target as HTMLElement) === kb.current) return;
    const hitDrawing = drawingAtPoint(e.clientX, e.clientY);
    if (hitDrawing && !(editing?.text.startsWith('='))) {
      // 点在图片 / 图表上：选中它（左键还能拖动），不动单元格选区
      if (editing) commit(null);
      focusGrid();
      setSelDrawing(hitDrawing.i);
      if (e.button === 0) {
        e.preventDefault();
        startDrawingDrag(hitDrawing.i, 'move', e.clientX, e.clientY);
      }
      return;
    }
    if (selDrawing !== null) setSelDrawing(null);
    if (e.button === 2) {
      const h = hitTest(e.clientX, e.clientY);
      if (editing) commit(null);
      if (h.area === 'cell' && !contains(sel.range, h.r, h.c)) selectCell({ r: h.r, c: h.c }, false);
      else if (h.area === 'col' && !(fullCols && h.c >= sel.range.c0 && h.c <= sel.range.c1)) extendTo({ r: 0, c: h.c }, { r: 0, c: h.c }, { r: ext.rows - 1, c: h.c }, false);
      else if (h.area === 'row' && !(fullRows && h.r >= sel.range.r0 && h.r <= sel.range.r1)) extendTo({ r: h.r, c: 0 }, { r: h.r, c: 0 }, { r: h.r, c: ext.cols - 1 }, false);
      return;
    }
    if (e.button !== 0) return;
    e.preventDefault();
    // 手机选区模式：点格子把选区扩展到这里
    if (mobile && mRange && !editing) {
      const h = hitTest(e.clientX, e.clientY);
      if (h.area === 'cell') extendTo(sel.anchor, sel.anchor, snap({ r: h.r, c: h.c }, merges));
      return;
    }
    // 编辑公式时点格子：把引用写进公式（同 Excel 的点选模式）
    if (editing?.text.startsWith('=')) {
      const h = hitTest(e.clientX, e.clientY);
      const el = editorEl();
      const caret = el?.selectionStart ?? editing.text.length;
      const pr = pointRef.current;
      const span = pr && caret === pr.end ? pr : canInsertRef(editing.text, caret) ? { start: caret, end: caret } : null;
      if (h.area === 'cell' && span) {
        const p = snap({ r: h.r, c: h.c }, merges);
        insertPointRef(span, p, p);
        drag.current = { mode: 'point', anchor: p, active: p, x: e.clientX, y: e.clientY };
        return;
      }
    }
    if (editing) commit(null);
    focusGrid();
    if (onHandle(e.clientX, e.clientY)) {
      drag.current = { mode: 'fill', anchor: sel.anchor, active: sel.active, x: e.clientX, y: e.clientY };
      setFillTo(sel.range);
      return;
    }
    const edge = edgeTest(e.clientX, e.clientY);
    if (edge) {
      const base = sizeOf(edge.axis === 'col' ? cols : rows, edge.i) / zoom;
      resize.current = { ...edge, start: edge.axis === 'col' ? e.clientX : e.clientY, base, px: base };
      // 拖出表头时光标也保持不变
      document.body.style.cursor = edge.axis === 'col' ? 'col-resize' : 'row-resize';
      dragMove.current(e.clientX, e.clientY);
      return;
    }
    const h = hitTest(e.clientX, e.clientY);
    if (h.area === 'corner') return selectAll();
    // 复选框列：点方框本身直接切换
    if (bundle && h.area === 'cell' && !e.shiftKey && !e.ctrlKey && !e.metaKey && metaAt(sheet, h.r, h.c)?.type === 'check') {
      const left = view.current!.getBoundingClientRect().left;
      if (Math.abs(e.clientX - left - (xOf(h.c) + sizeOf(cols, h.c) / 2)) <= 11 * zoom) {
        selectCell({ r: h.r, c: h.c }, false);
        toggleCheck(h.r, h.c);
        return;
      }
    }
    const mode = h.area === 'col' ? 'col' : h.area === 'row' ? 'row' : 'cell';
    if (e.shiftKey) {
      extendTo(sel.anchor, sel.active, pointAt(h, mode), false);
      drag.current = { mode, anchor: sel.anchor, active: sel.active, x: e.clientX, y: e.clientY };
      return;
    }
    if ((e.ctrlKey || e.metaKey) && mode === 'cell' && followLink(h.r, h.c)) return;
    const anchor = mode === 'col' ? { r: 0, c: h.c } : mode === 'row' ? { r: h.r, c: 0 } : snap({ r: h.r, c: h.c }, merges);
    // Ctrl+点击：原来的选区保留，再加一块（不连续选区）
    const extra = e.ctrlKey || e.metaKey ? [...(sel.extra ?? []), sel.range] : undefined;
    if (mode === 'cell' && !extra) selectCell(anchor, false);
    else extendTo(anchor, anchor, pointAt(h, mode), false, extra);
    drag.current = { mode, anchor, active: anchor, x: e.clientX, y: e.clientY, extra };
  };

  /** 填充柄拖到 h：沿离开选区较远的方向扩展选区 */
  const fillTarget = (h: Hit): Range => {
    const g = sel.range;
    const dr = h.r > g.r1 ? h.r - g.r1 : h.r < g.r0 ? h.r - g.r0 : 0;
    const dc = h.c > g.c1 ? h.c - g.c1 : h.c < g.c0 ? h.c - g.c0 : 0;
    if (!dr && !dc) return g;
    if (Math.abs(dr) >= Math.abs(dc)) return dr > 0 ? { ...g, r1: h.r } : { ...g, r0: h.r };
    return dc > 0 ? { ...g, c1: h.c } : { ...g, c0: h.c };
  };

  const dragMove = useRef<(x: number, y: number) => void>(() => {});
  dragMove.current = (x, y) => {
    const dd = dragDrawing.current;
    if (dd) {
      const dx = (x - dd.sx) / zoom;
      const dy = (y - dd.sy) / zoom;
      let { x: rx, y: ry, w: rw, h: rh } = dd.rect;
      const MIN = 24;
      if (dd.dir === 'move') {
        rx += dx;
        ry += dy;
      } else {
        if (dd.dir.includes('e')) rw = Math.max(MIN, rw + dx);
        if (dd.dir.includes('s')) rh = Math.max(MIN, rh + dy);
        if (dd.dir.includes('w')) {
          const nw = Math.max(MIN, rw - dx);
          rx += rw - nw;
          rw = nw;
        }
        if (dd.dir.includes('n')) {
          const nh = Math.max(MIN, rh - dy);
          ry += rh - nh;
          rh = nh;
        }
      }
      ghostRef.current = { x: Math.max(0, rx), y: Math.max(0, ry), w: rw, h: rh };
      setDrawGhost(ghostRef.current);
      return;
    }
    const z = resize.current;
    if (z) {
      const px = Math.max(MIN_SIZE, Math.round(z.base + ((z.axis === 'col' ? x : y) - z.start) / zoom));
      z.px = px;
      (z.axis === 'col' ? setFit : setRowFit)({ [z.i]: px });
      const rect = view.current!.getBoundingClientRect();
      const label = z.axis === 'col' ? `宽度：${Math.round((px / 7) * 100) / 100}（${px} 像素）` : `高度：${Math.round(px * 0.75 * 4) / 4}（${px} 像素）`;
      setResizeTip({ x: x - rect.left, y: y - rect.top, text: label });
      return;
    }
    const d = drag.current;
    if (!d) return;
    d.x = x;
    d.y = y;
    if (d.mode === 'fill') {
      setFillTo(fillTarget(hitTest(x, y)));
      return;
    }
    if (d.mode === 'point') {
      const pr = pointRef.current;
      const h = hitTest(x, y);
      if (pr) insertPointRef(pr, pr.anchor, { r: h.r, c: h.c });
      return;
    }
    extendTo(d.anchor, d.active, pointAt(hitTest(x, y), d.mode), false, d.extra);
  };
  const dragEnd = useRef<() => void>(() => {});
  dragEnd.current = () => {
    const dd = dragDrawing.current;
    if (dd) {
      dragDrawing.current = null;
      const g = ghostRef.current;
      ghostRef.current = null;
      setDrawGhost(null);
      // 没怎么动（只是点了一下选中）就不改
      if (g && (Math.abs(g.x - dd.rect.x) > 1 || Math.abs(g.y - dd.rect.y) > 1 || Math.abs(g.w - dd.rect.w) > 1 || Math.abs(g.h - dd.rect.h) > 1)) {
        const from = anchorAtPx(g.x * zoom, g.y * zoom);
        const to = anchorAtPx((g.x + g.w) * zoom, (g.y + g.h) * zoom);
        run('移动图表', (tx) => moveDrawing(tx, si, dd.i, from, to), ui());
      }
      return;
    }
    const d = drag.current;
    drag.current = null;
    if (painter && d && (d.mode === 'cell' || d.mode === 'col' || d.mode === 'row')) {
      doPaste(painter.block, 'formats');
      if (painter.mode === 'once') setPainter(null);
    }
    if (d?.mode === 'fill') {
      const to = fillTo;
      setFillTo(null);
      const g = sel.range;
      if (to && (to.r0 !== g.r0 || to.r1 !== g.r1 || to.c0 !== g.c0 || to.c1 !== g.c1)) {
        run('填充', (tx) => fillSeries(tx, si, g, to), ui());
        selectRange(to);
      }
    }
    const z = resize.current;
    if (z) {
      resize.current = null;
      document.body.style.cursor = '';
      setResizeTip(null);
      // 拖动的列 / 行在整列 / 整行选区里时，选中的都改成同样的尺寸（同 Excel）
      if (z.axis === 'col') {
        const all = fullCols && z.i >= sel.range.c0 && z.i <= sel.range.c1;
        run('列宽', (tx) => setColWidth(tx, si, all ? sel.range.c0 : z.i, all ? sel.range.c1 : z.i, { width: z.px }), ui());
        setFit({});
      } else {
        const all = fullRows && z.i >= sel.range.r0 && z.i <= sel.range.r1;
        run('行高', (tx) => setRowHeight(tx, si, all ? sel.range.r0 : z.i, all ? sel.range.r1 : z.i, { height: z.px }), ui());
        setRowFit({});
      }
    }
  };
  useEffect(() => {
    const move = (e: MouseEvent) => dragMove.current(e.clientX, e.clientY);
    const up = () => dragEnd.current();
    // 拖出可视区时自动滚动
    const timer = window.setInterval(() => {
      const d = drag.current;
      const el = scroller.current;
      if (!d || !el || !view.current) return;
      const rect = view.current.getBoundingClientRect();
      const dx = d.x < rect.left + hw ? d.x - rect.left - hw : d.x > rect.left + vpRef.current.w ? d.x - rect.left - vpRef.current.w : 0;
      const dy = d.y < rect.top + hh ? d.y - rect.top - hh : d.y > rect.top + vpRef.current.h ? d.y - rect.top - vpRef.current.h : 0;
      if (!dx && !dy) return;
      if (d.mode !== 'row') el.scrollLeft += dx;
      if (d.mode !== 'col') el.scrollTop += dy;
      dragMove.current(d.x, d.y);
    }, 40);
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
    return () => {
      window.clearInterval(timer);
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      if (resize.current) document.body.style.cursor = '';
    };
  }, [hw, hh]);
  const vpRef = useRef(vp);
  vpRef.current = vp;

  /** 双击填充柄：按左边（或右边）相邻列的数据长度向下填充（同 Excel） */
  const autoFillDown = () => {
    const g = sel.range;
    const filled = (r: number, c: number) => !!sheet.cells[r]?.[c];
    const side = g.c0 > 0 && filled(g.r1 + 1, g.c0 - 1) ? g.c0 - 1 : filled(g.r1 + 1, g.c1 + 1) ? g.c1 + 1 : -1;
    if (side < 0) return;
    let end = g.r1 + 1;
    while (filled(end + 1, side)) end++;
    const to = { ...g, r1: end };
    run('填充', (tx) => fillSeries(tx, si, g, to), ui());
    selectRange(to);
  };

  // ---------- 右键菜单 / 详情 ----------
  useEffect(() => {
    if (!menu && !freezeMenu && !detail && !sortMenu && !nameList) return;
    host().pushPopup();
    const close = () => {
      setMenu(null);
      setFreezeMenu(null);
      setSortMenu(null);
    };
    window.addEventListener('click', close);
    window.addEventListener('blur', close);
    return () => {
      host().popPopup();
      window.removeEventListener('click', close);
      window.removeEventListener('blur', close);
    };
  }, [menu, freezeMenu, detail, sortMenu, nameList]);

  // ---------- 键盘 ----------
  const pageRows = Math.max(1, Math.floor(zoneH / Math.max(1, sheet.defaultRowHeight * zoom)) - 1);
  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (e.nativeEvent.isComposing || k === 'Process') return;
    if (editing) {
      if (hintKey(e)) return;
      if (k === 'Enter' && e.altKey) {
        // Alt+Enter：格子里换行
        e.preventDefault();
        const el = e.currentTarget;
        const t = editing.text;
        const a = el.selectionStart;
        const next = t.slice(0, a) + '\n' + t.slice(el.selectionEnd);
        setEditing({ ...editing, text: next, mode: 'edit' });
        requestAnimationFrame(() => el.setSelectionRange(a + 1, a + 1));
        return;
      }
      if (k === 'Enter') {
        e.preventDefault();
        if (ctrl) commit(null, true);
        else commit([e.shiftKey ? -1 : 1, 0]);
        return;
      }
      if (k === 'Tab') {
        e.preventDefault();
        commit([0, e.shiftKey ? -1 : 1]);
        return;
      }
      if (k === 'Escape') {
        e.preventDefault();
        cancelEdit();
        return;
      }
      if (k === 'F2') {
        e.preventDefault();
        setEditing({ ...editing, mode: editing.mode === 'enter' ? 'edit' : 'enter' });
        return;
      }
      if (k === 'F4') {
        e.preventDefault();
        toggleAbsolute(e.currentTarget);
        return;
      }
      const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (arrows[k] && editing.mode === 'enter' && !e.shiftKey) {
        e.preventDefault();
        commit(arrows[k]);
      }
      return;
    }
    if (selDrawing !== null && (k === 'Delete' || k === 'Backspace')) {
      e.preventDefault();
      const i = selDrawing;
      setSelDrawing(null);
      run('删除图表', (tx) => deleteDrawing(tx, si, i), ui());
      return;
    }
    if (selDrawing !== null && k === 'Escape') {
      setSelDrawing(null);
      return;
    }
    if (k === 'F2' && e.shiftKey) {
      e.preventDefault();
      openNote(sel.active.r, sel.active.c);
      return;
    }
    if (ctrl && k === 'F3') {
      e.preventDefault();
      setNamesDlg({});
      return;
    }
    if (e.altKey && e.shiftKey && (k === 'ArrowRight' || k === 'ArrowLeft')) {
      e.preventDefault();
      groupSelection(k === 'ArrowRight');
      return;
    }
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) {
      e.preventDefault();
      const [dr, dc] = arrows[k];
      if (e.shiftKey) extendTo(sel.anchor, sel.active, moveBy(moveCtx, sel.end, dr, dc, ctrl));
      else selectCell(moveBy(moveCtx, sel.active, dr, dc, ctrl));
      return;
    }
    if ((k === 'PageDown' || k === 'PageUp') && ctrl) {
      // Ctrl+PageUp / PageDown：切换工作表
      e.preventDefault();
      const list = book.sheets.map((s, i) => (s.hidden ? -1 : i)).filter((i) => i >= 0);
      const at = list.indexOf(si);
      const next = list[at + (k === 'PageDown' ? 1 : -1)];
      if (next !== undefined) onSheet(next);
      return;
    }
    if (k === 'F11' && e.shiftKey) {
      e.preventDefault();
      bookCmd('插入工作表', (tx) => addSheet(tx, si), (i) => i as number);
      return;
    }
    if (k === 'PageDown' || k === 'PageUp') {
      e.preventDefault();
      const r = Math.max(0, Math.min(ext.rows - 1, (e.shiftKey ? sel.end : sel.active).r + (k === 'PageDown' ? pageRows : -pageRows)));
      if (e.shiftKey) extendTo(sel.anchor, sel.active, { r, c: sel.end.c });
      else selectCell({ r, c: sel.active.c });
      return;
    }
    if (k === 'Home') {
      e.preventDefault();
      selectCell(ctrl ? { r: 0, c: 0 } : { r: sel.active.r, c: 0 });
      return;
    }
    if (k === 'End' && ctrl) {
      e.preventDefault();
      selectCell({ r: Math.max(0, sheet.rows - 1), c: Math.max(0, sheet.cols - 1) });
      return;
    }
    if (k === 'Enter' || k === 'Tab') {
      e.preventDefault();
      const back = e.shiftKey ? -1 : 1;
      moveActive(k === 'Enter' ? back : 0, k === 'Tab' ? back : 0);
      return;
    }
    // 复选框列：空格切换
    if (k === ' ' && !ctrl && !e.shiftKey && !e.altKey && activeMeta?.type === 'check') {
      e.preventDefault();
      toggleCheck(sel.active.r, sel.active.c);
      return;
    }
    if (k === 'F2') {
      e.preventDefault();
      startEdit('edit');
      return;
    }
    if (k === 'Delete') {
      e.preventDefault();
      clear('contents');
      return;
    }
    if (e.altKey && k === 'ArrowDown') {
      e.preventDefault();
      if (pickable) setPicker({ r: sel.active.r, c: sel.active.c });
      else if (activeDv?.type === 'list') setListOpen(true);
      return;
    }
    if (e.altKey && (k === '=' || k === '+')) {
      e.preventDefault();
      autoSum();
      return;
    }
    if (k === 'Backspace') {
      e.preventDefault();
      run('清除内容', (tx) => clearRange(tx, si, { r0: sel.active.r, c0: sel.active.c, r1: sel.active.r, c1: sel.active.c }), ui());
      startEdit('enter', '');
      return;
    }
    if (k === 'Escape') {
      setMenu(null);
      setDetail(null);
      setMarquee(null);
      setPainter(null);
      setFilterMenu(null);
      setSortMenu(null);
      clip = clip?.cut ? null : clip;
      return;
    }
    if (!ctrl) return;
    const low = k.toLowerCase();
    if (low === 'a') {
      e.preventDefault();
      selectAll();
    } else if (low === 'c' || low === 'x') {
      e.preventDefault();
      copySelection(e.shiftKey && low === 'c' ? 'raw' : 'table', sel.range, low === 'x');
    } else if (low === 'h') {
      e.preventDefault();
      props.setFindMode('replace');
    } else if (low === 'p') {
      e.preventDefault();
      setPrintOpen(true);
    } else if (low === 'l' && e.shiftKey) {
      e.preventDefault();
      filterToggle();
    } else if (low === 'b' || low === 'i' || low === 'u' || k === '5') {
      e.preventDefault();
      const key = low === 'b' ? 'bold' : low === 'i' ? 'italic' : low === 'u' ? 'underline' : 'strike';
      styleCmd({ [key]: !activeSt[key] }, '格式');
    } else if (low === 'z') {
      e.preventDefault();
      restore(props.undo());
    } else if (low === 'y') {
      e.preventDefault();
      restore(props.redo());
    } else if (low === 'v' && e.altKey) {
      // Ctrl+Alt+V 选择性粘贴
      e.preventDefault();
      if (clip) setPasteDlg(true);
    } else if (low === 's') {
      e.preventDefault();
      void props.save(e.shiftKey);
    } else if (low === 'd' || low === 'r') {
      // Ctrl+D 向下填充（用选区第一行）、Ctrl+R 向右填充（用第一列）
      e.preventDefault();
      const g = sel.range;
      const src = low === 'd' ? { ...g, r1: g.r0 } : { ...g, c1: g.c0 };
      if ((low === 'd' && g.r1 > g.r0) || (low === 'r' && g.c1 > g.c0)) run('填充', (tx) => fillCopy(tx, src, g), ui());
      else if (low === 'd' && g.r0 > 0) run('填充', (tx) => fillCopy(tx, { ...g, r0: g.r0 - 1, r1: g.r0 - 1 }, g), ui());
      else if (low === 'r' && g.c0 > 0) run('填充', (tx) => fillCopy(tx, { ...g, c0: g.c0 - 1, c1: g.c0 - 1 }, g), ui());
    } else if (k === '1' && !e.shiftKey && !e.altKey) {
      // Ctrl+1：设置单元格格式（同 Excel）
      e.preventDefault();
      setFmtDlg(true);
    } else if (low === 'f') {
      e.preventDefault();
      searchInput.current?.focus();
      searchInput.current?.select();
    } else if (k === '+' || (k === '=' && e.shiftKey)) {
      // Ctrl+Shift+= 插入、Ctrl+- 删除（整行 / 整列选区时插删行列，否则插删行）；复制后虚线框还在时插入复制的单元格
      e.preventDefault();
      if (insClip) insertCopiedCmd(fullCols && !fullRows ? 'right' : 'down');
      else if (fullCols && !fullRows) insertCols();
      else insertRows();
    } else if (k === '-') {
      e.preventDefault();
      if (fullCols && !fullRows) deleteCols();
      else deleteRows();
    } else if (k === '=') {
      e.preventDefault();
      setZoom((z) => stepZoom(z, 1));
    } else if (k === '0') {
      e.preventDefault();
      setZoom(1);
    }
  };
  /** Ctrl+D / Ctrl+R：把源行 / 列原样复制到区域（公式平移） */
  const fillCopy = (tx: Tx, src: Range, to: Range) => {
    const block = copyBlock(book, si, src);
    if (src.r0 === src.r1) for (let r = to.r0; r <= to.r1; r++) r !== src.r0 && pasteBlock(tx, si, r, to.c0, block);
    else for (let c = to.c0; c <= to.c1; c++) c !== src.c0 && pasteBlock(tx, si, to.r0, c, block);
  };

  // ---------- 状态 ----------
  const multi = sel.range.r0 !== sel.range.r1 || sel.range.c0 !== sel.range.c1;
  const merge = merges.get(cellKey(sel.active.r, sel.active.c));
  const singleMerge = !!merge && merge.r0 === sel.range.r0 && merge.c0 === sel.range.c0 && merge.r1 === sel.range.r1 && merge.c1 === sel.range.c1;
  const stats = useMemo(() => {
    if (!multi && !sel.extra?.length) return null;
    // 不连续选区：各块分别统计再合起来
    const parts = allRanges.map((g) => statsOf(sheet, g, hiddenRow, hiddenCol));
    const nums = parts.reduce((a, p) => a + p.nums, 0);
    const sum = Number(parts.reduce((a, p) => a + p.sum, 0).toPrecision(15));
    const withNums = parts.filter((p) => p.nums);
    return {
      count: parts.reduce((a, p) => a + p.count, 0),
      nums,
      sum,
      avg: nums ? Number((sum / nums).toPrecision(15)) : 0,
      min: withNums.length ? Math.min(...withNums.map((p) => p.min)) : 0,
      max: withNums.length ? Math.max(...withNums.map((p) => p.max)) : 0,
    };
  }, [multi, sheet, sel.range, sel.extra, hiddenRow, hiddenCol, version]); // eslint-disable-line react-hooks/exhaustive-deps
  /** 选中了多个格子（不只是一个合并单元格） */
  const rangeSel = multi && !singleMerge;
  const activeRange = merge ?? { r0: sel.active.r, c0: sel.active.c, r1: sel.active.r, c1: sel.active.c };
  const selLabel = rangeSel ? rangeName(sel.range) : cellName(sel.active);
  const hasFreeze = fr > 0 || fc > 0;
  // 冻结选项：行 / 列数按选区算（多选时冻结到选区最后一行 / 列），冻结区太大（超过可视区）就不能用
  const freezeRowsTo = (rangeSel ? sel.range.r1 : activeRange.r1) + 1;
  const freezeColsTo = (rangeSel ? sel.range.c1 : activeRange.c1) + 1;
  const tooTall = (n: number) => rows.pos[Math.min(n, ext.rows)] > vp.h - hh - 40;
  const tooWide = (n: number) => cols.pos[Math.min(n, ext.cols)] > vp.w - hw - 60;
  const freezeOptions: { label: string; rows: number; cols: number; disabled?: boolean }[] = [
    { label: '冻结首行', rows: 1, cols: 0 },
    { label: `冻结前 ${freezeRowsTo} 行（至第 ${freezeRowsTo} 行）`, rows: freezeRowsTo, cols: 0, disabled: freezeRowsTo < 2 || tooTall(freezeRowsTo) },
    { label: '冻结首列', rows: 0, cols: 1 },
    { label: `冻结前 ${freezeColsTo} 列（至 ${colName(freezeColsTo - 1)} 列）`, rows: 0, cols: freezeColsTo, disabled: freezeColsTo < 2 || tooWide(freezeColsTo) },
    {
      label: `冻结窗格（${cellName(sel.active)} 上方的行和左侧的列）`,
      rows: sel.active.r,
      cols: sel.active.c,
      disabled: sel.active.r === 0 || sel.active.c === 0 || tooTall(sel.active.r) || tooWide(sel.active.c),
    },
  ];
  const applyFreeze = (r: number, c: number) => {
    run('冻结窗格', (tx) => setFreeze(tx, si, r, c), ui());
    setMenu(null);
    setFreezeMenu(null);
    focusGrid();
  };
  const freezeItems = () => (
    <>
      {freezeOptions.map((o) => (
        <button key={o.label} role="menuitem" disabled={o.disabled} className={fr === o.rows && fc === o.cols ? 'on' : undefined} onClick={() => applyFreeze(o.rows, o.cols)}>
          {o.label}
        </button>
      ))}
      {hasFreeze && (
        <button role="menuitem" onClick={() => applyFreeze(0, 0)}>
          取消冻结
        </button>
      )}
    </>
  );
  const hitAt = hitIdx >= 0 ? hits[hitIdx] : null;

  // ---------- 渲染一个区域 ----------
  const visible = (a: Axis, from: number, to: number, start: number, end: number): [number, number] => {
    if (end <= start) return [from, from - 1];
    return [Math.max(from, indexAt(a, start)), Math.min(to, indexAt(a, end - 1))];
  };
  const colRangeMain = visible(cols, fc, ext.cols - 1, fw + scroll.x, fw + scroll.x + zoneW);
  const rowRangeMain = visible(rows, fr, ext.rows - 1, fh + scroll.y, fh + scroll.y + zoneH);
  const colRangeFrozen: [number, number] = [0, fc - 1];
  const rowRangeFrozen: [number, number] = [0, fr - 1];

  const renderZone = (key: string, rr: [number, number], cr: [number, number], ox: number, oy: number, style: CSSProperties, frozenRows: boolean, frozenCols: boolean) => {
    const out: ReactNode[] = [];
    for (let r = rr[0]; r <= rr[1]; r++) {
      const h = sizeOf(rows, r);
      if (!h) continue;
      const top = rows.pos[r] - oy;
      const row = sheet.cells[r];
      for (let c = cr[0]; c <= cr[1]; c++) {
        const w = sizeOf(cols, c);
        if (!w || merges.has(cellKey(r, c))) continue;
        out.push(renderCell(r, c, cols.pos[c] - ox, top, w, h, row?.[c]));
      }
    }
    // 合并单元格：左上角属于本区域、且与可视范围相交的
    for (const m of sheet.merges) {
      if (m.r0 < fr !== frozenRows || m.c0 < fc !== frozenCols || m.r1 < rr[0] || m.r0 > rr[1] || m.c1 < cr[0] || m.c0 > cr[1]) continue;
      const w = cols.pos[m.c1 + 1] - cols.pos[m.c0];
      const h = rows.pos[m.r1 + 1] - rows.pos[m.r0];
      if (!w || !h) continue;
      out.push(renderCell(m.r0, m.c0, cols.pos[m.c0] - ox, rows.pos[m.r0] - oy, w, h, sheet.cells[m.r0]?.[m.c0], true));
    }
    // 自动筛选：标题行每格右边一个下拉按钮
    const af = sheet.autoFilter;
    if (af && af.ref.r0 >= rr[0] && af.ref.r0 <= rr[1] && sizeOf(rows, af.ref.r0)) {
      for (let c = Math.max(af.ref.c0, cr[0]); c <= Math.min(af.ref.c1, cr[1]); c++) {
        const w = sizeOf(cols, c);
        if (!w) continue;
        const h = Math.min(sizeOf(rows, af.ref.r0) - 2, Math.round(18 * zoom));
        const on = !!af.cols[c - af.ref.c0];
        out.push(
          <button
            key={`f${c}`}
            className={`sheet-filter-btn${on ? ' on' : ''}`}
            title={on ? '已筛选' : '筛选'}
            style={{ left: cols.pos[c + 1] - ox - h - 2, top: rows.pos[af.ref.r0 + 1] - oy - h - 2, width: h, height: h }}
            onMouseDown={(e) => {
              e.stopPropagation();
              e.preventDefault();
              if (editing) commit(null);
              const b = e.currentTarget.getBoundingClientRect();
              setFilterMenu((m) => (m?.c === c ? null : { c, x: Math.min(b.left, window.innerWidth - 300), y: b.bottom + 2 }));
            }}
          >
            {on ? <FunnelIcon /> : <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true"><path d="M1 2.5l3 3 3-3" fill="currentColor" /></svg>}
          </button>,
        );
      }
    }
    // 表格（套用表格格式）的标题行筛选按钮：只显示，筛选请用“排序和筛选”
    for (const t of sheet.tables ?? []) {
      if (!t.filter || !t.headerRows || (af && af.ref.r0 === t.ref.r0)) continue;
      const r = t.ref.r0;
      if (r < rr[0] || r > rr[1] || !sizeOf(rows, r)) continue;
      for (let c = Math.max(t.ref.c0, cr[0]); c <= Math.min(t.ref.c1, cr[1]); c++) {
        if (!sizeOf(cols, c)) continue;
        const h = Math.min(sizeOf(rows, r) - 2, Math.round(18 * zoom));
        out.push(
          <button
            key={`tf${c}`}
            className="sheet-filter-btn"
            title="表格筛选（这里暂不支持，请选中数据后用“排序和筛选”）"
            style={{ left: cols.pos[c + 1] - ox - h - 2, top: rows.pos[r + 1] - oy - h - 2, width: h, height: h }}
            onMouseDown={(e) => {
              e.stopPropagation();
              e.preventDefault();
              host().toast('表格自带的筛选暂不支持，请选中数据后用“排序和筛选”');
            }}
          >
            <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
              <path d="M1 2.5l3 3 3-3" fill="currentColor" />
            </svg>
          </button>,
        );
      }
    }
    // 边框层：在所有格子之上（后画的相邻格子不会盖住压在网格线上的边框），选区框之下
    const bordered: { r: number; c: number; b: CellBorder; fill?: string }[] = [];
    for (let r = rr[0]; r <= rr[1]; r++) {
      if (!sizeOf(rows, r)) continue;
      for (let c = cr[0]; c <= cr[1]; c++) {
        if (!sizeOf(cols, c)) continue;
        const s = styleAt(sheet, r, c);
        const own = s !== undefined ? book.styles[s] : undefined;
        const tbl = tableStyleAt(sheet.tables, book.styleBook?.palette.theme ?? [], r, c);
        const st = tbl ? overTable(tbl, own) : own;
        const b = cellBorder(st, r, c, merges.get(cellKey(r, c)));
        if (b) bordered.push({ r, c, b, fill: st?.fill });
      }
    }
    // 竖线：右边缘对齐 cols.pos[c]（和网格线一样画成框的右边框）；横线同理对齐 rows.pos[r]
    const lines = borderEdges(bordered).map((e) => (
      <div
        key={`${e.v ? 'v' : 'h'}${cellKey(e.r, e.c)}`}
        className="sheet-border"
        style={
          e.v
            ? { left: cols.pos[e.c] - ox - e.w, top: rows.pos[e.r] - oy - 1, width: e.w, height: sizeOf(rows, e.r) + 1, borderRight: e.css }
            : { left: cols.pos[e.c] - ox - 1, top: rows.pos[e.r] - oy - e.w, width: sizeOf(cols, e.c) + 1, height: e.w, borderBottom: e.css }
        }
      />
    ));
    return (
      <div key={key} className="sheet-zone" style={style}>
        {out}
        {lines}
        {drawingLayer(ox, oy)}
        {selectionBoxes(ox, oy)}
      </div>
    );
  };

  const renderCell = (r: number, c: number, left: number, top: number, width: number, height: number, cell: SCell | undefined, merged = false) => {
    const s = styleAt(sheet, r, c);
    const cfr = cfMap.get(cellKey(r, c));
    const own = s !== undefined ? book.styles[s] : undefined;
    // 表格样式在最下面，单元格自己设的格式盖在它上面，条件格式再盖在最上面
    const tbl = tableStyleAt(sheet.tables, book.styleBook?.palette.theme ?? [], r, c);
    const base = tbl ? overTable(tbl, own) : own;
    const st = cfr?.style ? { ...base, ...cfr.style } : base;
    const cm = bundle ? metaAt(sheet, r, c) : undefined;
    const align = cm?.type === 'check' ? 'center' : st?.align ?? (typeof cell?.v === 'number' ? 'right' : typeof cell?.v === 'boolean' || cell?.k === 'e' ? 'center' : 'left');
    const k = cellKey(r, c);
    let cls = 'sc';
    if (merged) cls += ' merged';
    if (typeof cell?.v === 'number') cls += ' num';
    if (st?.bold) cls += ' bold';
    if (st?.italic) cls += ' italic';
    if (st?.wrap) cls += ' wrap';
    const hit = hitKeys.has(k);
    // 像 Excel：选区内铺浅灰底（有底色时压暗底色），活动单元格保持原底色
    const inSel = !hit && ((rangeSel && contains(sel.range, r, c)) || !!sel.extra?.some((g) => contains(g, r, c))) && !contains(activeRange, r, c);
    const na = !!cell?.f && !!props.engine?.isUnsupported(si, r, c);
    if (na) cls += ' fx-na';
    const note = noteMap.get(k);
    if (note) cls += ' has-note';
    const link = sheet.links ? linkAt(r, c) : undefined;
    if (hit) cls += hitAt && hitAt.r === r && hitAt.c === c ? ' hit on' : ' hit';
    else if (inSel) cls += ' in-sel';
    // 列属性：单选 / 多选画成彩色标签，复选框画方框，不合规的标红角、必填空着的标红边
    let content: ReactNode = cell ? <span>{cell.t}</span> : null;
    if (cm) {
      const text = cell?.t ?? '';
      const opts = optsOf(cm);
      if (cm.type === 'select' && text) {
        content = (
          <span className="ct-tags">
            <span className="ct-tag" style={{ background: optionColor(opts, text) ?? '#94a3b8' }}>
              {text}
            </span>
          </span>
        );
      } else if (cm.type === 'multi' && text) {
        content = (
          <span className="ct-tags">
            {splitMulti(text).map((v) => (
              <span key={v} className="ct-tag" style={{ background: optionColor(opts, v) ?? '#94a3b8' }}>
                {v}
              </span>
            ))}
          </span>
        );
      } else if (cm.type === 'check' && (cell || r < sheet.rows)) {
        const on = isCheckOn(text);
        content = <span className={`ct-check${on ? ' on' : ''}`}>{on ? '✓' : ''}</span>;
      }
      if (text && checkCell(book, sheet, cm, text)) cls += ' col-bad';
      else if (cm.required && !text.trim() && sheet.cells[r]?.some((x) => x && x.t !== '')) cls += ' col-miss';
    }
    const colors = cellColors(st, inSel);
    // 搜索命中的高亮底色优先
    if (hit) delete colors.background;
    const deco = [st?.underline && 'underline', st?.strike && 'line-through'].filter(Boolean).join(' ');
    return (
      <div
        key={merged ? `m${k}` : k}
        className={cls}
        style={{
          left,
          top,
          width,
          height,
          ...(width < 2 * CELL_PAD + 1 && { padding: `0 ${cellPadding(width)}px` }),
          justifyContent: align === 'right' ? 'flex-end' : align === 'center' ? 'center' : 'flex-start',
          textAlign: align,
          alignItems: st?.valign === 'top' ? 'flex-start' : st?.valign === 'center' ? 'center' : st?.valign === 'bottom' ? 'flex-end' : 'center',
          ...(st?.fontSize && st.fontSize !== defFont && { fontSize: Math.round(((13 * st.fontSize) / defFont) * zoom) }),
          ...(st?.fontName && st.fontName !== defFontName && { fontFamily: `"${st.fontName}", var(--font, sans-serif)` }),
          ...(deco && { textDecoration: deco }),
          ...colors,
        }}
        title={
          note
            ? `${note.author ? `${note.author}：\n` : ''}${note.text}`
            : link
              ? `${link.tooltip ?? link.url ?? link.location}\nCtrl+单击打开链接`
              : na
                ? '此公式含暂不支持计算的函数，显示的是文件里保存的结果（保存后 Excel 打开时会重算）'
                : undefined
        }
      >
        {cfr?.bar && <i className="sheet-databar" style={{ width: `${cfr.bar.pct * 100}%`, background: `linear-gradient(90deg, ${cfr.bar.color}, color-mix(in srgb, ${cfr.bar.color} 15%, #fff))` }} />}
        {content}
      </div>
    );
  };

  /** 某行 / 列起点的坐标（超出显示范围的按默认尺寸往后推） */
  const posAt = (a: Axis, i: number, def: number) => (i <= a.n ? a.pos[i] : a.pos[a.n] + (i - a.n) * def * zoom);
  /** 第 i 个图片 / 图表在整张表里的位置（缩放后的像素） */
  const drawingRect = (i: number) => {
    const d = sheet.drawings?.[i];
    if (!d) return null;
    const x = posAt(cols, d.from.c, sheet.defaultColWidth) + d.from.cOff * zoom;
    const y = posAt(rows, d.from.r, sheet.defaultRowHeight) + d.from.rOff * zoom;
    const w = d.to ? posAt(cols, d.to.c, sheet.defaultColWidth) + d.to.cOff * zoom - x : (d.ext?.w ?? 200) * zoom;
    const h = d.to ? posAt(rows, d.to.r, sheet.defaultRowHeight) + d.to.rOff * zoom - y : (d.ext?.h ?? 150) * zoom;
    return { x, y, w, h };
  };
  /** 整张表里的位置（缩放后的像素）→ 锚点（列 / 行 + 格内偏移，未缩放） */
  const anchorAtPx = (x: number, y: number): Anchor => {
    const one = (a: Axis, v: number, def: number) => {
      const i = v < a.pos[a.n] ? indexAt(a, Math.max(0, v)) : a.n + Math.floor((v - a.pos[a.n]) / (def * zoom));
      return { i, off: (v - posAt(a, i, def)) / zoom };
    };
    const cc = one(cols, x, sheet.defaultColWidth);
    const rr = one(rows, y, sheet.defaultRowHeight);
    return { c: cc.i, cOff: cc.off, r: rr.i, rOff: rr.off };
  };
  /** 鼠标下的图片 / 图表（最上面的）：返回序号和相对整张表的坐标 */
  const drawingAtPoint = (clientX: number, clientY: number): { i: number; x: number; y: number } | null => {
    if (!view.current) return null;
    const rect = view.current.getBoundingClientRect();
    const px = clientX - rect.left - hw - fw + fw + scroll.x;
    const py = clientY - rect.top - hh - fh + fh + scroll.y;
    if (clientX - rect.left < hw + fw || clientY - rect.top < hh + fh) return null;
    for (let i = (sheet.drawings?.length ?? 0) - 1; i >= 0; i--) {
      const r = drawingRect(i);
      if (r && px >= r.x && px <= r.x + r.w && py >= r.y && py <= r.y + r.h) return { i, x: px, y: py };
    }
    return null;
  };
  /** 图片和图表：不挡鼠标；选中的那个画选择框和 8 个拖动点（拖动、缩放见 onMouseDown 里的 dragDrawing） */
  const drawingLayer = (ox: number, oy: number) =>
    (sheet.drawings ?? []).map((d, i) => {
      const base = drawingRect(i);
      if (!base) return null;
      const g = drawGhost && dragDrawing.current?.i === i ? { x: drawGhost.x * zoom, y: drawGhost.y * zoom, w: drawGhost.w * zoom, h: drawGhost.h * zoom } : base;
      const { x, y, w, h } = g;
      if (w < 2 || h < 2) return null;
      const style: CSSProperties = { left: x - ox, top: y - oy, width: w, height: h };
      const sel = selDrawing === i;
      return (
        <div key={`d${i}`}>
          {d.kind === 'image' ? (
            <img className="sheet-drawing" src={imageUrls[i]} alt={d.name ?? '图片'} style={style} draggable={false} />
          ) : (
            <div className="sheet-drawing" style={style}>
              {d.chart && <SheetChart spec={d.chart} w={Math.round(w)} h={Math.round(h)} book={book} />}
            </div>
          )}
          {sel && (
            <div className="sheet-drawing-sel" style={style}>
              {['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map((dir) => (
                <i
                  key={dir}
                  className={`dh dh-${dir}`}
                  onMouseDown={(e) => {
                    e.stopPropagation();
                    e.preventDefault();
                    startDrawingDrag(i, dir, e.clientX, e.clientY);
                  }}
                />
              ))}
            </div>
          )}
        </div>
      );
    });
  const startDrawingDrag = (i: number, dir: string, clientX: number, clientY: number) => {
    const r = drawingRect(i);
    if (!r) return;
    dragDrawing.current = { i, dir, sx: clientX, sy: clientY, rect: { x: r.x / zoom, y: r.y / zoom, w: r.w / zoom, h: r.h / zoom } };
    ghostRef.current = dragDrawing.current.rect;
    setDrawGhost(dragDrawing.current.rect);
  };

  // 选区框压在网格线上（向左上多出 1px，盖住上一格的右 / 下网格线），与表头的选区色边线连成一条，不会出现两条线
  const boxOf = (g: Range, ox: number, oy: number): CSSProperties => ({
    left: cols.pos[g.c0] - ox - 1,
    top: rows.pos[g.r0] - oy - 1,
    width: cols.pos[g.c1 + 1] - cols.pos[g.c0] + 1,
    height: rows.pos[g.r1 + 1] - rows.pos[g.r0] + 1,
  });
  const refBoxes = useMemo(() => {
    if (!editing?.text.startsWith('=')) return [];
    const own = sheet.name.toLowerCase();
    return refSpans(editing.text)
      .filter((x) => x.ref.sheet === undefined || x.ref.sheet.toLowerCase() === own)
      .map((x, i) => {
        const a = x.ref.a;
        const b = x.ref.b ?? a;
        const g = {
          r0: Math.min(a.r ?? 0, b.r ?? 0),
          c0: Math.min(a.c ?? 0, b.c ?? 0),
          r1: Math.min(ext.rows - 1, Math.max(a.r ?? ext.rows - 1, b.r ?? ext.rows - 1)),
          c1: Math.min(ext.cols - 1, Math.max(a.c ?? ext.cols - 1, b.c ?? ext.cols - 1)),
        };
        return { g, color: REF_COLORS[i % REF_COLORS.length] };
      });
  }, [editing?.text, sheet.name, ext.rows, ext.cols]); // eslint-disable-line react-hooks/exhaustive-deps
  const selectionBoxes = (ox: number, oy: number) => (
    <>
      {refBoxes.map((x, i) => (
        <div key={`ref${i}`} className="sheet-refbox" style={{ ...boxOf(x.g, ox, oy), borderColor: x.color, background: `color-mix(in srgb, ${x.color} 10%, transparent)` }} />
      ))}
      {sel.extra?.map((g, i) => <div key={`x${i}`} className="sheet-sel-extra" style={boxOf(g, ox, oy)} />)}
      <div className={`${rangeSel ? 'sheet-sel' : 'sheet-active'}${copied ? ' flash' : ''}${editing ? ' editing' : ''}`} key={copied?.n} style={boxOf(rangeSel ? sel.range : activeRange, ox, oy)} />
      {rangeSel && <div className="sheet-active-in" style={boxOf(activeRange, ox, oy)} />}
      {marquee && <div className="sheet-marquee" style={boxOf(marquee.range, ox, oy)} />}
      {fillTo && <div className="sheet-fill-preview" style={boxOf(fillTo, ox, oy)} />}
    </>
  );

  const colHeads = (cr: [number, number], ox: number, style: CSSProperties, key: string) => {
    const out: ReactNode[] = [];
    for (let c = cr[0]; c <= cr[1]; c++) {
      const w = sizeOf(cols, c);
      if (!w) continue;
      const on = c >= sel.range.c0 && c <= sel.range.c1;
      const cm = bundle ? columnAt(sheet.columns, c) : undefined;
      out.push(
        <div key={c} className={`sh col${on ? ' on' : ''}${on && fullCols ? ' full' : ''}`} style={{ left: cols.pos[c] - ox, width: w, height: hh0 }}>
          {cm && (
            <span className="ct-glyph" title={`${COLUMN_TYPE_LABELS[cm.type]}列`}>
              {COLUMN_TYPE_GLYPHS[cm.type]}
            </span>
          )}
          {colName(c)}
          {bundle && w >= 40 && (
            <button
              className="ct-btn"
              aria-label="列属性"
              title="列属性"
              onMouseDown={(e) => {
                e.stopPropagation();
                e.preventDefault();
                const b = e.currentTarget.getBoundingClientRect();
                openColumnPanel(c, { x: b.left - 340, y: b.bottom + 6 });
              }}
            >
              ▾
            </button>
          )}
        </div>,
      );
    }
    return (
      <div key={key} className="sheet-zone head" style={style}>
        {out}
      </div>
    );
  };
  const rowHeads = (rr: [number, number], oy: number, style: CSSProperties, key: string) => {
    const out: ReactNode[] = [];
    for (let r = rr[0]; r <= rr[1]; r++) {
      const h = sizeOf(rows, r);
      if (!h) continue;
      const on = r >= sel.range.r0 && r <= sel.range.r1;
      out.push(
        <div key={r} className={`sh row${on ? ' on' : ''}${on && fullRows ? ' full' : ''}`} style={{ top: rows.pos[r] - oy, height: h, width: hw0 }}>
          {r + 1}
        </div>,
      );
    }
    return (
      <div key={key} className="sheet-zone head" style={style}>
        {out}
      </div>
    );
  };

  /** 分组栏：每行 / 列画它所在的各级分组竖线（横线），汇总行 / 列上画折叠按钮 */
  const outlineZone = (axis: OutlineAxis, range: [number, number], off: number, style: CSSProperties, key: string) => {
    const pos = axis === 'row' ? rows : cols;
    const out: ReactNode[] = [];
    const after = summaryAfter(sheet, axis);
    for (let i = range[0]; i <= range[1]; i++) {
      const size = sizeOf(pos, i);
      if (!size) continue;
      const lv = levelAt(sheet, axis, i);
      const at = pos.pos[i] - off;
      const box = axis === 'row' ? { top: at, height: size } : { left: at, width: size };
      for (let l = 1; l <= lv; l++) {
        const line = (l - 1) * OG + Math.floor(OG / 2);
        out.push(<i key={`${i}b${l}`} className={`og-bar ${axis}`} style={axis === 'row' ? { ...box, left: line } : { ...box, top: line }} />);
      }
      const neighbour = levelAt(sheet, axis, after ? i - 1 : i + 1);
      if (neighbour > lv) {
        const folded = collapsedAt(sheet, axis, i);
        const slot = lv * OG + 1;
        const mid = Math.max(0, Math.floor((size - (OG - 2)) / 2));
        out.push(
          <button
            key={`${i}t`}
            className="og-btn"
            aria-label={folded ? '展开分组' : '折叠分组'}
            title={folded ? '展开' : '折叠'}
            style={axis === 'row' ? { top: at + Math.min(mid, 4), left: slot } : { left: at + Math.min(mid, 4), top: slot }}
            onMouseDown={(e) => (e.stopPropagation(), e.preventDefault())}
            onClick={(e) => {
              e.stopPropagation();
              run(folded ? '展开分组' : '折叠分组', (tx) => toggleGroup(tx, si, axis, i), ui());
              focusGrid();
            }}
          >
            {folded ? '+' : '−'}
          </button>,
        );
      }
    }
    return (
      <div key={key} className="sheet-zone og" style={style} onMouseDown={(e) => e.stopPropagation()} onMouseMove={(e) => e.stopPropagation()} onDoubleClick={(e) => e.stopPropagation()} onContextMenu={(e) => e.stopPropagation()}>
        {out}
      </div>
    );
  };
  /** 左上角的级别按钮：1 = 只看没分组的，2 = 再多看一级 … */
  const levelButtons = (axis: OutlineAxis) => {
    const top = axis === 'row' ? olRow : olCol;
    if (!top) return null;
    const x0 = axis === 'row' ? 0 : olRow ? (olRow + 1) * OG : 0;
    const y0 = axis === 'row' ? (olCol ? (olCol + 1) * OG : 0) : 0;
    return Array.from({ length: top + 1 }, (_, k) => (
      <button
        key={`${axis}${k}`}
        className="og-btn lv"
        aria-label={`显示${axis === 'row' ? '行' : '列'}分组的第 ${k + 1} 级`}
        title={k === top ? '显示全部' : `只显示前 ${k + 1} 级`}
        style={axis === 'row' ? { left: x0 + k * OG + 1, top: y0 + 1 } : { left: x0 + 1, top: y0 + k * OG + 1 }}
        onMouseDown={(e) => (e.stopPropagation(), e.preventDefault())}
        onClick={(e) => {
          e.stopPropagation();
          run('分组级别', (tx) => showLevel(tx, si, axis, k + 1), ui());
          focusGrid();
        }}
      >
        {k + 1}
      </button>
    ));
  };
  /** 组合 / 取消组合选中的行（整列选中时是列）；Alt+Shift+→ / ← */
  const groupSelection = (group: boolean) => {
    const g = sel.range;
    const axis: OutlineAxis = fullCols && !fullRows ? 'col' : 'row';
    const from = axis === 'row' ? g.r0 : g.c0;
    const to = axis === 'row' ? Math.min(g.r1, Math.max(g.r0, ext.rows - 1)) : Math.min(g.c1, Math.max(g.c0, ext.cols - 1));
    if (group && maxLevel(sheet, axis) >= 7 && levelAt(sheet, axis, from) >= 7) return host().toast('分组最多 8 层');
    run(group ? '组合' : '取消组合', (tx) => (group ? groupRange(tx, si, axis, from, to) : ungroupRange(tx, si, axis, from, to)), ui());
  };

  const ox = fw + scroll.x;
  const oy = fh + scroll.y;
  const zones = [
    renderZone('main', rowRangeMain, colRangeMain, ox, oy, { left: hw + fw, top: hh + fh, width: zoneW, height: zoneH }, false, false),
    fr > 0 && renderZone('top', rowRangeFrozen, colRangeMain, ox, 0, { left: hw + fw, top: hh, width: zoneW, height: fh }, true, false),
    fc > 0 && renderZone('left', rowRangeMain, colRangeFrozen, 0, oy, { left: hw, top: hh + fh, width: fw, height: zoneH }, false, true),
    fr > 0 && fc > 0 && renderZone('corner', rowRangeFrozen, colRangeFrozen, 0, 0, { left: hw, top: hh, width: fw, height: fh }, true, true),
    colHeads(colRangeMain, ox, { left: hw + fw, top: gh, width: zoneW, height: hh0 }, 'ch'),
    fc > 0 && colHeads(colRangeFrozen, 0, { left: hw, top: gh, width: fw, height: hh0 }, 'chf'),
    rowHeads(rowRangeMain, oy, { left: gw, top: hh + fh, width: hw0, height: zoneH }, 'rh'),
    fr > 0 && rowHeads(rowRangeFrozen, 0, { left: gw, top: hh, width: hw0, height: fh }, 'rhf'),
    ...(gw > 0
      ? [
          outlineZone('row', rowRangeMain, oy, { left: 0, top: hh + fh, width: gw, height: zoneH }, 'orw'),
          fr > 0 && outlineZone('row', rowRangeFrozen, 0, { left: 0, top: hh, width: gw, height: fh }, 'orwf'),
        ]
      : []),
    ...(gh > 0
      ? [
          outlineZone('col', colRangeMain, ox, { left: hw + fw, top: 0, width: zoneW, height: gh }, 'ocl'),
          fc > 0 && outlineZone('col', colRangeFrozen, 0, { left: hw, top: 0, width: fw, height: gh }, 'oclf'),
        ]
      : []),
  ];

  const openChart = () => {
    const d = dataRegion();
    const g = rangeSel ? sel.range : d?.g;
    if (!g || (g.r1 === g.r0 && g.c1 === g.c0)) return host().toast('请先选中要画成图表的数据');
    setChartDlg({ range: g });
  };
  const openNote = (r: number, c: number) => {
    const n = noteMap.get(cellKey(r, c));
    if (n && isThreaded(n)) return host().toast('这是 Excel 365 的对话式批注，请在 Excel 里编辑');
    setNoteDlg({ r, c });
  };
  const gotoRange = (toSi: number, g: Range) => {
    if (toSi !== si) props.onSheet(toSi);
    const a = { r: g.r0, c: g.c0 };
    if (g.r1 === a.r && g.c1 === a.c) selectCell(a);
    else extendTo(a, a, { r: g.r1, c: g.c1 });
    focusGrid();
  };
  /** 当前选区写成 Sheet1!$A$1:$B$5（名称的引用位置） */
  const selectionRef = () => {
    const g = rangeSel ? sel.range : { ...sel.range, r1: sel.range.r0, c1: sel.range.c0 };
    const a = `$${colName(g.c0)}$${g.r0 + 1}`;
    return `${quoteSheet(sheet.name)}!${g.r1 === g.r0 && g.c1 === g.c0 ? a : `${a}:$${colName(g.c1)}$${g.r1 + 1}`}`;
  };
  const commitAddr = () => {
    const g = addr === null ? null : parseRef(addr);
    const typed = addr;
    setAddr(null);
    if (!g && typed) {
      // 不是地址：当成名称，指向单个区域就跳过去
      const hit = nameTarget(book, typed.trim(), si);
      if (hit?.range) return gotoRange(hit.si, hit.range);
      if (typed.trim()) host().toast(`找不到“${typed.trim()}”`);
    }
    if (!g) return focusGrid();
    const a = { r: g.r0, c: g.c0 };
    if (g.r1 === a.r && g.c1 === a.c) selectCell(a);
    else extendTo(a, a, { r: g.r1, c: g.c1 });
    focusGrid();
  };

  // ---------- 单元格编辑框 ----------
  const edR = editing ? (merges.get(cellKey(editing.r, editing.c)) ?? { r0: editing.r, c0: editing.c, r1: editing.r, c1: editing.c }) : activeRange;
  const edStyle = (() => {
    const s = styleAt(sheet, edR.r0, edR.c0);
    const st = s !== undefined ? book.styles[s] : undefined;
    const left = Math.max(hw, xOf(edR.c0));
    const top = Math.max(hh, yOf(edR.r0));
    const css: CSSProperties = {
      left,
      top,
      minWidth: Math.max(8, xOf(edR.c1 + 1) - left),
      minHeight: Math.max(8, yOf(edR.r1 + 1) - top),
      maxWidth: Math.max(60, vp.w - left - 4),
      maxHeight: Math.max(40, vp.h - top - 4),
    };
    if (st?.bold) css.fontWeight = 600;
    if (st?.italic) css.fontStyle = 'italic';
    if (st?.fontSize && st.fontSize !== defFont) css.fontSize = Math.round(((13 * st.fontSize) / defFont) * zoom);
    if (st?.fill) {
      css.background = st.fill;
      css.color = st.color ?? '#000000';
    } else if (st?.color && st.color !== '#000000') css.color = st.color;
    return css;
  })();
  const onKbChange = (text: string) => {
    pointRef.current = null;
    if (editing) setEditing({ ...editing, text });
    else {
      // 直接打字（或输入法上屏）：覆盖原内容进入编辑
      const p = sel.active;
      const orig = rawText(p);
      setEditing({ r: p.r, c: p.c, text, orig: `\u0000${orig}`, mode: 'enter' });
      setMarquee(null);
      reveal(p);
    }
  };

  const detailCell = detail ? cellAt(sheet, detail) : undefined;
  const barText = editing ? editing.text : rawText(sel.active);
  /** 编辑栏输入框的属性：桌面的 fx 栏和手机底部的编辑栏共用（同一套编辑 / 提交 / 点选引用逻辑） */
  const barProps = {
    ref: barInput,
    value: barText,
    onFocus: () => {
      if (!editing) startEdit('edit', undefined, true);
      else if (!editing.bar) setEditing({ ...editing, bar: true, mode: 'edit' });
    },
    onChange: (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      pointRef.current = null;
      setEditing((ed) => (ed ? { ...ed, text: e.target.value } : ed));
      setCaretAt(e.target.selectionStart);
      setHintIdx(0);
    },
    onKeyUp: (e: React.KeyboardEvent<HTMLTextAreaElement>) => trackCaret(e.currentTarget),
    onClick: (e: React.MouseEvent<HTMLTextAreaElement>) => trackCaret(e.currentTarget),
    onKeyDown: (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (hintKey(e)) return;
      if (e.key === 'Enter' && !e.altKey) {
        e.preventDefault();
        commit(e.ctrlKey ? null : [e.shiftKey ? -1 : 1, 0], e.ctrlKey);
      } else if (e.key === 'Enter' && e.altKey) {
        e.preventDefault();
        const el = e.currentTarget;
        const a = el.selectionStart;
        setEditing((ed) => (ed ? { ...ed, text: ed.text.slice(0, a) + '\n' + ed.text.slice(el.selectionEnd) } : ed));
        requestAnimationFrame(() => el.setSelectionRange(a + 1, a + 1));
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commit([0, e.shiftKey ? -1 : 1]);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
      } else if (e.key === 'F4') {
        e.preventDefault();
        toggleAbsolute(e.currentTarget);
      }
    },
  };
  const notes = noteHidden ? [] : props.notes;

  return (
    <div className={`sheet${enter ? ' enter' : ''}${mobile ? ' mobile' : ''}`} style={{ '--sheet-font': `${Math.round(13 * zoom)}px` } as CSSProperties}>
      {!mobile && (<>
      <div className="ed-tabbar">
        <div className="ed-tabs" role="tablist">
          {(['home', 'page'] as const).map((t) => (
            <button key={t} role="tab" aria-selected={barTab === t} className={barTab === t ? 'on' : ''} onClick={() => setBarTab(t)}>
              {t === 'home' ? '常用' : '页面'}
            </button>
          ))}
        </div>
        <div className="ed-quick ed-quick-tab">
          {host().newDocument && (
            <button className="icon-btn ed-raised" title="新建 Excel 工作簿" aria-label="新建" onClick={() => host().newDocument?.('sheet')}>
              <Icon name="newDocColor" size={20} />
            </button>
          )}
          {host().openDocument && (
            <button className="icon-btn ed-raised" title="打开文件…" aria-label="打开" onClick={() => void host().openDocument?.('sheet')}>
              <Icon name="openColor" size={20} />
            </button>
          )}
          <div className="ed-split" ref={saveMenuRef}>
            <button className={`icon-btn ed-raised${props.dirty ? ' dirty' : ''}`} title="保存（Ctrl+S）" aria-label="保存" onClick={() => void props.save(false)}>
              <Icon name="saveColor" size={20} />
            </button>
            <button className="icon-btn ed-raised ed-split-arrow" aria-label="保存选项" aria-haspopup="menu" aria-expanded={saveMenu} onClick={() => setSaveMenu(!saveMenu)}>
              <Icon name="chevron" size={16} className="ed-split-chev" />
            </button>
            {saveMenu && (
              <div className="ed-split-menu" role="menu">
                <button role="menuitem" onClick={() => (setSaveMenu(false), void props.save(false))}>保存</button>
                <button role="menuitem" title="另存为（Ctrl+Shift+S）" onClick={() => (setSaveMenu(false), void props.save(true))}>另存为</button>
                <button role="menuitem" title="打开打印设置，在里面导出 PDF" onClick={() => (setSaveMenu(false), setPrintOpen(true))}>另存为 PDF</button>
              </div>
            )}
          </div>
        <button className="icon-btn ed-raised" title="撤销（Ctrl+Z）" aria-label="撤销" disabled={!props.canUndo} onClick={() => (restore(props.undo()), focusGrid())}>
          <Icon name="undo" size={15} />
        </button>
        <button className="icon-btn ed-raised" title="重做（Ctrl+Y）" aria-label="重做" disabled={!props.canRedo} onClick={() => (restore(props.redo()), focusGrid())}>
          <Icon name="redo" size={15} />
        </button>
        <button
          className={`icon-btn ed-raised${painter ? ' on' : ''}`}
          aria-label="格式刷"
          aria-pressed={!!painter}
          title="格式刷：单击刷一次，双击可以连续刷（Esc 结束）"
          onClick={(e) => (painterCmd(e.detail === 2), focusGrid())}
        >
          <Icon name="brush" size={20} />
        </button>
        <button className="icon-btn ed-raised" aria-label="清除格式" title="清除格式" onClick={() => (clear('formats'), focusGrid())}>
          <Icon name="eraser" size={16} />
        </button>
        <button className="icon-btn ed-raised" title="打印 / 导出 PDF（Ctrl+P）" aria-label="打印" onClick={() => setPrintOpen(true)}>
          <Icon name="printer" size={15} />
        </button>
        </div>
        <span className="grow" />
        <div className="ed-tabbar-right">
        <div className={`sheet-search${query ? ' has' : ''}`}>
          <Icon name="search" size={13} />
          <input
            ref={searchInput}
            placeholder="搜索 Ctrl+F"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                gotoHit(e.shiftKey ? -1 : 1);
              }
              if (e.key === 'Escape') {
                setQuery('');
                focusGrid();
              }
            }}
          />
          {query && <span className="mono sheet-search-count">{hits.length ? `${hitIdx + 1 || '–'} / ${hits.length}` : '0'}</span>}
        </div>
        <button className={`btn${props.findMode ? ' soft on' : ''}`} title="查找和替换（Ctrl+H）" onClick={() => props.setFindMode(props.findMode ? null : 'replace')}>
          查找替换
        </button>
        {props.hasPath && (
          <>
            <button className="btn sheet-excel-btn" title="用系统默认程序（Excel / WPS）打开" onClick={onExternal}>
              <ExcelIcon />
              用 Excel 打开
            </button>
          </>
        )}
        <FloatViewButton tabId={props.tabId} path={props.path} className="icon-btn ed-raised" />
        </div>
      </div>
      {barTab === 'page' && (
        <div className="toolbar sheet-bar">
        <button className={`btn${sheet.cf?.length ? ' soft on' : ''}`} title="条件格式：按条件给单元格着色、加数据条和色阶" onClick={() => (setSortMenu(null), setMenu(null), setFreezeMenu(null), setCfOpen(true))}>
          条件格式
        </button>
        <button
          className={`btn${hasFreeze ? ' soft on' : ''}`}
          title="冻结行 / 列：先选中要冻结到的行、列或单元格"
          aria-haspopup="menu"
          onClick={(e) => {
            // 不让这次点击冒泡到 window：菜单打开后立刻注册的“点击关闭”监听会收到同一次点击，菜单一闪就没了
            e.stopPropagation();
            const b = e.currentTarget.getBoundingClientRect();
            setMenu(null);
            setFreezeMenu((m) => (m ? null : { x: b.left, y: b.bottom + 4 }));
          }}
        >
          <Icon name="pin" size={13} />
          {hasFreeze ? (fc === 0 ? `已冻结 ${fr} 行` : fr === 0 ? `已冻结 ${fc} 列` : `已冻结 ${fr} 行 ${fc} 列`) : '冻结'}
          <Icon name="chevron" size={15} className="sheet-caret" />
        </button>
            <span className="grow" />
        </div>
      )}

      {barTab === 'home' && (
      <FormatBar
        extra={
          <>
        {props.onAssistant && (
          <button className="btn" title="AI 助手：用对话让 AI 阅读、整理和计算这张表（每次修改都可以撤销）" onClick={props.onAssistant}>
            AI 助手
          </button>
        )}
        <button className="btn" title="自动求和（Alt+=）：单个格子求上方或左边的数字；选中一片数字时在每列下面写合计" onClick={() => (autoSum(), !editing && focusGrid())}>
          Σ 求和
        </button>
        <button
          className={`btn${sheet.autoFilter ? ' soft on' : ''}`}
          title="排序和筛选"
          aria-haspopup="menu"
          onClick={(e) => {
            e.stopPropagation();
            const b = e.currentTarget.getBoundingClientRect();
            setMenu(null);
            setFreezeMenu(null);
            setSortMenu((m) => (m ? null : { x: b.left, y: b.bottom + 4 }));
          }}
        >
          <FunnelIcon />
          排序和筛选
          <Icon name="chevron" size={15} className="sheet-caret" />
        </button>
        <button className="btn" title="用选中的数据插入图表（柱形 / 条形 / 折线 / 面积 / 饼图）" onClick={() => openChart()}>
          插入图表
        </button>
        {bundle && (
          <button className={`btn${formPanel ? ' soft on' : ''}`} title="表单录入：按列属性逐条录入新记录（也可以改当前行）" onClick={() => (formPanel ? setFormPanel(null) : openForm())}>
            表单录入
          </button>
        )}
          </>
        }
        st={activeSt}
        fmt={activeFmt}
        defaultFont={defFontName}
        defaultSize={defFont}
        merged={!!merge}
        onStyle={styleCmd}
        onBorder={borderCmd}
        onMerge={mergeCmd}
        onNumFmt={numFmtCmd}
        onCustomFmt={() => void customFmtCmd()}
        onDecimals={decimalsCmd}
        onDone={focusGrid}
      />
      )}
      </>)}

      {!mobile && (
      <div className="sheet-fbar">
        <div className="sheet-addr mono" title="输入地址后回车跳转，如 D120 或 A1:C20">
          <input
            aria-label="单元格地址"
            value={addr ?? selLabel}
            onFocus={(e) => {
              if (editing) commit(null);
              setAddr(selLabel);
              e.currentTarget.select();
            }}
            onChange={(e) => setAddr(e.target.value)}
            onBlur={() => setAddr(null)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                commitAddr();
              }
              if (e.key === 'Escape') {
                setAddr(null);
                focusGrid();
              }
            }}
          />
          <button
            className="icon-btn sheet-addr-btn"
            aria-label="名称列表"
            title="定义的名称"
            onMouseDown={(e) => e.preventDefault()}
            onClick={(e) => {
              const r = e.currentTarget.getBoundingClientRect();
              setNameList({ x: r.left - 150, y: r.bottom + 4 });
            }}
          >
            <Icon name="chevron" size={12} />
          </button>
        </div>
        {sheet.protected && (
          <span className="sheet-lock" title="这张工作表受保护：只能修改未锁定的单元格（右键 → 撤销工作表保护）">
            已保护
          </span>
        )}
        <span className="sheet-fx mono">fx</span>
        <textarea className="sheet-formula grow" aria-label="编辑栏" rows={1} spellCheck={false} {...barProps} />
      </div>
      )}

      {(notes.length > 0 || props.external) && (
        <div className="sheet-note" role="status">
          {props.external ? (
            <>
              <span className="grow">文件已被其它程序修改。</span>
              <button className="btn ghost" onClick={onReload}>
                重新读取（放弃我的修改）
              </button>
              <button className="btn ghost" onClick={props.onKeepMine}>
                保留我的修改
              </button>
            </>
          ) : (
            <>
              <span className="grow">{notes.join('；')}</span>
              <button className="btn ghost" onClick={() => setNoteHidden(true)}>
                知道了
              </button>
            </>
          )}
        </div>
      )}

      <div className="sheet-body">
        <div className="sheet-scroll" ref={scroller} onScroll={onScroll}>
          <div ref={gridContent} style={{ width: hw + cols.pos[ext.cols], height: hh + rows.pos[ext.rows], position: 'relative' }}>
            <div
              ref={view}
              className={`sheet-view${edgeHover ? ` resize-${edgeHover}` : ''}${painter ? ' painting' : ''}`}
              style={{ width: vp.w, height: vp.h }}
              onMouseDown={onMouseDown}
              onMouseMove={(e) => {
                if (resize.current || drag.current) return;
                const edge = onHandle(e.clientX, e.clientY) ? 'fill' : (edgeTest(e.clientX, e.clientY)?.axis ?? null);
                if (edge !== edgeHover) setEdgeHover(edge);
              }}
              onMouseLeave={() => !resize.current && setEdgeHover(null)}
              onDoubleClick={(e) => {
                if (editing || mobile) return;
                const dh = drawingAtPoint(e.clientX, e.clientY);
                if (dh) {
                  if (sheet.drawings?.[dh.i]?.kind === 'chart') setChartDlg({ index: dh.i });
                  return;
                }
                if (onHandle(e.clientX, e.clientY)) return autoFillDown();
                // 双击分隔线：列宽按内容调整，行高恢复自动
                const edge = edgeTest(e.clientX, e.clientY);
                if (edge?.axis === 'col') return autoFit(edge.i);
                if (edge?.axis === 'row') return void run('行高', (tx) => setRowHeight(tx, si, edge.i, edge.i, { height: null }), ui());
                const h = hitTest(e.clientX, e.clientY);
                if (h.area === 'cell') startEdit('edit');
                else if (h.area === 'col') autoFit(h.c);
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                if (mobile) return;
                setFreezeMenu(null);
                const h = hitTest(e.clientX, e.clientY);
                setMenuCol(h.area === 'cell' || h.area === 'col' ? h.c : sel.active.c);
                setMenu({ x: e.clientX, y: e.clientY });
              }}
            >
              {zones}
              <div className="sheet-corner" style={{ width: hw, height: hh }} title="全选" />
              {(gw > 0 || gh > 0) && (
                <div className="og-corner" style={{ width: hw, height: hh }} onMouseDown={(e) => e.stopPropagation()} onMouseMove={(e) => e.stopPropagation()}>
                  {levelButtons('row')}
                  {levelButtons('col')}
                </div>
              )}
              {fr > 0 && <div className="sheet-freeze-h" style={{ top: hh + fh - 1, left: 0, width: vp.w }} />}
              {fc > 0 && <div className="sheet-freeze-v" style={{ left: hw + fw - 1, top: 0, height: vp.h }} />}
              <textarea
                ref={kb}
                className={`sheet-editor${editing && !editing.bar ? ' on' : editing ? ' on mirror' : ''}`}
                aria-label="单元格内容"
                inputMode={mobile ? 'none' : undefined}
                spellCheck={false}
                style={editing ? edStyle : { left: edStyle.left, top: edStyle.top }}
                value={editing ? editing.text : ''}
                readOnly={!!editing?.bar}
                onChange={(e) => {
                  onKbChange(e.target.value);
                  setCaretAt(e.target.selectionStart);
                  setHintIdx(0);
                }}
                onKeyDown={onKeyDown}
                onKeyUp={(e) => trackCaret(e.currentTarget)}
                onClick={(e) => trackCaret(e.currentTarget)}
                onPaste={(e) => {
                  if (editing) return;
                  e.preventDefault();
                  pasteText(e.clipboardData.getData('text/plain'), e.clipboardData.getData('text/html') || undefined);
                }}
              />
              {!editing && pickable && (
                <button
                  className="sheet-dv-btn"
                  aria-label={activeMeta?.type === 'date' ? '选择日期' : '从选项中选择'}
                  style={{ left: xOf(activeRange.c1 + 1) + 1, top: yOf(activeRange.r1 + 1) - Math.round(20 * zoom), height: Math.round(19 * zoom), width: Math.round(17 * zoom) }}
                  onMouseDown={(e) => {
                    e.stopPropagation();
                    e.preventDefault();
                    setPicker((p) => (p ? null : { r: sel.active.r, c: sel.active.c }));
                  }}
                >
                  <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
                    <path d="M1 2.5l3 3 3-3" fill="currentColor" />
                  </svg>
                </button>
              )}
              {!editing && picker && pickerMeta && (
                <ColumnPicker
                  meta={pickerMeta}
                  options={optsOf(pickerMeta)}
                  value={colCellText(sheet, picker.r, picker.c)}
                  at={{ left: xOf(picker.c), top: yOf(picker.r + 1) + 2 }}
                  onPick={(text) => pickValue(picker.r, picker.c, text)}
                  onCreate={pickerMeta.optionsRef ? undefined : (text) => addOption(pickerMeta, text)}
                  onManage={() => {
                    const c = picker.c;
                    setPicker(null);
                    openColumnPanel(c, { x: Math.max(8, xOf(c)), y: yOf(picker.r + 1) + 6 });
                  }}
                  onClose={() => {
                    setPicker(null);
                    focusGrid();
                  }}
                />
              )}
              {!editing && !pickable && activeDv?.type === 'list' && !activeDv.noDropDown && (
                <button
                  className="sheet-dv-btn"
                  aria-label="从列表中选择"
                  style={{ left: xOf(activeRange.c1 + 1) + 1, top: yOf(activeRange.r1 + 1) - Math.round(20 * zoom), height: Math.round(19 * zoom), width: Math.round(17 * zoom) }}
                  onMouseDown={(e) => {
                    e.stopPropagation();
                    e.preventDefault();
                    setListOpen((o) => !o);
                  }}
                >
                  <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
                    <path d="M1 2.5l3 3 3-3" fill="currentColor" />
                  </svg>
                </button>
              )}
              {!editing && listOpen && activeDv && (
                <div className="sheet-hint sheet-dv-list" style={{ left: xOf(activeRange.c0), top: yOf(activeRange.r1 + 1) + 2 }} onMouseDown={(e) => (e.stopPropagation(), e.preventDefault())}>
                  {listItems(activeDv, book, si).map((x) => (
                    <button
                      key={x}
                      className={sheet.cells[sel.active.r]?.[sel.active.c]?.t === x ? 'on' : ''}
                      onClick={() => {
                        setListOpen(false);
                        run('输入', (tx) => setInput(tx, si, sel.active.r, sel.active.c, /^[\d.,%-]+$|^(true|false)$/i.test(x) ? x : `'${x}`), ui());
                        focusGrid();
                      }}
                    >
                      {x}
                    </button>
                  ))}
                </div>
              )}
              {!editing && !listOpen && activeDv?.prompt && (
                <div className="sheet-dv-prompt" style={{ left: xOf(activeRange.c0) + 8, top: yOf(activeRange.r1 + 1) + 4 }}>
                  {activeDv.promptTitle && <strong>{activeDv.promptTitle}</strong>}
                  <span>{activeDv.prompt}</span>
                </div>
              )}
              {editing && (hints.length > 0 || call) && (
                <div className="sheet-hint" style={{ left: Math.max(hw, xOf(edR.c0)), top: Math.max(hh, yOf(edR.r1 + 1)) + 4 }} onMouseDown={(e) => e.preventDefault()}>
                  {hints.length > 0
                    ? hints.map((n, i) => (
                        <button key={n} className={i === Math.min(hintIdx, hints.length - 1) ? 'on' : ''} onClick={() => acceptHint(n, editing.bar ? barInput.current : kb.current)}>
                          <span className="mono">{n}</span>
                          <span className="sub ellipsis">{FUNCTION_INFO[n].desc}</span>
                        </button>
                      ))
                    : call && (
                        <div className="sheet-hint-sig">
                          <span className="mono">
                            {call.name}(
                            {FUNCTION_INFO[call.name].args.map((a, i, all) => (
                              <span key={i}>
                                <span className={i === Math.min(call.arg, all.length - 1) || (all[i] === '…' && call.arg >= i) ? 'cur' : ''}>{a}</span>
                                {i < all.length - 1 ? ', ' : ''}
                              </span>
                            ))}
                            )
                          </span>
                          <span className="sub">{FUNCTION_INFO[call.name].desc}</span>
                        </div>
                      )}
                </div>
              )}
              {resizeTip && (
                <div className="sheet-resize-tip mono" style={{ left: resizeTip.x + 12, top: resizeTip.y + 14 }}>
                  {resizeTip.text}
                </div>
              )}
            </div>
          </div>
        </div>
        {copied && (
          <div className="sheet-copied mono" key={copied.n}>
            <Icon name="check" size={13} />
            {copied.text}
          </div>
        )}
      </div>

      <div className="sheet-foot">
        <SheetTabs
          sheets={book.sheets}
          current={sheetIndex}
          onSelect={(i) => {
            if (editing) commit(null);
            if (i !== si) onSheet(i);
          }}
          onAdd={() => bookCmd('插入工作表', (tx) => addSheet(tx, si + 1), (i) => i as number)}
          onRename={sheetName}
          onDelete={async (i) => {
            const s = book.sheets[i];
            if (s.rows && !(await host().confirm(`删除工作表「${s.name}」？引用它的公式会变成 #REF!（可以撤销）。`))) return;
            bookCmd(
              '删除工作表',
              (tx) => {
                if (!deleteSheet(tx, i)) host().toast('工作簿里至少要有一张可见的工作表');
              },
              () => Math.max(0, book.sheets.findIndex((x, j) => !x.hidden && j >= Math.min(i, book.sheets.length - 1))),
            );
          }}
          onCopy={(i) => bookCmd('复制工作表', (tx) => copySheet(tx, i), (n) => n as number)}
          onMove={(from, to) => bookCmd('移动工作表', (tx) => moveSheet(tx, from, to), () => to)}
          onHide={(i) =>
            bookCmd(
              '隐藏工作表',
              (tx) => {
                if (!setSheetHidden(tx, i, true)) host().toast('工作簿里至少要有一张可见的工作表');
              },
              () => book.sheets.findIndex((x) => !x.hidden),
            )
          }
          onUnhide={(i) => bookCmd('取消隐藏工作表', (tx) => setSheetHidden(tx, i, false), () => i)}
          onTabColor={(i, c) => bookCmd('标签颜色', (tx) => setTabColor(tx, i, c))}
          onDone={focusGrid}
        />
        <div className="sheet-hud mono" style={mobile ? { display: 'none' } : undefined}>
          {missing.length > 0 && (
            <button className="sheet-col-hud" title="点击跳到第一个" onClick={() => (selectCell({ r: missing[0].r, c: missing[0].c }), focusGrid())}>
              {missing.length} 个必填单元格为空
            </button>
          )}
          {stats ? (
            <>
              <HudItem label="计数" value={String(stats.count)} onCopy={copyText} />
              {stats.nums > 0 && (
                <>
                  <HudItem label="求和" value={fmtStat(stats.sum)} raw={String(stats.sum)} onCopy={copyText} />
                  <HudItem label="平均" value={fmtStat(stats.avg)} raw={String(stats.avg)} onCopy={copyText} />
                  <HudItem label="最小" value={fmtStat(stats.min)} raw={String(stats.min)} onCopy={copyText} />
                  <HudItem label="最大" value={fmtStat(stats.max)} raw={String(stats.max)} onCopy={copyText} />
                </>
              )}
            </>
          ) : (
            <span className="sub">
              {sheet.rows} 行 × {sheet.cols} 列
            </span>
          )}
        </div>
        <div className="sheet-zoom" style={mobile ? { display: 'none' } : undefined}>
          <button className="icon-btn" aria-label="缩小" onClick={() => setZoom((z) => stepZoom(z, -1))}>
            <Icon name="zoomOut" size={14} />
          </button>
          <button className="mono sheet-zoom-val" title="恢复 100%（Ctrl+0）" onClick={() => setZoom(1)}>
            {Math.round(zoom * 100)}%
          </button>
          <button className="icon-btn" aria-label="放大" onClick={() => setZoom((z) => stepZoom(z, 1))}>
            <Icon name="zoomIn" size={14} />
          </button>
        </div>
      </div>

      {mobile && (
        <SheetMobileBar
          edit={mEdit}
          onEdit={(on) => {
            if (!on && editing) void commit(null);
            setMEdit(on);
            if (!on) setMRange(false);
          }}
          range={mRange}
          onRange={setMRange}
          rangeText={`${colName(sel.range.c0)}${sel.range.r0 + 1}${sel.range.r1 > sel.range.r0 || sel.range.c1 > sel.range.c0 ? `:${colName(sel.range.c1)}${sel.range.r1 + 1}` : ''}`}
          barProps={barProps}
          editing={!!editing}
          onCommit={() => void commit([1, 0])}
          onCancel={cancelEdit}
          canUndo={props.canUndo}
          canRedo={props.canRedo}
          onUndo={() => restore(props.undo())}
          onRedo={() => restore(props.redo())}
          st={activeSt}
          fmt={activeFmt}
          onStyle={styleCmd}
          onBorder={borderCmd}
          onNumFmt={numFmtCmd}
          onRows={(op) => (op === 'rowAbove' ? insertRows(false) : op === 'rowBelow' ? insertRows(true) : op === 'colLeft' ? insertCols(false) : op === 'colRight' ? insertCols(true) : op === 'delRow' ? deleteRows() : deleteCols())}
          onSort={sortCmd}
          filtered={!!sheet.autoFilter}
          onFilter={filterToggle}
          onCopy={() => copySelection('table')}
          onPaste={() => void pasteFromMenu()}
          onClear={() => clear('contents')}
          address={`${colName(sel.active.c)}${sel.active.r + 1}`}
          content={(() => {
            const cell = cellAt(sheet, sel.active);
            return cell ? rawOf(cell) : '';
          })()}
          stats={stats ? { count: stats.count, sum: stats.nums > 0 ? fmtStat(stats.sum) : null } : null}
          query={query}
          hits={hits.length}
          hitIdx={hitIdx}
          onQuery={setQuery}
          onHit={gotoHit}
        />
      )}
      {menu && (
        <div className="ctx-menu sheet-menu" role="menu" style={{ left: menu.x, top: Math.max(8, Math.min(menu.y, window.innerHeight - 640)) }} onMouseDown={(e) => e.stopPropagation()}>
          <button role="menuitem" onClick={() => (copySelection('table', sel.range, true), setMenu(null))}>
            剪切<span className="grow" />
            <span className="sub mono">Ctrl+X</span>
          </button>
          <button role="menuitem" onClick={() => (copySelection('table'), setMenu(null))}>
            复制<span className="grow" />
            <span className="sub mono">Ctrl+C</span>
          </button>
          <button role="menuitem" onClick={() => void pasteFromMenu()}>
            粘贴<span className="grow" />
            <span className="sub mono">Ctrl+V</span>
          </button>
          <button role="menuitem" disabled={!clip} onClick={() => (setMenu(null), setPasteDlg(true))}>
            选择性粘贴…<span className="grow" />
            <span className="sub mono">Ctrl+Alt+V</span>
          </button>
          {bundle && (
            <>
              <div className="ctx-sep" />
              <button role="menuitem" onClick={() => (setMenu(null), openColumnPanel(menuCol, { x: Math.min(menu.x, window.innerWidth - 400), y: menu.y }))}>
                列属性…（{colName(menuCol)} 列）
                <span className="grow" />
                {columnAt(sheet.columns, menuCol) && <span className="sub">{COLUMN_TYPE_LABELS[columnAt(sheet.columns, menuCol)!.type]}</span>}
              </button>
              {columnAt(sheet.columns, menuCol) && (
                <button role="menuitem" onClick={() => (setMenu(null), applyColumn(menuCol, null))}>
                  清除列属性
                </button>
              )}
            </>
          )}
          <div className="ctx-sep" />
          {selDrawing !== null && sheet.drawings?.[selDrawing]?.kind === 'chart' && (
            <button role="menuitem" onClick={() => (setMenu(null), setChartDlg({ index: selDrawing }))}>
              编辑图表…
            </button>
          )}
          {selDrawing !== null && (
            <button role="menuitem" onClick={() => (setMenu(null), run('删除图表', (tx) => deleteDrawing(tx, si, selDrawing), ui()), setSelDrawing(null), focusGrid())}>
              删除{sheet.drawings?.[selDrawing]?.kind === 'chart' ? '图表' : '图片'}
              <span className="grow" />
              <span className="sub mono">Delete</span>
            </button>
          )}
          <button role="menuitem" onClick={() => (setMenu(null), setFmtDlg(true))}>
            设置单元格格式…<span className="grow" />
            <span className="sub mono">Ctrl+1</span>
          </button>
          <div className="ctx-sep" />
          <button role="menuitem" onClick={() => (setMenu(null), openNote(sel.active.r, sel.active.c))}>
            {noteMap.has(cellKey(sel.active.r, sel.active.c)) ? '编辑批注…' : '插入批注…'}
            <span className="grow" />
            <span className="sub mono">Shift+F2</span>
          </button>
          {noteMap.has(cellKey(sel.active.r, sel.active.c)) && (
            <button role="menuitem" onClick={() => (setMenu(null), run('删除批注', (tx) => deleteNote(tx, si, sel.active.r, sel.active.c), ui()), focusGrid())}>
              删除批注
            </button>
          )}
          <div className="ctx-sep" />
          {/* 复制 / 剪切后（虚线框还在）：插入复制（剪切）的单元格，代替“插入”。整行在上方插入行、整列在左边插入列；
              复制的是部分单元格时，选的是整列就右移，否则下移 */}
          {insClip ? (
            <button role="menuitem" onClick={() => insertCopiedCmd(fullCols && !fullRows ? 'right' : 'down')}>
              插入{insClip.cut ? '剪切' : '复制'}的单元格<span className="grow" />
              <span className="sub mono">Ctrl+Shift+=</span>
            </button>
          ) : (
            /* 插入：选的是整列时在左边插入列，否则在上方插入行（同 Ctrl+Shift+=） */
            <button role="menuitem" onClick={() => (fullCols && !fullRows ? insertCols() : insertRows(), setMenu(null))}>
              插入<span className="grow" />
              <span className="sub mono">Ctrl+Shift+=</span>
            </button>
          )}
          {/* 删除：选的是整列时删除所选列，否则删除所选行（同 Ctrl+-） */}
          <button role="menuitem" onClick={() => (fullCols && !fullRows ? deleteCols() : deleteRows(), setMenu(null))}>
            删除<span className="grow" />
            <span className="sub mono">Ctrl+-</span>
          </button>
          <div className="ctx-sep" />
          {hasBreak ? (
            <button role="menuitem" onClick={() => pageBreakCmd(false)}>
              删除分页符
            </button>
          ) : (
            <button role="menuitem" disabled={ar === 0 && ac === 0} onClick={() => pageBreakCmd(true)}>
              插入分页符
            </button>
          )}
          <div className="ctx-sep" />
          <button role="menuitem" onClick={() => (clear('formats'), setMenu(null))}>
            清除格式
          </button>
          <div className="ctx-sep" />
          {fullRows ? (
            <>
              <button role="menuitem" onClick={() => void askRowHeight()}>
                行高…
              </button>
              <button role="menuitem" onClick={() => (hideRows(true), setMenu(null))}>
                隐藏行
              </button>
              <button role="menuitem" onClick={() => (hideRows(false), setMenu(null))}>
                取消隐藏行
              </button>
            </>
          ) : fullCols ? (
            <>
              <button role="menuitem" onClick={() => void askColWidth()}>
                列宽…
              </button>
              <button role="menuitem" onClick={() => (hideCols(true), setMenu(null))}>
                隐藏列
              </button>
              <button role="menuitem" onClick={() => (hideCols(false), setMenu(null))}>
                取消隐藏列
              </button>
            </>
          ) : (
            <>
              <button role="menuitem" onClick={() => void askRowHeight()}>
                行高…
              </button>
              <button role="menuitem" onClick={() => void askColWidth()}>
                列宽…
              </button>
            </>
          )}
        </div>
      )}

      {freezeMenu && (
        <div className="ctx-menu" role="menu" style={{ left: freezeMenu.x, top: freezeMenu.y }} onMouseDown={(e) => e.stopPropagation()} onClick={(e) => e.stopPropagation()}>
          {freezeItems()}
        </div>
      )}

      {props.findMode && (
        <FindReplace
          book={book}
          si={si}
          at={{ si, r: sel.active.r, c: sel.active.c }}
          mode={props.findMode}
          version={version}
          onMode={props.setFindMode}
          onGoto={gotoHit2}
          onReplace={replaceOne}
          onReplaceAll={replaceEvery}
          onClose={() => (props.setFindMode(null), focusGrid())}
        />
      )}

      {printOpen && (
        <PrintDialog
          sheet={sheet}
          styles={book.styles}
          defSize={defFont}
          defFont={defFontName ?? '等线'}
          sel={sel.range}
          multi={rangeSel}
          fileName={props.fileName}
          onSavePage={(page) =>
            run(
              '页面设置',
              (tx) => {
                tx.meta(si);
                sheet.page = page;
              },
              ui(),
            )
          }
          onPrint={(urls, w, h) => {
            setPrintOpen(false);
            void host().printImages(urls, `${w}pt ${h}pt`);
            focusGrid();
          }}
          onPdf={(urls, w, h) => {
            setPrintOpen(false);
            focusGrid();
            void imagesToPdf(urls, w, h)
              .then(props.exportPdf)
              .then(focusGrid);
          }}
          onClose={() => (setPrintOpen(false), focusGrid())}
        />
      )}

      {sortMenu && (
        <div className="ctx-menu" role="menu" style={{ left: sortMenu.x, top: sortMenu.y }} onMouseDown={(e) => e.stopPropagation()} onClick={(e) => e.stopPropagation()}>
          <button role="menuitem" onClick={() => (sortCmd(false), focusGrid())}>
            升序（按 {colName(sel.active.c)} 列）
          </button>
          <button role="menuitem" onClick={() => (sortCmd(true), focusGrid())}>
            降序（按 {colName(sel.active.c)} 列）
          </button>
          <button
            role="menuitem"
            onClick={() => {
              setSortMenu(null);
              const d = dataRegion();
              if (!d) return host().toast('请先选中要排序的数据');
              setSortDialog({ range: d.g, header: d.header });
            }}
          >
            自定义排序…
          </button>
          <div className="ctx-sep" />
          <button
            role="menuitem"
            onClick={() => {
              setSortMenu(null);
              const d = dataRegion();
              if (!d) return host().toast('请先选中要处理的数据');
              setDataDialog({ kind: 'dup', range: d.g, header: d.header });
            }}
          >
            删除重复项…
          </button>
          <button
            role="menuitem"
            onClick={() => {
              setSortMenu(null);
              const g = sel.range;
              if (g.c1 !== g.c0) return host().toast('分列一次只能处理一列，请选中一列里的单元格');
              const r1 = Math.min(g.r1, ext.rows - 1);
              setDataDialog({ kind: 'split', range: { ...g, r1: Math.max(g.r0, r1) }, header: false });
            }}
          >
            分列…
          </button>
          <div className="ctx-sep" />
          {bundle && (
            <>
              <button
                role="menuitem"
                onClick={() => {
                  setSortMenu(null);
                  if (sheet.rows < 2) return host().toast('表格里还没有数据可以分析');
                  setSmartDlg(true);
                }}
              >
                转为智能列…
              </button>
              <div className="ctx-sep" />
            </>
          )}
          <button
            role="menuitem"
            onClick={() => {
              setSortMenu(null);
              const d = dataRegion();
              if (!d || d.g.r1 <= d.g.r0) return host().toast('请先选中要汇总的数据（第一行为字段名，至少再有一行数据）');
              setPivotDlg(d.g);
            }}
          >
            数据透视表…
          </button>
          {sheet.pivot && (
            <button
              role="menuitem"
              onClick={() => {
                setSortMenu(null);
                const ok = run('刷新透视表', (tx) => refreshPivot(tx, si), ui());
                if (ok === false) host().toast('找不到透视表的源数据所在的工作表');
                else if (ok) host().toast('透视表已刷新');
                focusGrid();
              }}
            >
              刷新透视表
            </button>
          )}
          <div className="ctx-sep" />
          <button role="menuitem" onClick={() => (setSortMenu(null), groupSelection(true), focusGrid())}>
            组合{fullCols && !fullRows ? '列' : '行'}
            <span className="grow" />
            <span className="sub mono">Alt+Shift+→</span>
          </button>
          <button role="menuitem" onClick={() => (setSortMenu(null), groupSelection(false), focusGrid())}>
            取消组合{fullCols && !fullRows ? '列' : '行'}
            <span className="grow" />
            <span className="sub mono">Alt+Shift+←</span>
          </button>
          <div className="ctx-sep" />
          <button role="menuitem" className={sheet.autoFilter ? 'on' : undefined} onClick={() => (filterToggle(), focusGrid())}>
            {sheet.autoFilter ? '取消筛选' : '筛选'}
            <span className="grow" />
            <span className="sub mono">Ctrl+Shift+L</span>
          </button>
          <button role="menuitem" disabled={!sheet.autoFilter || !Object.keys(sheet.autoFilter.cols).length} onClick={() => (setSortMenu(null), run('清除筛选', (tx) => setAutoFilter(tx, si, sheet.autoFilter!.ref), ui()), focusGrid())}>
            清除筛选条件
          </button>
          <button role="menuitem" disabled={!sheet.autoFilter} onClick={() => (setSortMenu(null), run('重新应用筛选', (tx) => refilter(tx, si), ui()), focusGrid())}>
            重新应用
          </button>
        </div>
      )}

      {filterMenu && sheet.autoFilter && (
        <FilterMenu
          key={filterMenu.c}
          x={filterMenu.x}
          y={filterMenu.y}
          {...(() => {
            const af = sheet.autoFilter!;
            const ref = { ...af.ref, r1: Math.max(af.ref.r1, currentRegion(sheet, af.ref.r0, af.ref.c0).r1) };
            const env = filterEnv(book, si);
            const hidden = hiddenByOthers(book, si, filterMenu.c);
            const ex = columnExtras(book, si, ref, filterMenu.c, hidden);
            return { values: columnValues(sheet, ref, filterMenu.c, (r) => hidden(r) || !!env.date(r, filterMenu.c)), dates: ex.dates, fills: ex.fills, colors: ex.colors };
          })()}
          current={sheet.autoFilter.cols[filterMenu.c - sheet.autoFilter.ref.c0]}
          onSort={(desc) => {
            const af = sheet.autoFilter!;
            setFilterMenu(null);
            doSort(af.ref, [{ c: filterMenu.c, desc }], true);
            focusGrid();
          }}
          onApply={(f) => columnFilter(filterMenu.c, f)}
          onClose={() => (setFilterMenu(null), focusGrid())}
        />
      )}

      {cfOpen && (
        <CfDialog
          sheet={sheet}
          range={rangeSel ? sel.range : { r0: sel.active.r, c0: sel.active.c, r1: sel.active.r, c1: sel.active.c }}
          onAdd={(rule) => run('条件格式', (tx) => addCfRule(tx, si, rule), ui())}
          onRemove={(rule) => run('删除条件格式', (tx) => removeCfRule(tx, si, rule), ui())}
          onClear={(g) => {
            const n = run('清除条件格式', (tx) => clearCfRules(tx, si, g), ui());
            host().toast(n ? `已清除 ${n} 条规则` : '没有可清除的规则');
          }}
          onClose={() => (setCfOpen(false), focusGrid())}
        />
      )}
      {sortDialog && (
        <SortDialog
          sheet={sheet}
          range={sortDialog.range}
          header={sortDialog.header}
          onOk={(keys, header) => {
            setSortDialog(null);
            doSort(sortDialog.range, keys, header);
            focusGrid();
          }}
          onClose={() => (setSortDialog(null), focusGrid())}
        />
      )}
      {dataDialog?.kind === 'dup' && (
        <DuplicatesDialog
          sheet={sheet}
          range={dataDialog.range}
          header={dataDialog.header}
          onOk={(cols, header) => {
            const g = dataDialog.range;
            setDataDialog(null);
            const res = run('删除重复项', (tx) => removeDuplicates(tx, si, g, cols, header), ui());
            const st = host();
            if (res === null) st.toast('区域里有合并单元格，不能删除重复项');
            else if (res) st.toast(res[0] ? `删除了 ${res[0]} 个重复行，保留 ${res[1]} 行` : '没有发现重复行');
            focusGrid();
          }}
          onClose={() => (setDataDialog(null), focusGrid())}
        />
      )}
      {nameList && (
        <div className="ctx-menu sheet-menu sheet-addr-list" role="menu" style={{ left: Math.max(8, nameList.x), top: nameList.y }} onMouseDown={(e) => e.stopPropagation()}>
          {(book.names ?? [])
            .filter((n) => n.local === undefined || n.local === si)
            .map((n) => (
              <button
                key={`${n.local ?? ''}|${n.name}`}
                role="menuitem"
                onClick={() => {
                  setNameList(null);
                  const hit = nameTarget(book, n.name, si);
                  if (hit?.range) gotoRange(hit.si, hit.range);
                  else host().toast(`“${n.name}”不是一个单元格区域：${n.f}`);
                }}
              >
                {n.name}
                <span className="grow" />
                <span className="sub mono ellipsis" style={{ maxWidth: 140 }}>
                  {n.f}
                </span>
              </button>
            ))}
          {!(book.names ?? []).length && <div className="sub" style={{ padding: '6px 10px' }}>还没有定义名称</div>}
          <div className="ctx-sep" />
          <button role="menuitem" onClick={() => (setNameList(null), setNamesDlg({}))}>
            名称管理器…
          </button>
          <button role="menuitem" onClick={() => (setNameList(null), setNamesDlg({ draft: { f: selectionRef() } }))}>
            用选区定义名称…
          </button>
        </div>
      )}
      {chartDlg && (
        <ChartDialog
          book={book}
          si={si}
          range={chartDlg.range}
          initial={chartDlg.index !== undefined ? sheet.drawings?.[chartDlg.index]?.chart : undefined}
          onClose={() => (setChartDlg(null), focusGrid())}
          onOk={(spec) => {
            const dlg = chartDlg;
            setChartDlg(null);
            if (dlg.index !== undefined) run('编辑图表', (tx) => updateChart(tx, si, dlg.index!, spec), ui());
            else {
              // 放在选区右边（隔一列），480 × 288 像素
              const g = dlg.range!;
              const x = cols.pos[Math.min(g.c1 + 2, cols.n)] ?? 0;
              const y = rows.pos[Math.min(g.r0, rows.n)] ?? 0;
              const at = run('插入图表', (tx) => addChart(tx, si, spec, anchorAtPx(x, y), anchorAtPx(x + 480 * zoom, y + 288 * zoom)), ui());
              if (at !== undefined) setSelDrawing(at);
            }
            focusGrid();
          }}
        />
      )}
      {noteDlg && (
        <NoteDialog
          cell={cellName({ r: noteDlg.r, c: noteDlg.c })}
          initial={noteMap.get(cellKey(noteDlg.r, noteDlg.c))?.text ?? ''}
          author={noteMap.get(cellKey(noteDlg.r, noteDlg.c))?.author ?? userName}
          exists={noteMap.has(cellKey(noteDlg.r, noteDlg.c))}
          onOk={(text) => {
            const d = noteDlg;
            setNoteDlg(null);
            run(text.trim() ? '批注' : '删除批注', (tx) => setNote(tx, si, d.r, d.c, text, userName), ui());
            focusGrid();
          }}
          onDelete={() => {
            const d = noteDlg;
            setNoteDlg(null);
            run('删除批注', (tx) => deleteNote(tx, si, d.r, d.c), ui());
            focusGrid();
          }}
          onClose={() => (setNoteDlg(null), focusGrid())}
        />
      )}
      {protectDlg && (
        <ProtectDialog
          kind={protectDlg}
          title={{ sheet: '保护工作表', 'sheet-off': '撤销工作表保护', book: '保护工作簿结构', 'book-off': '撤销工作簿结构保护' }[protectDlg]}
          initial={sheet.protection?.attrs}
          onClose={() => (setProtectDlg(null), focusGrid())}
          onOk={async (pw, allow) => {
            if (protectDlg === 'sheet') run('保护工作表', (tx) => setSheetProtection(tx, si, protectSheetAttrs(sheet.protection?.attrs, allow, pw)), ui());
            else if (protectDlg === 'book') run('保护工作簿结构', (tx) => setStructureLock(tx, { password: pw }), ui());
            else {
              const attrs = protectDlg === 'sheet-off' ? sheet.protection?.attrs : book.wbProtection?.attrs;
              const good = await verifyPassword(attrs ?? {}, pw, protectDlg === 'book-off' ? 'workbook' : '');
              if (good === null) return '无法校验这种密码算法，请在 Excel 里撤销保护';
              if (!good) return '密码不对';
              if (protectDlg === 'sheet-off') run('撤销工作表保护', (tx) => setSheetProtection(tx, si, null), ui());
              else run('撤销工作簿结构保护', (tx) => setStructureLock(tx, false), ui());
            }
            setProtectDlg(null);
            focusGrid();
            return null;
          }}
        />
      )}
      {bundle && colPanel && (
        <ColumnPanel
          book={book}
          sheet={sheet}
          col={colPanel.c}
          existing={columnAt(sheet.columns, colPanel.c)}
          header={headerRow()}
          at={{ x: colPanel.x, y: colPanel.y }}
          onApply={(meta) => applyColumn(colPanel.c, meta)}
          onClose={() => (setColPanel(null), focusGrid())}
        />
      )}
      {bundle && smartDlg && (
        <SmartColumnsDialog
          book={book}
          sheet={sheet}
          header={headerRow()}
          onClose={() => (setSmartDlg(false), focusGrid())}
          onApply={(header, list) => {
            setSmartDlg(false);
            const n = run('转为智能列', (tx) => applyInferences(tx, si, header, list), ui());
            if (n) host().toast(`已设置 ${n} 列，可撤销`);
            focusGrid();
          }}
        />
      )}
      {bundle && formPanel && (
        <FormFill
          key={formPanel.formId}
          book={book}
          sheet={sheet}
          forms={formsOf()}
          formId={formPanel.formId}
          user={userName}
          activeRow={sel.active.r}
          onSubmit={submitFromForm}
          onDesign={(id) => setDesigner({ form: formsOf().find((f) => f.id === id) ?? null })}
          onExportHtml={(f) => void exportFormHtml(f)}
          onImportResult={(f) => void importFormFile(f)}
          onClose={() => (setFormPanel(null), focusGrid())}
        />
      )}
      {bundle && designer && (
        <FormDesigner
          sheet={sheet}
          header={headerRow()}
          itemId={bundle.itemId}
          form={designer.form}
          newId={bundle.newFormId()}
          items={bundle.items()}
          onClose={() => setDesigner(null)}
          onDelete={
            designer.form
              ? () => {
                  const id = designer.form!.id;
                  setDesigner(null);
                  void host()
                    .guard(bundle.deleteForm(id))
                    .then(() => setFormPanel((p) => (p ? { formId: '' } : p)));
                }
              : undefined
          }
          onSave={(form, created) => {
            setDesigner(null);
            if (created.length) {
              const header = headerRow();
              run(
                '表单字段设为文本列',
                (tx) => {
                  if (!sheet.columns) setHeaderRow(tx, si, header);
                  for (const m of created) setColumn(tx, si, m.col, m);
                },
                ui(),
              );
            }
            void host()
              .guard(bundle.saveForm(form))
              .then((ok) => {
                if (ok === undefined) return;
                host().toast(`已保存表单「${form.title}」`);
                setFormPanel({ formId: form.id });
              });
          }}
        />
      )}
      {namesDlg && (
        <NamesDialog
          book={book}
          si={si}
          selectionRef={selectionRef()}
          draft={namesDlg.draft}
          onSave={(def, replace) => {
            const r = run('定义名称', (tx) => saveName(tx, def, replace), ui());
            return r === undefined ? '保存失败' : r;
          }}
          onDelete={(def) => void run('删除名称', (tx) => deleteName(tx, def), ui())}
          onGoto={(def: DefinedName) => {
            const hit = nameTarget(book, def.name, def.local ?? si);
            setNamesDlg(null);
            if (hit?.range) gotoRange(hit.si, hit.range);
          }}
          onClose={() => (setNamesDlg(null), focusGrid())}
        />
      )}
      {pivotDlg && (
        <PivotDialog
          range={pivotDlg}
          fields={fieldNames(sheet, pivotDlg)}
          initial={defaultConfig(sheet, pivotDlg)}
          onOk={(cfg) => {
            const g = pivotDlg;
            setPivotDlg(null);
            const at = run('数据透视表', (tx) => createPivot(tx, si, g, cfg), ui());
            if (at !== undefined) props.onSheet(at);
          }}
          onClose={() => (setPivotDlg(null), focusGrid())}
        />
      )}
      {fmtDlg && (
        <FormatCellsDialog
          st={activeSt}
          fmt={activeFmt}
          value={sheet.cells[sel.active.r]?.[sel.active.c]?.v}
          date1904={book.date1904}
          defaultFont={defFontName ?? '等线'}
          defaultSize={defFont}
          merged={!!merge}
          multi={sel.range.r0 !== sel.range.r1 || sel.range.c0 !== sel.range.c1}
          onOk={(r) => void formatCellsCmd(r)}
          onClose={() => (setFmtDlg(false), focusGrid())}
        />
      )}
      {pasteDlg && (
        <PasteSpecialDialog
          onOk={(c) => {
            setPasteDlg(false);
            void pasteFromMenu(c.mode, c.transpose, c.opts);
          }}
          onClose={() => (setPasteDlg(false), focusGrid())}
        />
      )}
      {dataDialog?.kind === 'split' && (
        <SplitDialog
          sheet={sheet}
          range={dataDialog.range}
          onOk={async (o) => {
            const g = dataDialog.range;
            setDataDialog(null);
            const w = splitWidth(sheet, g, o);
            // 右边要被覆盖的列里有内容时先确认（同 Excel 的“是否替换目标单元格内容”）
            let busy = false;
            for (let r = g.r0; r <= g.r1 && !busy; r++) for (let c = g.c0 + 1; c < g.c0 + w; c++) if (sheet.cells[r]?.[c]) busy = true;
            if (busy && (await host().askText('右边的单元格里已有内容，确定要替换吗？（输入“是”确认）', '是')) !== '是') return focusGrid();
            run('分列', (tx) => textToColumns(tx, si, g, o), ui());
            focusGrid();
          }}
          onClose={() => (setDataDialog(null), focusGrid())}
        />
      )}

      {detail && (
        <div className="sheet-detail-mask" onMouseDown={() => setDetail(null)}>
          <div className="sheet-detail" role="dialog" aria-label="单元格内容" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.key === 'Escape' && (setDetail(null), focusGrid())}>
            <div className="sheet-detail-head">
              <span className="sheet-detail-addr mono">{cellName(detail)}</span>
              <span className="sub">{detailCell ? kindOf(detailCell) : '空单元格'}</span>
              <span className="grow" />
              <button className="icon-btn" aria-label="关闭" onClick={() => (setDetail(null), focusGrid())}>
                <Icon name="close" size={12} />
              </button>
            </div>
            <textarea readOnly autoFocus value={detailCell?.t ?? ''} />
            {detailCell && rawOf(detailCell) !== detailCell.t && (
              <div className="sheet-detail-raw mono">
                <span className="sub">{detailCell.f ? '公式' : '原始值'}</span> {rawOf(detailCell)}
              </div>
            )}
            <div className="sheet-detail-foot">
              <button className="btn" onClick={() => copyText(detailCell ? rawOf(detailCell) : '', '原始值')} disabled={!detailCell}>
                复制原始值
              </button>
              <button className="btn primary" onClick={() => copyText(detailCell?.t ?? '', '内容')} disabled={!detailCell}>
                复制
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );

  /** 双击列头：按内容估算列宽 */
  function autoFit(c: number) {
    let max = 0;
    for (let r = 0; r < sheet.rows; r++) {
      const t = sheet.cells[r]?.[c]?.t;
      if (!t) continue;
      let w = 0;
      for (const ch of t) w += ch.charCodeAt(0) > 255 ? 13 : 7.5;
      if (w > max) max = w;
    }
    const w = Math.min(600, Math.max(40, Math.round(max + 16)));
    const all = fullCols && c >= sel.range.c0 && c <= sel.range.c1;
    run('列宽', (tx) => setColWidth(tx, si, all ? sel.range.c0 : c, all ? sel.range.c1 : c, { width: w }), ui());
  }
}

function FunnelIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
      <path d="M1.5 2h9L7 6.2V10L5 9V6.2z" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinejoin="round" />
    </svg>
  );
}

function HudItem({ label, value, raw, onCopy }: { label: string; value: string; raw?: string; onCopy: (text: string, label: string) => void }) {
  return (
    <button className="sheet-hud-item" title={`点击复制${label}`} onClick={() => onCopy(raw ?? value, label)}>
      <span className="k">{label}</span>
      <span className="v">{value}</span>
    </button>
  );
}

function cellAt(sheet: SheetData, p: Pos): SCell | undefined {
  return sheet.cells[p.r]?.[p.c];
}

function rawOf(cell: SCell): string {
  if (cell.f) return `=${cell.f}`;
  return typeof cell.v === 'boolean' ? (cell.v ? 'TRUE' : 'FALSE') : String(cell.v);
}

function kindOf(cell: SCell): string {
  if (cell.f) return '公式';
  if (typeof cell.v === 'number') return '数值';
  if (typeof cell.v === 'boolean') return '逻辑值';
  return `文本 · ${[...cell.v].length} 字`;
}

function stepZoom(z: number, dir: 1 | -1): number {
  const i = ZOOMS.findIndex((x) => x >= z - 1e-6);
  const cur = i < 0 ? ZOOMS.length - 1 : i;
  return ZOOMS[Math.max(0, Math.min(ZOOMS.length - 1, cur + dir))];
}

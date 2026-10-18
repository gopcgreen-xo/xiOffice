import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { AUTOSAVE_MS, type AiActionInfo, type EditorTab, type FieldPalette, basename, clearDraft, dirname, draftTime, fitWidthZoom, host, isMobile, loadDraft, pagesToPdf, pickSave, saveDraft, usePinchZoom } from '@wz/editor-shared';
import { applyParaStyle, blockPages, headings, insertToc, PARA_STYLES, paraStyleOf, tocRange, updateToc, type ParaStyle } from '../core/toc';
import { Icon } from '@wz/editor-shared/Icon';
import { FloatViewButton } from '@wz/editor-shared/FloatViewButton';
import '@wz/editor-shared/controls.css';
import './word.css';
import { writeDocx } from '../ooxml/docx';
import { openDocx, patchDocx, type DocxOrigin } from '../ooxml/patch';
import { fragmentToHtml, fragmentToText, htmlToFragment } from '../ooxml/clipboard';
import { replaceAll as replaceAllText, replaceMatch, type Match } from '../core/find';
import { WordFind } from './WordFind';
import { WordMobileBar, type MobileEditApi, type MobileMode, type MobileReviewApi, type TableOp } from './WordMobile';
import { registerOfficeFallback } from '../core/officeMetrics';
import { applyTemplateHits, templateParagraphs } from '../core/templatize';
import { BookmarkDialog, LinkDialog, MergeDialog, ProtectDialog, StylesDialog, WatermarkDialog } from './WordDialogs';
import { activeProtection, addPermRange, checkPassword, clearPermRanges, editError, makeProtection, permRanges, protectLabel, type ProtectMode } from '../core/protect';
import { addBookmark, bookmarkText, clearLink, linkAt, listBookmarks, removeBookmark, setLink, type LinkTarget } from '../core/links';
import { applyCharacterStyle, applyParagraphStyle, catalogFromDoc, clearFormatting, checkStyleName, createStyle, deleteStyle, redefineStyle, renameStyle, withEdits, type StyleCatalog } from '../core/styles';
import { combineRecords, fillRecords, insertMergeField, insertNextRecord, placeholders } from '../core/merge';
import { isDoc, readDoc } from '../doc97/read';
import { WordRuler } from './WordRuler';
import { BORDER_TARGETS, BordersDialog, FontDialog, FontSizeBox, ParagraphDialog } from './FormatDialogs';
import { WordMenu, type MenuItem } from './WordMenu';
import { FieldDialog } from './FieldDialog';
import { allFieldIds, fieldIdsInRange, findField, setFieldInstr, updateFields, type FieldEnv, type UpdateResult } from '../core/fields';
import { fieldAt, setShowCode, showsCode } from '../core/fieldView';
import { nudgeFontSize, stepFontSize } from '../core/format';
import { WordVRuler } from './WordVRuler';
import { marginPatch, marginPatchV, paragraphArea } from '../core/ruler';
import { loadImages, registerImage } from '../core/images';
import { applyPictureFormat, pictureAt, pictureBounds, pictureLocalPoint, pictureFormat, pictureKey, type PictureFormat } from '../core/picture';
import { PictureDialog } from './PictureDialog';
import {
  CN_SIZES,
  FIELD_CHAR,
  imageIds,
  PAPERS,
  PT_PER_CM,
  SECTION_BREAK_LABEL,
  STORY_LABEL,
  charPitch,
  hfList,
  hfPath,
  newDocument,
  paraText,
  sectionList,
  textHeight,
  textWidth,
  type Align,
  type CellEdge,
  type Doc,
  type GridType,
  type HFStory,
  type PageGeom,
  type Paragraph,
  type ParaProps,
  type RunProps,
  type SectionBreak,
  type SectionInfo,
  type Story,
  type Table,
  type TextBox,
  type Watermark,
} from '../core/model';
import { NUM_PRESETS } from '../core/numbering';
import { bodyLineAt, caretRect, comparePos, findLine, hitTest, layoutDocument, moveVertical, rowEnds, selectionRects, type PageLayout, type Pos } from '../core/layout';
import {
  History,
  insertImage,
  addComment,
  commentRanges,
  deleteComment,
  hasRevisions,
  nextRevision,
  revisionAt,
  replyComment,
  resolveRevisions,
  tracking,
  copyFragment,
  insertFragment,
  applyCellBorders,
  applyNumbering,
  applyRunProps,
  changeIndent,
  changeNumLevel,
  mapRuns,
  backspace,
  deleteColumn,
  deleteForward,
  deleteRange,
  deleteRow,
  deleteTable,
  getPara,
  hfLinked,
  insertColumn,
  cellRect,
  clearCellContents,
  distributeColumns,
  mergeTableCells,
  moveColumnLine,
  setCellFormat,
  splitCell,
  unmergeCell,
  insertField,
  insertRaw,
  insertRow,
  insertSectionBreak,
  insertTable,
  insertText,
  moveHorizontal,
  paraLength,
  runPropsAt,
  applyAiText,
  selectedText,
  setHFLinked,
  setParaProps,
  setSectionPage,
  splitParagraph,
  storyEdge,
  tableAt,
  wordCount,
  paragraphPaths,
  type BorderTarget,
} from '../core/ops';
import { registerFontTables } from '../core/fontfile';
import { CanvasMeasurer, EDITOR_FONTS_EA, EDITOR_FONTS_LATIN, drawPage, renderPagesToImages, type SelRect } from '../core/render';

interface Sel {
  anchor: Pos;
  focus: Pos;
}

const PX_PER_PT = 96 / 72;

/** 突出显示的颜色（Word 的常用几种） */
/** 表格底纹常用色（Word 主题色的浅色档） */
const SHADING_OPTIONS: [string, string][] = [
  ['浅灰', 'F2F2F2'],
  ['灰色', 'D9D9D9'],
  ['浅蓝', 'DEEAF6'],
  ['浅绿', 'E2EFD9'],
  ['浅黄', 'FFF2CC'],
  ['浅橙', 'FBE4D5'],
];

const HIGHLIGHT_OPTIONS: [string, string][] = [
  ['黄色', 'FFFF00'],
  ['鲜绿', '00FF00'],
  ['青绿', '00FFFF'],
  ['粉红', 'FF00FF'],
  ['浅灰', 'C0C0C0'],
];

/** 浮动对象所在的段落和文本段下标（正文和表格里） */
function findBoxRun(d: Doc, box: TextBox): { p: import('../core/model').Paragraph; i: number } | null {
  const paras = d.body.flatMap((b) => (b.kind === 'p' ? [b] : b.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs))));
  for (const p of paras) {
    const i = p.runs.findIndex((r) => r.box === box);
    if (i >= 0) return { p, i };
  }
  return null;
}

/** 编辑器内部剪贴板：复制时记下原样的片段；粘贴时系统剪贴板的文字还是它，就用原样的片段（不经 HTML 转换，不丢格式） */
let wordClip: { text: string; blocks: import('../core/model').Block[] } | null = null;
const NO_SEL: SelRect[] = [];

/** 页面内容签名：内容没变的页不重绘。 */
function pageSignature(page: PageLayout): string {
  const part = (lines: PageLayout['body']) =>
    lines.map((l) => `${l.path.join('.')}:${l.line.start}-${l.line.end}${l.line.brk}${l.sectionBreak ?? ''}@${l.x.toFixed(1)},${l.y.toFixed(1)}#${l.line.glyphs.map((g) => g.ch + (g.image ? pictureKey(g.image) : '') + g.x.toFixed(1) + (g.props.bold ? 'b' : '') + (g.props.italic ? 'i' : '') + (g.props.underline ? 'u' : '') + g.props.size + g.family + g.props.color).join('')}`).join('|');
  const floats = page.floats.map((f) => `${f.x.toFixed(1)},${f.y.toFixed(1)},${f.box.w}x${f.box.h}:${f.box.image ? pictureKey(f.box.image) : ''}:${part(f.lines)}`).join('|');
  // 表格框线、底纹：只改线型 / 粗细 / 颜色或底纹时文字不动，也要重画
  const borders = page.borders.map((b) => `${b.x1.toFixed(1)},${b.y1.toFixed(1)},${b.x2.toFixed(1)},${b.y2.toFixed(1)},${b.w ?? ''}${b.color ?? ''}${b.style ?? ''}`).join(';');
  const fills = page.fills.map((f) => `${f.x.toFixed(1)},${f.y.toFixed(1)},${f.w.toFixed(1)},${f.h.toFixed(1)}${f.color}`).join(';');
  return `${part(page.body)}§${part(page.header)}§${part(page.footer)}§${borders}§${fills}§${floats}§${page.watermark ? JSON.stringify(page.watermark) : ''}`;
}
const PAGE_GAP = 16;

type MarksMode = 'para' | 'all' | 'none';
const MARKS_LABEL: Record<MarksMode, string> = { para: '只显示段落标记', all: '显示全部（含分页符、分节符）', none: '不显示' };
let measurer: CanvasMeasurer | null = null;
/** 系统字体的度量只读一次（所有 Word 标签页共用） */
let systemFonts: Promise<number> | null = null;
const loadSystemFonts = () =>
  (systemFonts ??= (host().fontFaces?.() ?? Promise.resolve([]))
    .then(registerFontTables)
    .catch((e) => {
      host().log('warn', 'word', `读取系统字体度量失败：${e}`);
      return 0;
    })
    // 手机上没有宋体、Times New Roman 等：系统字体之后补上它们的度量（行高、分页和电脑一致）
    .then((n) => (isMobile() ? (registerOfficeFallback(), n + 1) : n)));
const getMeasurer = () => (measurer ??= new CanvasMeasurer());

const collapsed = (s: Sel) => comparePos(s.anchor, s.focus) === 0 && s.anchor.story === s.focus.story;
const ordered = (s: Sel): [Pos, Pos] => (comparePos(s.anchor, s.focus) <= 0 ? [s.anchor, s.focus] : [s.focus, s.anchor]);
const caretSel = (p: Pos): Sel => ({ anchor: p, focus: p });

const LINE_OPTIONS: { label: string; value: ParaProps['line'] }[] = [
  { label: '单倍行距', value: { rule: 'auto', value: 1 } },
  { label: '1.15 倍', value: { rule: 'auto', value: 1.15 } },
  { label: '1.5 倍行距', value: { rule: 'auto', value: 1.5 } },
  { label: '2 倍行距', value: { rule: 'auto', value: 2 } },
  { label: '固定值 20 磅', value: { rule: 'exact', value: 20 } },
  { label: '固定值 28 磅', value: { rule: 'exact', value: 28 } },
  { label: '固定值 28.95 磅（公文）', value: { rule: 'exact', value: 28.95 } },
];

export function WordEditor({ tab }: { tab: EditorTab }) {
  const doc = useRef<Doc>(newDocument());
  const [version, setVersion] = useState(0);
  const latestVersion = useRef(version);
  latestVersion.current = version;
  const [savedVersion, setSavedVersion] = useState(0);
  const [loaded, setLoaded] = useState(!tab.path);
  const [sel, setSel] = useState<Sel>(caretSel({ story: 'body', path: [0], offset: 0 }));
  const [pageHint, setPageHint] = useState(0);
  // 按上次使用的缩放比例打开，第一次为 100%
  const [zoom, setZoomState] = useState(() => {
    try {
      const z = Number(localStorage.getItem('word.zoom'));
      return z >= 0.5 && z <= 3 ? z : 1;
    } catch {
      return 1;
    }
  });
  const setZoom = (fn: (z: number) => number) =>
    setZoomState((z) => {
      const next = fn(z);
      try {
        localStorage.setItem('word.zoom', String(next));
      } catch {
        /* 无痕模式等情况下忽略 */
      }
      return next;
    });
  const [pending, setPending] = useState<RunProps | null>(null);
  const [setupOpen, setSetupOpen] = useState(false);
  const [lossy, setLossy] = useState<string[]>([]);
  /** 打开的原文件（字节级保留保存用）；新建文档、原文件读不了时为 null，保存时整篇重写 */
  const origin = useRef<DocxOrigin | null>(null);
  const [focused, setFocused] = useState(false);
  const [fontsTick, setFontsTick] = useState(0);
  // 编辑标记：all 显示段落标记和分页符 / 分节符（默认）→ para 只显示段落标记 → none 都不显示。
  // 默认改成 all 时换了存储键（word.marks.v2），以前存下的 para（旧默认值）不再沿用
  const [marks, setMarks] = useState<MarksMode>(() => {
    try {
      const v = localStorage.getItem('word.marks.v2');
      return v === 'para' || v === 'all' || v === 'none' ? v : 'all';
    } catch {
      return 'all';
    }
  });
  const changeMarks = (next: MarksMode) => {
    setMarks(next);
    try {
      localStorage.setItem('word.marks.v2', next);
    } catch {
      /* 无痕模式等情况下忽略 */
    }
  };
  const history = useRef(new History<Sel>());
  const protectToast = useRef(0);
  const scroller = useRef<HTMLDivElement>(null);
  const input = useRef<HTMLTextAreaElement>(null);
  // 安卓手机版（docs/ANDROID.md §6.2）：没有功能区 / 标尺 / 状态栏，底部工具条；默认查看模式（不弹键盘），按屏宽缩放
  const mobile = isMobile();
  const [mMode, setMModeState] = useState<MobileMode>('view');
  const mModeRef = useRef(mMode);
  mModeRef.current = mMode;
  /** 手机上只有编辑模式才让输入框拿焦点（拿到焦点就会弹键盘） */
  const canFocus = () => !mobile || mModeRef.current === 'edit';
  const [mFit, setMFit] = useState(true);
  /** 手机编辑模式里的“选择”：手指在页面上拖动 = 选文字（不滚动） */
  const [mSelect, setMSelect] = useState(false);
  /** 手机审阅：批注列表、“最终状态”预览（全部接受后的样子，只看） */
  const [mComments, setMComments] = useState(false);
  const [mFinal, setMFinal] = useState(false);
  const pagesEl = useRef<HTMLDivElement>(null);
  const mTouch = useRef<{ x: number; y: number; t: number; timer: number } | null>(null);
  const pageEls = useRef<(HTMLDivElement | null)[]>([]);
  const composing = useRef(false);
  /** 打开的是 .doc / .dot：保存一律另存为 .docx */
  const legacyDoc = useRef(false);
  const [findMode, setFindMode] = useState<'find' | 'replace' | null>(null);
  /** 审阅：修订模式、审阅工具栏、批注栏 */
  const [trackOn, setTrackOnState] = useState(false);
  /** 修订开关存在文档里（settings.xml 的 w:trackRevisions，同 Word）：切换算一次修改，保存后下次打开仍是这个状态 */
  const setTrackOn = (on: boolean) => {
    setTrackOnState(on);
    if (!!doc.current.trackRevisions === on) return;
    doc.current.trackRevisions = on || undefined;
    bump();
  };
  const [reviewBar, setReviewBar] = useState(false);
  /** 插入字段面板的“AI 生成”：进行中（可取消）和结果说明 */
  const [tplBusy, setTplBusy] = useState(false);
  const [tplReport, setTplReport] = useState<string | null>(null);
  const tplCtrl = useRef<AbortController | null>(null);
  const [commentsOpen, setCommentsOpen] = useState(true);
  /** 格式刷：复制的文字格式和段落格式；keep = 双击后连续使用（Esc 结束） */
  const [painter, setPainter] = useState<{ run: RunProps; para: Partial<ParaProps>; keep: boolean } | null>(null);
  /** 导航窗格：按标题跳转 */
  const [navOpen, setNavOpen] = useState(false);
  const [searchText, setSearchText] = useState('');
  const [searchSeq, setSearchSeq] = useState(0);
  const runSearch = () => {
    setSearchSeq((n) => n + 1);
    setFindMode('find');
  };
  const [barTab, setBarTab] = useState<'home' | 'page' | 'tools'>('home');
  /** 打开的对话框：链接、书签、水印、样式、模板与批量生成 */
  const [dialog, setDialog] = useState<null | 'link' | 'bookmark' | 'watermark' | 'styles' | 'merge' | 'protect' | 'font' | 'para' | 'borders'>(null);
  /** 右键菜单的位置（屏幕坐标） */
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  /** 编辑域对话框：要改的域 */
  const [fieldEdit, setFieldEdit] = useState<{ id: string; instr: string } | null>(null);
  /** 表格框线的“笔”：框线下拉按它画，边框对话框里改 */
  const [pen, setPen] = useState<CellEdge>({ w: 0.5, color: 'auto' });
  /** 合并后的字体框显示哪一种：最近选的是中文字体还是西文字体 */
  const [fontPick, setFontPick] = useState<'ea' | 'latin'>('ea');
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [saveMenu, setSaveMenu] = useState(false);
  const saveMenuRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!saveMenu) return;
    const close = (e: MouseEvent) => {
      if (!saveMenuRef.current?.contains(e.target as Node)) setSaveMenu(false);
    };
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && setSaveMenu(false);
    // 捕获阶段监听：文档区域会拦截鼠标事件，冒泡阶段收不到
    document.addEventListener('mousedown', close, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', close, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [saveMenu]);
  /** 标尺：显示 / 隐藏（记住上次的选择）、页面左边在标尺里的位置 */
  const [rulerOn, setRulerOnState] = useState(() => {
    try {
      return localStorage.getItem('word.ruler') !== '0';
    } catch {
      return true;
    }
  });
  const setRulerOn = (on: boolean) => {
    setRulerOnState(on);
    try {
      localStorage.setItem('word.ruler', on ? '1' : '0');
    } catch {
      /* 无痕模式等情况下忽略 */
    }
  };
  const [rulerLeft, setRulerLeft] = useState(0);
  const [rulerTop, setRulerTop] = useState(0);
  const centerEl = useRef<HTMLDivElement>(null);
  // AI 动作（只有 AI 版的工作台提供 host().ai；没有时不显示入口）
  const ai = host().ai;
  const [aiActions, setAiActions] = useState<AiActionInfo[]>([]);
  useEffect(() => {
    void ai?.actions('text-selection').then(setAiActions);
  }, [ai]);
  const author = host().useUserName() || '工作台用户';
  /** 选中的浮动图片（模型里那个 TextBox 对象）与正在进行的拖动 / 缩放 */
  const [floatSel, setFloatSel] = useState<TextBox | null>(null);
  /** 光标停在表格某行的行尾标记上（Word：在这里回车，在下面插入一行）；pos 为同时放在该行最后一格末尾的光标 */
  const [rowEnd, setRowEnd] = useState<{ bi: number; r: number; page: number; pos: Pos } | null>(null);
  /** 拖动表格列线：网格列 g 的右边线 */
  const colDrag = useRef<{ page: number; bi: number; g: number; x0: number } | null>(null);
  const [colGhost, setColGhost] = useState<{ page: number; x: number; y: number; h: number; dx?: number } | null>(null);
  const floatDrag = useRef<{ mode: 'move' | 'resize'; page: number; x0: number; y0: number; fx: number; fy: number; w: number; h: number; rot: number } | null>(null);
  const [floatGhost, setFloatGhost] = useState<{ page: number; x: number; y: number; w: number; h: number } | null>(null);
  const imageInput = useRef<HTMLInputElement>(null);
  const [pictureEdit, setPictureEdit] = useState<{ at: Pos; value: PictureFormat } | null>(null);
  /** Ctrl+Shift+V：下一次粘贴只要文字 */
  const plainPaste = useRef(false);
  const goalX = useRef<number | null>(null);
  const dragging = useRef(false);
  const ppt = PX_PER_PT * zoom;
  const bump = () => setVersion((v) => v + 1);
  /** 命令执行后把键盘焦点还给文档 */
  const refocus = () => setTimeout(() => canFocus() && input.current?.focus({ preventScroll: true }), 0);

  // 打开文件
  useEffect(() => {
    if (!tab.path) return;
    let cancelled = false;
    void (async () => {
      const s = host();
      const bytes = await s.guard(host().readBytes(tab.path!));
      if (!bytes || cancelled) return;
      try {
        // Word 97-2003 二进制文件（.doc / .dot）：转换成文档模型打开，保存时另存为 .docx，原文件不改动
        const legacy = isDoc(bytes);
        legacyDoc.current = legacy;
        const first = legacy ? { doc: readDoc(bytes), origin: null } : await openDocx(bytes);
        let opened = first.doc;
        origin.current = first.origin;
        if (legacy) s.toast('这是 Word 97-2003 格式（.doc），已转换后打开；保存时将另存为 .docx，原文件不会被改动', 6000);
        // 上次没保存就退出了：有自动保存的草稿时提示恢复（草稿也是字节级保留保存的结果，以它为原文件继续编辑）
        let recovered = false;
        const draft = await loadDraft(tab.path!);
        if (draft && !cancelled && typeof draft.data !== 'string') {
          const ok = (await host().confirm(`「${basename(tab.path!)}」有 ${draftTime(draft.time)} 自动保存的修改，上次没有保存就关闭了。\n\n恢复这些修改吗？选择“取消”将打开文件里保存的内容。`));
          if (ok) {
            try {
              const r = await openDocx(draft.data);
              opened = r.doc;
              origin.current = r.origin;
              recovered = true;
            } catch {
              s.toast('草稿已损坏，打开文件里保存的内容');
            }
          }
          if (!recovered) await clearDraft(tab.path!);
        }
        // 图片先解码好再排版绘制（绘制是同步的）
        await loadImages(imageIds(opened));
        if (cancelled) return;
        doc.current = opened;
        setTrackOnState(!!opened.trackRevisions);
        if (hasRevisions(opened) || opened.comments?.length || opened.trackRevisions) {
          setReviewBar(true);
          // 手机：有修订 / 批注的文档直接进审阅（docs/ANDROID.md §5.2）
          if (isMobile()) {
            setMModeState('review');
            mModeRef.current = 'review';
          }
        }
        // 字节级保留保存时，不认识的内容都原样保留，不用提示
        setLossy(origin.current ? [] : doc.current.lossy);
        setSel(caretSel(storyEdge(doc.current, 'body', false)));
        // 刚打开时不算修改（恢复的草稿算未保存的修改）
        const v = Date.now();
        setVersion(v);
        setSavedVersion(recovered ? v - 1 : v);
      } catch (e) {
        s.toast(`无法打开：${e instanceof Error ? e.message : String(e)}`);
      }
      setLoaded(true);
    })();
    return () => {
      cancelled = true;
    };
    // 只在首次打开时读取（另存后路径变化不重新加载）
  }, []);

  // 字体就绪后重新测量；系统字体的度量（行高按字体文件算）读到后重新排版
  useEffect(() => {
    void document.fonts?.ready.then(() => {
      getMeasurer().reset();
      setFontsTick((t) => t + 1);
    });
    void loadSystemFonts().then((n) => n && setFontsTick((t) => t + 1));
  }, []);

  // viewTick：只改显示、不改文档的（显示域代码）重新排版用，不算修改
  const [viewTick, setViewTick] = useState(0);
  // 手机“最终状态”：排一份全部接受修订后的副本（文档本身不变）
  const layout = useMemo(() => {
    if (!mFinal) return layoutDocument(doc.current, getMeasurer());
    const copy = structuredClone(doc.current);
    resolveRevisions(copy, true);
    return layoutDocument(copy, getMeasurer());
  }, [version, fontsTick, viewTick, mFinal]); // eslint-disable-line react-hooks/exhaustive-deps
  const dirty = version !== savedVersion && loaded;
  // 有未保存的修改时定时存草稿；关闭标签时删除草稿
  useEffect(() => {
    if (!tab.path || !dirty) return;
    const path = tab.path;
    const t = window.setInterval(() => {
      void (async () => {
        try {
          await saveDraft(path, origin.current ? await patchDocx(doc.current, origin.current) : writeDocx(doc.current));
        } catch {
          /* 草稿只是保险，失败不影响编辑 */
        }
      })();
    }, AUTOSAVE_MS);
    return () => window.clearInterval(t);
  }, [tab.path, dirty]);
  useEffect(() => () => void (tab.path && clearDraft(tab.path)), [tab.path]);
  useEffect(() => {
    host().updateTab(tab.id, { dirty });
  }, [dirty, tab.id]);

  const story: Story = sel.focus.story;
  const caret = useMemo(() => caretRect(layout, sel.focus, pageHint), [layout, sel.focus, pageHint]);
  // 光标离开行尾标记（点别处、打字、命令移动了光标）时退出行尾状态
  useEffect(() => {
    if (rowEnd && sel.focus !== rowEnd.pos) setRowEnd(null);
  }, [sel.focus, rowEnd]);
  const caretPage = caret?.page ?? pageHint;
  const words = useMemo(() => wordCount(doc.current), [version]);

  // 标尺和光标所在页对齐：滚动、缩放、窗口 / 侧栏变化后重新量页面左边的位置
  const measureRuler = useCallback(() => {
    const el = pageEls.current[caretPage] ?? pageEls.current[0];
    const box = centerEl.current;
    if (!el || !box) return;
    const rect = el.getBoundingClientRect();
    const left = Math.round((rect.left - box.getBoundingClientRect().left - 28) * 10) / 10;
    setRulerLeft((v) => (Math.abs(v - left) < 0.2 ? v : left));
    const sc = scroller.current;
    if (sc) {
      const top = Math.round((rect.top - sc.getBoundingClientRect().top) * 10) / 10;
      setRulerTop((v) => (Math.abs(v - top) < 0.2 ? v : top));
    }
  }, [caretPage]);
  useLayoutEffect(() => {
    if (rulerOn) measureRuler();
  });
  useEffect(() => {
    const sc = scroller.current;
    if (!rulerOn || !sc) return;
    sc.addEventListener('scroll', measureRuler, { passive: true });
    const ro = new ResizeObserver(measureRuler);
    ro.observe(sc);
    return () => {
      sc.removeEventListener('scroll', measureRuler);
      ro.disconnect();
    };
  }, [rulerOn, measureRuler]);

  /** 选区两端在同一张表的不同单元格里：选中的是单元格矩形（行范围、网格列范围） */
  const cellBlock = useMemo(() => {
    const a = tableAt(doc.current, sel.anchor);
    const b = tableAt(doc.current, sel.focus);
    if (!a || !b || a.bi !== b.bi || (a.r === b.r && a.c === b.c)) return null;
    const ra = cellRect(a.table, a.r, a.c);
    const rb = cellRect(b.table, b.r, b.c);
    return { bi: a.bi, r0: Math.min(ra.r0, rb.r0), r1: Math.max(ra.r1, rb.r1), g0: Math.min(ra.g0, rb.g0), g1: Math.max(ra.g1, rb.g1) };
  }, [sel, version]); // eslint-disable-line react-hooks/exhaustive-deps

  const selRectsByPage = useMemo(() => {
    const map = new Map<number, SelRect[]>();
    if (collapsed(sel)) return map;
    if (cellBlock) {
      for (const pg of layout.pages)
        for (const c of pg.cells ?? [])
          if (c.bi === cellBlock.bi && c.r >= cellBlock.r0 && c.r <= cellBlock.r1 && c.g1 >= cellBlock.g0 && c.g0 <= cellBlock.g1) {
            if (!map.has(pg.index)) map.set(pg.index, []);
            map.get(pg.index)!.push({ x: c.x, y: c.y, w: c.w, h: c.h });
          }
      return map;
    }
    for (const r of selectionRects(layout, sel.anchor, sel.focus, pageHint)) {
      if (!map.has(r.page)) map.set(r.page, []);
      map.get(r.page)!.push(r);
    }
    return map;
  }, [layout, sel, pageHint, cellBlock]);

  // ---------- 修改 ----------

  const mutate = useCallback(
    (kind: string, fn: (d: Doc, s: Sel) => Sel | null | void): boolean => {
      // 文档保护：被挡住的操作不改文档、不进撤销（提示最多每 1.5 秒一次，免得连续打字时刷屏）
      const [pa, pb] = ordered(sel);
      const blocked = editError(doc.current, kind, pa, pb);
      if (blocked) {
        if (Date.now() - protectToast.current > 1500) {
          protectToast.current = Date.now();
          host().toast(blocked);
        }
        return false;
      }
      history.current.record(doc.current, sel, kind);
      let next: Sel | null | void;
      // 修订开关是模块级的：每次修改前按本编辑器的状态设置（“只能修订”的保护下强制打开）
      tracking.on = trackOn || activeProtection(doc.current)?.edit === 'trackedChanges';
      tracking.author = author;
      try {
        next = fn(doc.current, sel);
      } catch (e) {
        host().toast(e instanceof Error ? e.message : String(e));
        return false;
      }
      refocus();
      if (next === null) return false;
      if (next) setSel(next);
      bump();
      return true;
    },
    [sel, trackOn, author],
  );

  /** 替换选区与单独删除整表时，Word 对两端段落的合并规则不同。 */
  const clearSelection = (d: Doc, s: Sel, replacement = true): Pos | null => {
    if (collapsed(s)) return s.focus;
    const [a, b] = ordered(s);
    if (a.story !== b.story) return null;
    const p = deleteRange(d, a, b, replacement);
    if (!p) host().toast('选区跨越表格边界，请分别删除');
    return p;
  };

  const typeText = (text: string, paste = false) => {
    if (!text) return;
    goalX.current = null;
    mutate(paste ? 'paste' : 'type', (d, s) => {
      const at = clearSelection(d, s);
      if (!at) return null;
      const p = paste ? insertText(d, at, text, pending) : insertRaw(d, at, text, pending);
      setPending(null);
      return caretSel(p);
    });
  };

  /** AI 动作：把选中的文字交给工作台的预览框，用户确认后替换选区或插入到后面（一步撤销，修订模式下记为修订） */
  const runAi = async (id: string) => {
    if (!ai) return;
    if (activeProtection(doc.current)) return host().toast('文档受保护，AI 不能修改；要修改请先“停止保护”');
    if (collapsed(sel)) return host().toast('请先选中要处理的文字');
    const [a, b] = ordered(sel);
    if (a.story !== b.story) return host().toast('选区跨越了正文和页眉页脚，请分开处理');
    const r = await ai.run(id, selectedText(doc.current, a, b), ['replace', 'insert-after']);
    refocus();
    if (!r) return;
    goalX.current = null;
    mutate('ai', (d, s) => {
      const p = applyAiText(d, s.anchor, s.focus, r.text, r.mode);
      if (!p) host().toast('选区跨越表格边界，请分开处理');
      return p ? caretSel(p) : null;
    });
  };

  /** AI 助手（侧栏对话，AI 直接改文档）：接口里读的都是最新状态；每次修改记一步撤销，失败恢复原样 */
  const live = useRef({ sel, trackOn, author });
  live.current = { sel, trackOn, author };
  const closeAssistant = useRef<(() => void) | null>(null);
  useEffect(() => () => closeAssistant.current?.(), []);
  const openAssistant = () => {
    if (!ai?.openAssistant) return;
    closeAssistant.current?.();
    closeAssistant.current = ai.openAssistant({
      kind: 'word',
      title: tab.title,
      api: {
        doc: () => doc.current,
        selectionText: () => {
          const s = live.current.sel;
          if (collapsed(s)) return '';
          const [a, b] = ordered(s);
          return a.story === b.story ? selectedText(doc.current, a, b) : '';
        },
        edit: (_label, fn) => {
          if (activeProtection(doc.current)) throw new Error('文档受保护，AI 助手不能修改；请先停止保护');
          history.current.record(doc.current, live.current.sel, 'ai');
          tracking.on = live.current.trackOn;
          tracking.author = live.current.author;
          try {
            fn(doc.current);
          } catch (e) {
            const snap = history.current.dropLast();
            if (snap) doc.current = snap.doc;
            bump();
            throw e;
          }
          // 原来的光标位置可能已经不存在：放到文档开头
          setFloatSel(null);
          setSel(caretSel(storyEdge(doc.current, 'body', false)));
          bump();
        },
      },
    });
  };

  /** 复制：纯文本 + HTML（带格式），并记下原样的片段 */
  const writeClip = (dt: DataTransfer) => {
    const [a, b] = ordered(sel);
    if (a.story !== b.story) return;
    const blocks = copyFragment(doc.current, a, b);
    const text = blocks.some((x) => x.kind === 'table') ? fragmentToText(blocks) : selectedText(doc.current, a, b);
    dt.setData('text/plain', text);
    dt.setData('text/html', fragmentToHtml(blocks));
    wordClip = { text, blocks };
  };

  /** 插入图片文件（png / jpg / gif / bmp / webp）：按原始像素大小（96 dpi），比版心宽的缩小到版心宽 */
  const insertImageFile = async (file: Blob) => {
    const mime = file.type;
    if (!/^image\/(png|jpeg|gif|bmp|webp)$/.test(mime)) return host().toast('只支持 png、jpg、gif、bmp、webp 图片');
    const data = new Uint8Array(await file.arrayBuffer());
    let w = 200;
    let h = 150;
    try {
      const bmp = await createImageBitmap(file);
      w = bmp.width * 0.75;
      h = bmp.height * 0.75;
      bmp.close();
    } catch {
      /* 解码不了就用默认大小 */
    }
    const max = textWidth(doc.current.setup);
    if (w > max) {
      h = (h * max) / w;
      w = max;
    }
    const id = registerImage(data, mime);
    await loadImages([id]);
    mutate('image', (d, s) => {
      const at = clearSelection(d, s);
      return at ? caretSel(insertImage(d, at, { id, w: Math.round(w * 10) / 10, h: Math.round(h * 10) / 10 })) : null;
    });
  };

  /** 选中的正好是一张嵌入式图片：返回它所在的段落和文本段 */
  const selectedImage = (() => {
    if (collapsed(sel) || sel.anchor.story !== sel.focus.story || sel.anchor.path.join() !== sel.focus.path.join()) return null;
    const [a, b] = ordered(sel);
    if (b.offset - a.offset !== 1) return null;
    let acc = 0;
    for (const r of getPara(doc.current, a).runs) {
      if (acc === a.offset && r.image) return r.image;
      acc += r.text.length;
      if (acc > a.offset) break;
    }
    return null;
  })();
  const openPictureDialog = () => {
    let at = ordered(sel)[0];
    if (floatSel) {
      let found: Pos | null = null;
      for (const path of paragraphPaths(doc.current, 'body')) {
        let offset = 0;
        for (const run of getPara(doc.current, { story: 'body', path }).runs) {
          if (run.box === floatSel) found = { story: 'body', path, offset };
          offset += run.text.length;
        }
      }
      if (!found) return;
      at = found;
    }
    const run = pictureAt(doc.current, at);
    const value = run && pictureFormat(run);
    if (!value) return;
    setSel({ anchor: at, focus: { ...at, offset: at.offset + 1 } });
    setPictureEdit({ at, value });
  };
  const savePictureDialog = (value: PictureFormat) => {
    const target = pictureEdit;
    if (!target) return;
    const ok = mutate('image', (d) => {
      const blocked = editError(d, 'image', target.at, { ...target.at, offset: target.at.offset + 1 });
      if (blocked) { host().toast(blocked); return null; }
      const run = applyPictureFormat(d, target.at, value);
      if (!run) return null;
      setFloatSel(run.box ?? null);
      return { anchor: target.at, focus: { ...target.at, offset: target.at.offset + 1 } };
    });
    if (ok) setPictureEdit(null);
  };
  /** 改图片大小：k 为比例，fit = 宽度适合版心 */
  const resizeImage = (k: number | 'fit') =>
    mutate('image', (d, s) => {
      const [a] = ordered(s);
      let acc = 0;
      for (const r of getPara(d, a).runs) {
        if (acc === a.offset && r.image) {
          const max = textWidth(d.setup);
          const f = k === 'fit' ? max / r.image.w : k;
          r.image = { ...r.image, w: Math.round(r.image.w * f * 10) / 10, h: Math.round(r.image.h * f * 10) / 10 };
          return;
        }
        acc += r.text.length;
      }
      return null;
    });

  /** 粘贴：本编辑器复制的用原样片段；有 HTML 时按 HTML（保留格式和表格）；否则按纯文本 */
  const pasteFrom = (text: string, html: string) => {
    const norm = (t: string) => t.replace(/\r\n?/g, '\n').replace(/\n$/, '');
    let blocks = wordClip && norm(wordClip.text) === norm(text) && html ? structuredClone(wordClip.blocks) : null;
    if (!blocks && html) {
      try {
        blocks = htmlToFragment(html, pending ?? runPropsAt(doc.current, sel.focus), textWidth(doc.current.setup));
      } catch {
        blocks = null;
      }
    }
    if (!blocks) return typeText(text, true);
    goalX.current = null;
    const pasted = blocks;
    mutate('paste', (d, s) => {
      const at = clearSelection(d, s);
      if (!at) return null;
      setPending(null);
      return caretSel(insertFragment(d, at, pasted));
    });
  };

  // ---------- 审阅 ----------
  const gotoRevision = (back: boolean) => {
    const r = nextRevision(doc.current, back ? ordered(sel)[0] : ordered(sel)[1], back);
    if (!r) return host().toast('文档里没有修订');
    setSel({ anchor: r[0], focus: r[1] });
    refocus();
  };
  /** 接受 / 拒绝：有选区时处理选区里的，否则处理光标所在（或后面）的那一处，然后跳到下一处 */
  const resolveCmd = (accept: boolean, all = false, target?: [Pos, Pos]) => {
    let range: [Pos, Pos] | undefined;
    if (!all) {
      if (target) range = target;
      else if (!collapsed(sel)) range = ordered(sel);
      else {
        // 光标就在某一处修订上（含两端）时处理它；不在修订上才取后面最近的一处
        const r = revisionAt(doc.current, sel.focus) ?? nextRevision(doc.current, { ...sel.focus, offset: Math.max(0, sel.focus.offset - 1) });
        if (!r) return host().toast('文档里没有修订');
        range = r;
      }
    }
    let n = 0;
    mutate('review', (d) => {
      n = resolveRevisions(d, accept, range);
      if (!n) return null;
      if (!range) return undefined;
      // 处理完自动跳到下一处（选中它，悬浮的接受 / 拒绝按钮随之移过去）；没有修订了就把光标放在原处
      const nx = nextRevision(d, { ...range[0], offset: Math.max(0, range[0].offset - 1) });
      return nx ? { anchor: nx[0], focus: nx[1] } : caretSel(range[0]);
    });
    if (!n) host().toast(all ? '文档里没有修订' : '这里没有修订');
    // 按钮抢走了焦点，键盘（Delete / Backspace 等）就不再作用于文档
    refocus();
  };
  /** 光标（或选区起点）所在的那一处修订：在它下方显示悬浮的接受 / 拒绝 */
  const floatRev = useMemo(() => {
    if (!reviewBar || mFinal || sel.focus.story !== 'body') return null;
    const r = revisionAt(doc.current, ordered(sel)[0]);
    if (!r) return null;
    const rects = selectionRects(layout, r[0], r[1], pageHint);
    const last = rects[rects.length - 1];
    return last ? { rect: last, range: r } : null;
  }, [reviewBar, sel, layout, pageHint, version]); // eslint-disable-line react-hooks/exhaustive-deps
  const newComment = async () => {
    if (collapsed(sel) || sel.anchor.story !== 'body' || sel.focus.story !== 'body') return host().toast('请先在正文里选中要批注的文字');
    const [a, b] = ordered(sel);
    const text = await host().askText('批注内容', '');
    if (!text?.trim()) return refocus();
    setCommentsOpen(true);
    mutate('comment', (d) => (addComment(d, a, b, text.trim(), author) ? undefined : null));
  };
  const replyCmd = async (id: string) => {
    const text = await host().askText('回复批注', '');
    if (!text?.trim()) return refocus();
    mutate('comment', (d) => void replyComment(d, id, text.trim(), author));
  };
  const ranges = useMemo(() => commentRanges(doc.current), [version]);
  const commentRectsByPage = useMemo(() => {
    const map = new Map<number, SelRect[]>();
    if (mFinal) return map;
    for (const [a, b] of ranges.values()) {
      for (const r of selectionRects(layout, a, b, pageHint)) {
        if (!map.has(r.page)) map.set(r.page, []);
        map.get(r.page)!.push(r);
      }
    }
    return map;
  }, [layout, ranges, pageHint, mFinal]);
  /** 允许编辑的区域也用同样的浅黄底标出（有保护时才显示） */
  const permRectsByPage = useMemo(() => {
    const map = new Map<number, SelRect[]>();
    if (!doc.current.protection?.enforced) return map;
    for (const r0 of permRanges(doc.current)) {
      for (const r of selectionRects(layout, r0.start, r0.end, pageHint)) {
        if (!map.has(r.page)) map.set(r.page, []);
        map.get(r.page)!.push(r);
      }
    }
    return map;
  }, [layout, version, pageHint]); // eslint-disable-line react-hooks/exhaustive-deps

  /** 删除选中的浮动图片 */
  const deleteFloat = () => {
    const box = floatSel;
    if (!box) return;
    setFloatSel(null);
    mutate('float', (d) => {
      const hit = findBoxRun(d, box);
      if (!hit) return null;
      hit.p.runs.splice(hit.i, 1);
    });
  };
  /** 浮动图片移到页面上的 (x, y)（磅），改为相对页面定位；或改大小 */
  const moveFloat = (box: TextBox, patch: { x?: number; y?: number; w?: number; h?: number }) =>
    mutate('float', (d) => {
      const hit = findBoxRun(d, box);
      if (!hit) return null;
      const b = hit.p.runs[hit.i].box!;
      if (patch.x !== undefined) b.posH = { rel: 'page', offset: Math.round(patch.x * 10) / 10 };
      if (patch.y !== undefined) b.posV = { rel: 'page', offset: Math.round(patch.y * 10) / 10 };
      if (patch.w !== undefined && patch.h !== undefined) {
        b.w = patch.w;
        b.h = patch.h;
        if (b.image) b.image = { ...b.image, w: patch.w, h: patch.h };
      }
    });

  const undo = () => {
    if (activeProtection(doc.current)) return host().toast('文档受保护，不能撤销；要修改请先“停止保护”');
    const snap = history.current.undo(doc.current, sel);
    if (!snap) return;
    setFloatSel(null);
    // 修订开关不随撤销 / 重做变化（同 Word）
    snap.doc.trackRevisions = doc.current.trackRevisions;
    doc.current = snap.doc;
    setSel(snap.sel);
    bump();
  };
  const redo = () => {
    if (activeProtection(doc.current)) return host().toast('文档受保护，不能重做；要修改请先“停止保护”');
    const snap = history.current.redo(doc.current, sel);
    if (!snap) return;
    setFloatSel(null);
    snap.doc.trackRevisions = doc.current.trackRevisions;
    doc.current = snap.doc;
    setSel(snap.sel);
    bump();
  };

  // ---------- 当前格式 ----------

  const curRun: RunProps = pending ?? runPropsAt(doc.current, sel.focus);
  const fontShown = fontPick === 'latin' && curRun.fontLatin !== curRun.fontEA ? curRun.fontLatin : curRun.fontEA;
  const curPara: ParaProps = getPara(doc.current, sel.focus).props;
  const inTable = tableAt(doc.current, sel.focus);

  const setRun = (patch: Partial<RunProps>) => {
    if (collapsed(sel)) {
      setPending({ ...curRun, ...patch });
      refocus();
      return;
    }
    const [a, b] = ordered(sel);
    mutate('format', (d) => applyRunProps(d, a, b, patch));
  };
  const setPara = (patch: Partial<ParaProps>) => {
    const [a, b] = ordered(sel);
    mutate('para', (d) => setParaProps(d, a, b, patch));
  };
  /** 增大 / 减小字号：byList = 按字号表（Ctrl+Shift+> / <），否则每次 1 磅；选中多种字号时各自变 */
  const growFont = (dir: 1 | -1, byList: boolean) => {
    const next = (size: number) => (byList ? stepFontSize(size, dir) : nudgeFontSize(size, dir));
    if (collapsed(sel)) return setRun({ size: next(curRun.size) });
    const [a, b] = ordered(sel);
    mutate('format', (d) => void mapRuns(d, a, b, (p, run) => (run.opaque ? p : { ...p, size: next(p.size) })));
  };
  /** 清除格式：paragraphs = true 时连段落格式一起回到“正文”（同 Word 的“清除所有格式”），否则只清文字格式（Ctrl+Space） */
  const clearFormatCmd = (paragraphs: boolean) => {
    const [a, b] = ordered(sel);
    if (a.story !== b.story) return host().toast('选区跨越了正文和页眉页脚，请分开处理');
    if (collapsed(sel) && !paragraphs) {
      const p = getPara(doc.current, sel.focus);
      const cat = withEdits(styleCat(), doc.current.styleEdits);
      setPending({ ...cat.base(p.props.styleId ?? (p.props.style ? (cat.idOf(p.props.style) ?? undefined) : undefined)).run });
      return refocus();
    }
    setPending(null);
    mutate('format', (d) => void clearFormatting(d, styleCat(), a, b, paragraphs));
  };
  /** 增加 / 减少缩进量：编号段落改级别，其它段落左缩进移到下一个 / 上一个默认制表位 */
  const indentCmd = (dir: 1 | -1) => {
    if (curPara.num) return levelCmd(dir);
    const [a, b] = ordered(sel);
    mutate('para', (d) => (changeIndent(d, a, b, dir) ? undefined : null));
  };
  /** 一个字符多宽（段落对话框按字符换算缩进用；同排版：有字符网格且对齐网格时按网格，否则按段首文字的字号） */
  const charPtOf = (p: Paragraph): number => {
    const t = doc.current.setup.grid.type;
    if ((t === 'linesAndChars' || t === 'snapToChars') && p.props.snapToGrid) return charPitch(doc.current);
    return p.runs[0]?.props.size ?? p.mark.size;
  };
  /** 拿起格式：光标处的文字格式 + 所在段落的对齐、缩进、间距、行距 */
  const pickFormat = (keep: boolean) => {
    if (painter) return setPainter(null);
    const { rev: _r, ...run } = curRun;
    const p = curPara;
    setPainter({
      run,
      para: { align: p.align, indentLeft: p.indentLeft, indentRight: p.indentRight, firstLine: p.firstLine, firstLineChars: p.firstLineChars, spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter, spaceBeforeLines: p.spaceBeforeLines, spaceAfterLines: p.spaceAfterLines, line: p.line },
      keep,
    });
    refocus();
  };
  /** 刷到选区上：选中整段（含段落标记）时连段落格式一起刷，否则只刷文字格式（同 Word） */
  const applyPainter = (s: Sel) => {
    if (!painter || collapsed(s) || s.anchor.story !== s.focus.story) return;
    const [a, b] = ordered(s);
    mutate('format', (d) => {
      applyRunProps(d, a, b, painter.run);
      const wholePara = a.offset === 0 && (a.path.join() !== b.path.join() || b.offset >= paraLength(getPara(d, b)));
      if (wholePara) setParaProps(d, a, b, painter.para);
    });
    if (!painter.keep) setPainter(null);
  };

  // ---------- 键盘 ----------

  const lineEdge = (end: boolean): Pos => {
    const hit = findLine(layout, sel.focus, pageHint);
    if (!hit) return sel.focus;
    const { line } = hit.pl;
    const last = line.brk === 'line' || line.brk === 'page' || line.brk === 'column' ? line.end - 1 : line.end;
    return { ...sel.focus, offset: end ? last : line.start };
  };

  const moveTo = (p: Pos, extend: boolean, page?: number) => {
    setSel((s) => (extend ? { anchor: s.anchor, focus: p } : caretSel(p)));
    if (page !== undefined) setPageHint(page);
    setPending(null);
  };

  /** 表格第 r 行最后一格最后一段的末尾（行尾标记左边的位置） */
  const rowLastPos = (bi: number, r: number): Pos | null => {
    const t = doc.current.body[bi];
    if (!t || t.kind !== 'table' || !t.rows[r]) return null;
    const c = t.rows[r].cells.length - 1;
    const pi = t.rows[r].cells[c].paragraphs.length - 1;
    return { story: 'body', path: [bi, r, c, pi], offset: paraLength(t.rows[r].cells[c].paragraphs[pi]) };
  };
  /** 把光标放到第 r 行的行尾标记上 */
  const enterRowEnd = (bi: number, r: number, page: number) => {
    const pos = rowLastPos(bi, r);
    if (!pos) return false;
    // 行跨页时优先用给定的页，否则用这一行出现的第一页
    const has = (i: number) => !!layout.pages[i] && rowEnds(layout.pages[i]).some((e) => e.bi === bi && e.r === r);
    const onPage = has(page) ? page : layout.pages.findIndex((_, i) => has(i));
    if (onPage < 0) return false;
    setSel(caretSel(pos));
    setRowEnd({ bi, r, page: onPage, pos });
    setPending(null);
    goalX.current = null;
    return true;
  };

  const nextCell = (dir: 1 | -1) => {
    const t = tableAt(doc.current, sel.focus);
    if (!t) return false;
    const cells: number[][] = [];
    t.table.rows.forEach((row, r) => row.cells.forEach((_, c) => cells.push([t.bi, r, c, 0])));
    const i = cells.findIndex((c) => c[1] === t.r && c[2] === t.c);
    if (dir > 0 && i === cells.length - 1) {
      mutate('table', (d, s) => {
        const p = insertRow(d, s.focus, true);
        return p ? caretSel(p) : null;
      });
      return true;
    }
    const target = cells[i + dir];
    if (target) moveTo({ story: 'body', path: target, offset: 0 }, false);
    return true;
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.nativeEvent.isComposing || composing.current) return;
    if (floatSel && !e.ctrlKey && !e.metaKey) {
      const f = layout.pages.flatMap((pg) => pg.floats.map((x) => ({ ...x, page: pg.index }))).find((x) => x.box === floatSel);
      const step = e.shiftKey ? 10 : 1;
      const nudge: Record<string, [number, number]> = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] };
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        return deleteFloat();
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        return setFloatSel(null);
      }
      if (nudge[e.key] && f) {
        e.preventDefault();
        return void moveFloat(floatSel, { x: f.x + nudge[e.key][0], y: f.y + nudge[e.key][1] });
      }
      if (e.key.length === 1 || e.key === 'Enter') setFloatSel(null);
    }
    if (rowEnd && !e.ctrlKey && !e.metaKey && !e.altKey) {
      const { bi, r } = rowEnd;
      if (e.key === 'Enter' && !e.shiftKey) {
        // 行尾回车：在这一行下面插入一行（按这一行里没有纵向合并的格定位，避免插到合并块下面）
        e.preventDefault();
        setRowEnd(null);
        mutate('table', (d) => {
          const t = d.body[bi];
          if (!t || t.kind !== 'table') return null;
          const c = Math.max(0, t.rows[r].cells.findIndex((x) => x.vMerge === 'none'));
          const p = insertRow(d, { story: 'body', path: [bi, r, c, 0], offset: 0 }, true);
          return p ? caretSel({ ...p, path: [p.path[0], p.path[1], 0, 0] }) : null;
        });
        return;
      }
      if (e.key === 'ArrowLeft' && !e.shiftKey) {
        // 回到这一行最后一格的末尾
        e.preventDefault();
        setRowEnd(null);
        return;
      }
      if (e.key === 'ArrowRight' && !e.shiftKey) {
        // 到下一行第一格；最后一行则到表格后面的段落
        e.preventDefault();
        const t = doc.current.body[bi];
        const next: Pos = t?.kind === 'table' && r + 1 < t.rows.length ? { story: 'body', path: [bi, r + 1, 0, 0], offset: 0 } : { story: 'body', path: [Math.min(bi + 1, doc.current.body.length - 1)], offset: 0 };
        setRowEnd(null);
        return moveTo(next, false);
      }
      setRowEnd(null);
    }
    if (painter && e.key === 'Escape') {
      e.preventDefault();
      return setPainter(null);
    }
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    const shift = e.shiftKey;
    const handled = () => e.preventDefault();

    // 域：F9 更新，Shift+F9 切换光标处的域代码，Alt+F9 切换全部（同 Word）
    if (k === 'F9' && !mod) {
      handled();
      if (e.altKey) return toggleFieldCodes(true);
      return shift ? toggleFieldCodes(false) : updateFieldCmd();
    }

    if (mod) {
      const key = k.toLowerCase();
      if (key === 'z') return handled(), shift ? redo() : undo();
      if (key === 'y') return handled(), redo();
      if (key === 's') return handled(), void save(false);
      if (key === 'p') return handled(), void print();
      if (key === 'k') return handled(), openLinkDialog();
      if (key === 'b') return handled(), setRun({ bold: !curRun.bold });
      if (key === 'i') return handled(), setRun({ italic: !curRun.italic });
      if (key === 'u') return handled(), setRun({ underline: !curRun.underline });
      // Ctrl+Shift+= 上标，Ctrl+= 下标（同 Word）
      if (k === '+' || (k === '=' && shift)) return handled(), setRun({ vertAlign: curRun.vertAlign === 'superscript' ? undefined : 'superscript' });
      if (k === '=') return handled(), setRun({ vertAlign: curRun.vertAlign === 'subscript' ? undefined : 'subscript' });
      if (key === 'e') return handled(), setPara({ align: 'center' });
      if (key === 'l') return handled(), setPara({ align: 'left' });
      if (key === 'r') return handled(), setPara({ align: 'right' });
      if (key === 'j') return handled(), setPara({ align: 'both' });
      // 字号：Ctrl+] / Ctrl+[ 每次 1 磅，Ctrl+Shift+> / < 按字号表（同 Word）
      if (e.code === 'BracketRight' && !shift) return handled(), growFont(1, false);
      if (e.code === 'BracketLeft' && !shift) return handled(), growFont(-1, false);
      if (e.code === 'Period' && shift) return handled(), growFont(1, true);
      if (e.code === 'Comma' && shift) return handled(), growFont(-1, true);
      // Ctrl+Space 清除文字格式，Ctrl+M / Ctrl+Shift+M 增加 / 减少缩进量，Ctrl+D 字体对话框
      if (e.code === 'Space' && !shift) return handled(), clearFormatCmd(false);
      if (key === 'm') return handled(), indentCmd(shift ? -1 : 1);
      if (key === 'd' && !shift) return handled(), setDialog('font');
      if (key === 'f' || key === 'h') {
        handled();
        setFindMode(key === 'f' ? 'find' : 'replace');
        return;
      }
      if (key === 'v' && shift) {
        // 只粘贴文字：让浏览器照常触发 paste 事件，事件里按纯文本处理
        plainPaste.current = true;
        return;
      }
      if (key === 'a') {
        handled();
        setSel({ anchor: storyEdge(doc.current, story, false, sel.focus.path), focus: storyEdge(doc.current, story, true, sel.focus.path) });
        return;
      }
      if (k === 'Enter') return handled(), typeText('\f');
      if (k === 'Home' || k === 'End') return handled(), moveTo(storyEdge(doc.current, story, k === 'End', sel.focus.path), shift);
    }

    switch (k) {
      case 'ArrowLeft':
      case 'ArrowRight': {
        handled();
        goalX.current = null;
        const dir = k === 'ArrowLeft' ? -1 : 1;
        if (!shift && !collapsed(sel)) return moveTo(dir < 0 ? ordered(sel)[0] : ordered(sel)[1], false);
        // 一行最后一格的末尾按 →：先到行尾标记（同 Word）
        if (dir > 0 && !shift && story === 'body' && sel.focus.path.length === 4) {
          const [bi, r] = sel.focus.path;
          const last = rowLastPos(bi, r);
          if (last && last.path.join() === sel.focus.path.join() && last.offset === sel.focus.offset && enterRowEnd(bi, r, caretPage)) return;
        }
        return moveTo(moveHorizontal(doc.current, sel.focus, dir), shift);
      }
      case 'ArrowUp':
      case 'ArrowDown': {
        handled();
        if (!caret) return;
        goalX.current ??= caret.x;
        const r = moveVertical(layout, sel.focus, k === 'ArrowUp' ? -1 : 1, pageHint, goalX.current);
        if (r) {
          setSel((s) => (shift ? { anchor: s.anchor, focus: r.pos } : caretSel(r.pos)));
          setPageHint(r.page);
        }
        return;
      }
      case 'Home':
      case 'End':
        handled();
        goalX.current = null;
        return moveTo(lineEdge(k === 'End'), shift);
      case 'Backspace':
      case 'Delete':
        handled();
        goalX.current = null;
        mutate('delete', (d, s) => {
          if (!collapsed(s)) {
            if (k === 'Delete' && cellBlock) {
              const p = clearCellContents(d, s.anchor, s.focus);
              return p ? caretSel(p) : null;
            }
            const p = clearSelection(d, s, false);
            return p ? caretSel(p) : null;
          }
          const p = k === 'Backspace' ? backspace(d, s.focus) : deleteForward(d, s.focus);
          return p ? caretSel(p) : null;
        });
        return;
      case 'Enter':
        handled();
        if (shift) return typeText('\n');
        mutate('split', (d, s) => {
          const at = clearSelection(d, s);
          if (!at) return null;
          // 空的编号段落回车：结束编号（与 Word 一致）
          const p = getPara(d, at);
          if (p.props.num && paraLength(p) === 0) {
            p.props = { ...p.props, num: null, indentLeft: 0, firstLine: 0, firstLineChars: null };
            return caretSel(at);
          }
          return caretSel(splitParagraph(d, at));
        });
        return;
      case 'Tab':
        handled();
        if (nextCell(shift ? -1 : 1)) return;
        // 编号段落段首：Tab 降级，Shift+Tab 升级
        if (curPara.num && (sel.focus.offset === 0 || !collapsed(sel))) {
          const [a, b] = ordered(sel);
          mutate('para', (d) => (changeNumLevel(d, a, b, shift ? -1 : 1) ? undefined : null));
          return;
        }
        return typeText('\t');
      case 'Escape':
        if (story !== 'body') {
          handled();
          moveTo(storyEdge(doc.current, 'body', false), false);
        }
        return;
      default:
    }
  };

  const flushInput = (value: string) => {
    if (input.current) input.current.value = '';
    typeText(value);
  };

  // ---------- 鼠标 ----------

  const pointAt = (e: React.PointerEvent | React.MouseEvent, page: number) => {
    const el = pageEls.current[page]!;
    const r = el.getBoundingClientRect();
    return { x: (e.clientX - r.left) / ppt, y: (e.clientY - r.top) / ppt };
  };

  /** 页面上某个纵坐标属于正文还是该页实际使用的页眉 / 页脚分区 */
  const regionOf = (y: number, page: number, x?: number): Story => {
    const s = layout.pages[page]?.setup ?? doc.current.setup;
    const hf = layout.pages[page]?.hf;
    const outside = hf && (y < s.marginTop || y > s.height - s.marginBottom);
    if (!outside) return 'body';
    // 页边距里也可能有正文内容（浮动表格、超出版心的内容）：点在正文的某一行上，就按正文处理，
    // 不然这些文字永远点不中（会被当成点页眉 / 页脚）
    if (x !== undefined && bodyLineAt(layout.pages[page]!, x, y)) return 'body';
    return y < s.marginTop ? hf!.header : hf!.footer;
  };

  const onPagePointerDown = (e: React.PointerEvent<HTMLDivElement>, page: number) => {
    if (e.button !== 0) return;
    e.preventDefault();
    if (canFocus()) input.current?.focus();
    const pt = pointAt(e, page);
    // 浮动图片：浮于文字上方的优先，最上面的先命中；右下角是缩放手柄
    const floats = [...layout.pages[page].floats].filter((f) => f.box.image && !f.hf).sort((a, b) => Number(b.box.behind) - Number(a.box.behind)).reverse();
    const handle = 8 / ppt;
    const sel0 = floatSel && floats.find((f) => f.box === floatSel);
    const local = sel0 && pictureLocalPoint(sel0.box.image!, pt.x - sel0.x, pt.y - sel0.y);
    const onHandle = sel0 && local && Math.abs(local.x - sel0.box.w) <= handle && Math.abs(local.y - sel0.box.h) <= handle;
    const hitFloat = onHandle ? sel0 : floats.find((f) => {
      const p = pictureLocalPoint(f.box.image!, pt.x - f.x, pt.y - f.y);
      return p.x >= 0 && p.x <= f.box.w && p.y >= 0 && p.y <= f.box.h;
    });
    if (hitFloat && story === 'body') {
      setFloatSel(hitFloat.box);
      floatDrag.current = { mode: onHandle ? 'resize' : 'move', page, x0: pt.x, y0: pt.y, fx: hitFloat.x, fy: hitFloat.y, w: hitFloat.box.w, h: hitFloat.box.h, rot: hitFloat.box.image?.rot ?? 0 };
      e.currentTarget.setPointerCapture(e.pointerId);
      return;
    }
    setFloatSel(null);
    // Word selects an inline picture with one click rather than placing a text
    // caret on its left/right edge. Use the glyph's source offset, including
    // zero-width bookmarks and fields preceding the image.
    const pg = layout.pages[page];
    for (const pl of [...pg.body, ...pg.header, ...pg.footer]) {
      if (pl.story !== story) continue;
      for (const glyph of pl.line.glyphs) {
        if (!glyph.image || glyph.math || glyph.offset === undefined) continue;
        const bounds = pictureBounds(glyph.image);
        const x = pl.x + glyph.x + (bounds.w - glyph.image.w) / 2;
        const y = pl.y + pl.line.baseline - bounds.h + (bounds.h - glyph.image.h) / 2;
        const local = pictureLocalPoint(glyph.image, pt.x - x, pt.y - y);
        if (local.x < 0 || local.y < 0 || local.x > glyph.image.w || local.y > glyph.image.h) continue;
        const at: Pos = { story: pl.story, path: pl.path, offset: glyph.offset };
        if (!pictureAt(doc.current, at)?.image) continue;
        setSel({ anchor: at, focus: { ...at, offset: at.offset + 1 } });
        setPageHint(page);
        setRowEnd(null);
        dragging.current = false;
        return;
      }
    }
    const line = story === 'body' ? columnLineAt(page, pt) : null;
    if (line) {
      colDrag.current = { page, bi: line.bi, g: line.g, x0: pt.x };
      setColGhost({ page, x: line.x, y: line.y, h: line.h });
      e.currentTarget.setPointerCapture(e.pointerId);
      return;
    }
    const region = regionOf(pt.y, page, pt.x);
    // 表格行右边框外面：光标放到这一行的行尾标记上
    if (story === 'body' && region === 'body' && !e.shiftKey) {
      const re = rowEnds(layout.pages[page]).find((x) => pt.x > x.x + 1 && pt.y >= x.y && pt.y < x.y + x.h);
      if (re && enterRowEnd(re.bi, re.r, page)) {
        setPageHint(page);
        return;
      }
    }
    const target = story === 'body' ? 'body' : region === 'body' ? story : region;
    const hit = hitTest(layout, page, pt.x, pt.y, target);
    if (!hit) return;
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
      const lk = linkAt(doc.current, hit);
      if (lk) return followLink(lk.link);
    }
    setPageHint(page);
    goalX.current = null;
    setPending(null);
    setSel((s) => (e.shiftKey && s.focus.story === hit.story ? { anchor: s.anchor, focus: hit } : caretSel(hit)));
    dragging.current = true;
    e.currentTarget.setPointerCapture(e.pointerId);
  };

  /** 鼠标附近（3 像素内）的表格竖线：返回网格列 g（线在它的右边） */
  const columnLineAt = (page: number, pt: { x: number; y: number }) => {
    const tol = 3 / ppt;
    for (const c of layout.pages[page].cells ?? []) {
      if (pt.y < c.y || pt.y > c.y + c.h || Math.abs(pt.x - (c.x + c.w)) > tol) continue;
      // 整张表在本页的上下范围，拖动时画整条线
      const same = layout.pages[page].cells!.filter((o) => o.bi === c.bi);
      const y = Math.min(...same.map((o) => o.y));
      return { bi: c.bi, g: c.g1, x: c.x + c.w, y, h: Math.max(...same.map((o) => o.y + o.h)) - y };
    }
    return null;
  };

  const onPagePointerMove = (e: React.PointerEvent<HTMLDivElement>, page: number) => {
    const cd = colDrag.current;
    if (cd) {
      const pt = pointAt(e, cd.page);
      setColGhost((g) => g && { ...g, dx: pt.x - cd.x0 });
      return;
    }
    if (!dragging.current && !floatDrag.current && story === 'body') {
      e.currentTarget.style.cursor = columnLineAt(page, pointAt(e, page)) ? 'col-resize' : 'text';
    }
    const fd = floatDrag.current;
    if (fd) {
      const pt = pointAt(e, fd.page);
      const dx = pt.x - fd.x0;
      const dy = pt.y - fd.y0;
      if (fd.mode === 'move') setFloatGhost({ page: fd.page, x: fd.fx + dx, y: fd.fy + dy, w: fd.w, h: fd.h });
      else {
        // 等比缩放：按拖动距离较大的方向算比例
        const angle = fd.rot * Math.PI / 180, c = Math.cos(angle), s = Math.sin(angle);
        const k = Math.max(0.05, Math.max((fd.w + c * dx + s * dy) / fd.w, (fd.h - s * dx + c * dy) / fd.h));
        const dw = fd.w * (k - 1), dh = fd.h * (k - 1);
        setFloatGhost({ page: fd.page, x: fd.fx + ((c - 1) * dw - s * dh) / 2, y: fd.fy + (s * dw + (c - 1) * dh) / 2, w: fd.w * k, h: fd.h * k });
      }
      return;
    }
    if (!dragging.current) return;
    const pt = pointAt(e, page);
    // 拖出本页时按页内最近行处理
    const hit = hitTest(layout, page, pt.x, pt.y, sel.anchor.story);
    if (hit) setSel((s) => ({ anchor: s.anchor, focus: hit }));
  };

  const onPageDoubleClick = (e: React.MouseEvent<HTMLDivElement>, page: number) => {
    const pt = pointAt(e, page);
    const region = regionOf(pt.y, page, pt.x);
    // 双击页眉页脚区域进入编辑，编辑页眉页脚时双击正文返回
    if ((story === 'body' && region !== 'body') || (story !== 'body' && region === 'body')) {
      const hit = hitTest(layout, page, pt.x, pt.y, region) ?? storyEdge(doc.current, region, false, hfPath(layout.pages[page].hf.owner, 0));
      setPageHint(page);
      setSel(caretSel(hit));
      return;
    }
    // 双击选中连续的同类字符
    const p = sel.focus;
    const text = paraText(getPara(doc.current, p));
    const cls = (c: string | undefined) => (!c ? -1 : /\s/.test(c) ? 0 : /[A-Za-z0-9_]/.test(c) ? 1 : /[，。、；：？！“”‘’（）《》]/.test(c) ? 2 : 3);
    const c0 = cls(text[p.offset]) >= 0 ? cls(text[p.offset]) : cls(text[p.offset - 1]);
    let a = p.offset;
    let b = p.offset;
    while (a > 0 && cls(text[a - 1]) === c0 && c0 !== 3) a--;
    while (b < text.length && cls(text[b]) === c0 && c0 !== 3) b++;
    if (c0 === 3) {
      a = Math.max(0, p.offset - 1);
      b = Math.min(text.length, a + 1);
    }
    setSel({ anchor: { ...p, offset: a }, focus: { ...p, offset: b } });
  };

  const dblRef = useRef(onPageDoubleClick);
  dblRef.current = onPageDoubleClick;

  // ---------- 手机：触摸 ----------
  // 手指放在页面上不开始选择（让浏览器滚动）；轻点 = 单击（只在编辑模式放光标、弹键盘），长按 = 选词

  const userName = host().useUserName();
  const setMMode = (m: MobileMode) => {
    setMModeState(m);
    mModeRef.current = m;
    if (m === 'edit') input.current?.focus({ preventScroll: true });
    else input.current?.blur();
    if (m !== 'review') setMFinal(false);
    if (m === 'review') {
      setReviewBar(true);
      // 修订、批注要记作者：第一次审阅时让用户填名字
      const s = host();
      if (!userName && s.setUserName)
        void s.askText('审阅人姓名（修订和批注显示的作者）', '').then((v) => v?.trim() && s.setUserName?.(v.trim()));
    }
  };
  const mobileCancel = () => {
    if (mTouch.current) window.clearTimeout(mTouch.current.timer);
    mTouch.current = null;
  };
  /** 长按（编辑模式）：同桌面的双击——先在按下的位置放光标，再选中光标处的词 */
  const mSelectWordAt = (e: React.MouseEvent<HTMLDivElement>, page: number) => {
    mobileCancel();
    // 编辑和审阅都能长按选词（审阅时选中文字加批注，不弹键盘）
    if (mModeRef.current === 'view' || mFinal) return;
    try {
      onPagePointerDown(e as React.PointerEvent<HTMLDivElement>, page);
    } catch {
      /* 事件已不在派发中，setPointerCapture 可能失败，不影响放光标 */
    }
    dragging.current = false;
    floatDrag.current = null;
    colDrag.current = null;
    // 光标是 React 状态，下一轮渲染后才生效：等它生效再按光标选词（用最新的处理函数）
    const { clientX, clientY } = e;
    window.setTimeout(() => {
      dblRef.current({ clientX, clientY } as React.MouseEvent<HTMLDivElement>, page);
      if (mModeRef.current === 'edit') input.current?.focus({ preventScroll: true });
    }, 0);
  };
  const mobileDown = (e: React.PointerEvent<HTMLDivElement>, page: number) => {
    mobileCancel();
    // 编辑 / 审阅：不让浏览器在抬手后补发 mousedown（会抢走输入框焦点、打断长按选词）；不影响手指滚动
    if (mModeRef.current !== 'view') e.preventDefault();
    const timer = window.setTimeout(() => mTouch.current && mSelectWordAt(e, page), 450);
    mTouch.current = { x: e.clientX, y: e.clientY, t: Date.now(), timer };
  };
  const mobileUp = (e: React.PointerEvent<HTMLDivElement>, page: number) => {
    const t = mTouch.current;
    mobileCancel();
    if (!t || Math.hypot(e.clientX - t.x, e.clientY - t.y) > 10 || Date.now() - t.t > 500) return;
    // 查看模式只滚动；审阅模式点一下放光标（不弹键盘），接受 / 拒绝作用在光标处
    if (mModeRef.current === 'view' || mFinal) return;
    try {
      onPagePointerDown(e, page);
    } catch {
      /* 抬起时 setPointerCapture 可能失败，不影响放光标 */
    }
    dragging.current = false;
    floatDrag.current = null;
    colDrag.current = null;
    if (mModeRef.current === 'edit') input.current?.focus({ preventScroll: true });
  };

  // 按屏宽缩放（转屏 / 键盘弹出时跟着变）；双指缩放后不再自动适应，点“适应宽度”恢复
  const firstWidth = layout.pages[0]?.setup.width ?? 0;
  useLayoutEffect(() => {
    const sc = scroller.current;
    if (!mobile || !mFit || !sc || !firstWidth) return;
    const apply = () => sc.clientWidth && setZoomState(fitWidthZoom(sc.clientWidth, firstWidth * PX_PER_PT, 6, 0.2, 3));
    apply();
    const ro = new ResizeObserver(apply);
    ro.observe(sc);
    return () => ro.disconnect();
  }, [mobile, mFit, firstWidth]);
  usePinchZoom(scroller, pagesEl, zoom, mobile ? { min: 0.2, max: 3, onZoom: (z) => (setMFit(false), setZoomState(z)) } : null);

  // 手机上的页码提示：屏幕上方三分之一处是哪一页
  const [mPage, setMPage] = useState(0);
  useEffect(() => {
    const sc = scroller.current;
    if (!mobile || !sc) return;
    const onScroll = () => {
      const y = sc.scrollTop + sc.clientHeight / 3;
      let p = 0;
      pageEls.current.forEach((el, i) => el && el.offsetTop <= y && (p = i));
      setMPage(p);
    };
    onScroll();
    sc.addEventListener('scroll', onScroll, { passive: true });
    return () => sc.removeEventListener('scroll', onScroll);
  }, [mobile, layout]);

  // 光标滚动到可见
  useEffect(() => {
    const sc = scroller.current;
    const el = caret ? pageEls.current[caret.page] : null;
    if (!sc || !el || !caret || dragging.current) return;
    const top = el.offsetTop + caret.y * ppt;
    const bottom = top + caret.h * ppt;
    if (top < sc.scrollTop + 8) sc.scrollTop = top - 40;
    else if (bottom > sc.scrollTop + sc.clientHeight - 8) sc.scrollTop = bottom - sc.clientHeight + 40;
  }, [caret, ppt]);

  // ---------- 命令 ----------

  /** 保存；返回是否已保存（取消另存为 / 失败为 false） */
  const save = async (asNew: boolean): Promise<boolean> => {
    const s = host();
    let target = tab.path;
    const legacyFrom = legacyDoc.current ? tab.path : undefined;
    // .doc 不能就地保存：第一次保存另存为同名 .docx（重名时加序号，不覆盖已有文件），保存后标签页指向新文件
    if (legacyFrom) asNew = true;
    if (target && !asNew && !origin.current && doc.current.lossy.length) {
      const ok = (await host().confirm(
        `该文档包含本编辑器暂不支持的内容：${doc.current.lossy.join('、')}。\n覆盖保存会丢失这些内容。\n\n确定覆盖原文件吗？选择“取消”将另存为新文件。`,
      ));
      if (!ok) asNew = true;
    }
    if (!target || asNew) {
      const dir = tab.path ? dirname(tab.path) : undefined;
      let base = tab.path ? basename(tab.path).replace(/\.(docx|doc|dot)$/i, '') + (legacyFrom ? '' : '_副本') : '新建文档';
      if (legacyFrom) {
        const taken = new Set((await s.guard(s.listDir(dir ?? s.workFolders()[0] ?? '')))?.map((e) => e.name.toLowerCase()) ?? []);
        const stem = base;
        for (let n = 1; taken.has(`${base}.docx`.toLowerCase()); n++) base = `${stem} (${n})`;
      }
      const picked = await pickSave({ title: '另存为', name: `${base}.docx`, dir, label: 'Word 文档', exts: ['docx'] });
      if (!picked) return false;
      target = picked;
    }
    // 打开的文件：只重写改过的部分，其余逐字节保留；新建文档整篇写出
    const bytes = origin.current ? await patchDocx(doc.current, origin.current) : writeDocx(doc.current);
    if ((await s.guard(host().writeBytes(target, bytes).then(() => true))) === undefined) return false;
    doc.current.lossy = [];
    setLossy([]);
    setSavedVersion(version);
    void clearDraft(target);
    const unsaved = latestVersion.current !== version;
    s.updateTab(tab.id, { path: target, title: basename(target), dirty: unsaved });
    legacyDoc.current = false;
    s.toast(`已保存 ${basename(target)}`);
    return !unsaved;
  };

  const saveRef = useRef(save);
  saveRef.current = save;
  // 关闭有未保存修改的标签时，工作台通过它保存（新文档弹“另存为”）
  useEffect(() => host().registerSave?.(tab.id, () => saveRef.current(false)), [tab.id]);

  const print = async () => {
    host().toast('正在准备打印…');
    const first = layout.pages[0]?.setup ?? doc.current.setup;
    await host().printImages(renderPagesToImages(layout, 2), `${first.width}pt ${first.height}pt`);
  };

  /** 导出 PDF：按当前排版逐页画成图片（同打印），存到文档所在的工作文件夹 */
  const exportPdf = async () => {
    const s = host();
    const stem = tab.path ? basename(tab.path).replace(/\.docx$/i, '') : '新建文档';
    const target = await pickSave({ title: '导出 PDF', name: `${stem}.pdf`, dir: tab.path ? dirname(tab.path) : undefined, label: 'PDF 文档', exts: ['pdf'] });
    if (!target) return refocus();
    s.toast('正在导出 PDF…');
    const urls = renderPagesToImages(layout, 2);
    const bytes = await pagesToPdf(layout.pages.map((pg, i) => ({ url: urls[i], w: pg.setup.width, h: pg.setup.height })));
    if ((await s.guard(host().writeBytes(target, bytes).then(() => true))) === undefined) return;
    s.toast(`已导出 ${basename(target)}`);
    refocus();
  };

  const insertTableCmd = async () => {
    const v = await host().askText('插入表格：行数×列数', '3×4');
    const m = v?.match(/(\d+)\D+(\d+)/);
    if (!m) return refocus();
    const rows = Math.min(200, Math.max(1, Number(m[1])));
    const cols = Math.min(30, Math.max(1, Number(m[2])));
    mutate('table', (d, s) => {
      if (s.focus.story !== 'body' || s.focus.path.length !== 1) {
        host().toast('请把光标放在正文段落中（不能在表格内插入表格）');
        return null;
      }
      const p = insertTable(d, s.focus, rows, cols);
      return p ? caretSel(p) : null;
    });
  };

  /** 目录：没有时在光标处插入，有时按当前标题和页码更新；插入后页码会变，再排一次版更新页码 */
  const tocCmd = () => {
    const had = !!tocRange(doc.current);
    const ok = mutate('toc', (d, s) => {
      if (had) {
        updateToc(d, blockPages(layoutDocument(d, getMeasurer())));
        updateToc(d, blockPages(layoutDocument(d, getMeasurer())));
        return s;
      }
      const p = insertToc(d, s.focus, blockPages(layoutDocument(d, getMeasurer())));
      if (!p) return null;
      updateToc(d, blockPages(layoutDocument(d, getMeasurer())));
      return caretSel(p);
    });
    const st = host();
    if (!ok && !had) st.toast(sel.focus.story !== 'body' || sel.focus.path.length !== 1 ? '请把光标放在正文段落里再插入目录' : '文档里没有标题：先把章节标题设为“标题 1 / 2 / 3”样式');
    else if (had) st.toast('目录已更新');
  };
  const curStyle = paraStyleOf(getPara(doc.current, sel.focus));

  /** 合并：选中的单元格矩形 */
  const mergeCmd = () =>
    mutate('table', (d, s) => {
      const p = mergeTableCells(d, s.anchor, s.focus);
      if (!p) host().toast(cellBlock ? '选中的区域会切开其它合并单元格，不能合并' : '请先拖动选中要合并的几个单元格');
      return p ? caretSel(p) : null;
    });
  const splitCmd = async () => {
    const v = await host().askText('拆分成几列', '2');
    if (v === null) return refocus();
    const n = Math.floor(Number(v));
    if (!(n >= 2 && n <= 20)) return host().toast('请输入 2 到 20 之间的列数');
    tableCmd((d, p) => splitCell(d, p, n), '已合并的单元格请先取消合并再拆分');
  };
  /** 底纹 / 框线：作用于选中的单元格（没有跨格选择时为当前格） */
  const cellFormatCmd = (patch: Parameters<typeof setCellFormat>[3]) => mutate('table', (d, s) => void setCellFormat(d, s.anchor, s.focus, patch));
  /** 给选中的格子设框线（line 缺省用当前的笔） */
  const bordersCmd = (target: BorderTarget, line: CellEdge = pen) =>
    mutate('table', (d, s) => (applyCellBorders(d, s.anchor, s.focus, target, line) ? undefined : (host().toast('请把光标放在表格里，或在同一张表格里选中单元格'), null)));

  // ---------- 右键菜单 ----------

  /** 菜单里的剪切 / 复制走浏览器的 copy / cut 事件（和 Ctrl+C / X 同一套处理）；粘贴从系统剪贴板读 */
  const clipCmd = async (kind: 'copy' | 'cut' | 'paste' | 'plain') => {
    if (canFocus()) input.current?.focus({ preventScroll: true });
    if (kind === 'copy' || kind === 'cut') {
      if (!collapsed(sel) && !document.execCommand(kind)) host().toast(`请按 Ctrl+${kind === 'copy' ? 'C' : 'X'}`);
      return;
    }
    try {
      let text = '';
      let html = '';
      let img: Blob | null = null;
      for (const it of await navigator.clipboard.read()) {
        if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
        if (kind === 'paste' && it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
        const t = it.types.find((x) => x.startsWith('image/'));
        if (t && kind === 'paste') img = await it.getType(t);
      }
      if (img && !text) return void insertImageFile(img);
      if (text || html) pasteFrom(text, html);
    } catch {
      host().toast(`没能读取剪贴板，请按 ${kind === 'plain' ? 'Ctrl+Shift+V' : 'Ctrl+V'} 粘贴`);
    }
  };
  /** 右键：点在选区外时先把光标移过去（同 Word），再弹出菜单 */
  const onPageContextMenu = (e: React.MouseEvent<HTMLDivElement>, page: number) => {
    e.preventDefault();
    const pt = pointAt(e, page);
    const region = regionOf(pt.y, page, pt.x);
    const hit = hitTest(layout, page, pt.x, pt.y, story === 'body' ? 'body' : region === 'body' ? story : region);
    const [a, b] = ordered(sel);
    const inside = hit && !collapsed(sel) && hit.story === a.story && comparePos(a, hit) <= 0 && comparePos(hit, b) <= 0;
    if (hit && !inside && !floatSel) {
      setPageHint(page);
      setPending(null);
      setSel(caretSel(hit));
    }
    if (canFocus()) input.current?.focus({ preventScroll: true });
    setMenu({ x: e.clientX, y: e.clientY });
  };
  // ---------- 域：更新域（F9）、编辑域、切换域代码（Shift+F9，Alt+F9 全部） ----------

  /** 计算域用的环境：现在的时间、文件名、页数（书签由 updateFields 从文档里取） */
  const fieldEnv = (): FieldEnv => ({ now: new Date(), fileName: tab.path ? basename(tab.path) : tab.title, filePath: tab.path ?? undefined, numPages: layout.pages.length });
  /** 光标所在的域 */
  const fieldHere = () => {
    try {
      return fieldAt(getPara(doc.current, sel.focus), sel.focus.offset);
    } catch {
      return null;
    }
  };
  const fieldReport = (r: UpdateResult) => {
    const msg = [r.updated ? `已更新 ${r.updated} 个域` : '', r.kept.length ? `${[...new Set(r.kept)].join('、')} 域暂时算不出结果，保留了原结果` : '', r.locked ? `${r.locked} 个域已锁定，没有更新` : '']
      .filter(Boolean)
      .join('；');
    host().toast(msg || '域的结果没有变化');
  };
  const updateFieldCmd = (all = false) => {
    const [a, b] = ordered(sel);
    const ids = all ? allFieldIds(doc.current) : fieldIdsInRange(doc.current, a, b);
    if (!ids.length) return host().toast(all ? '文档里没有可以更新的域' : '光标所在处没有域');
    const env = fieldEnv();
    // 先在副本上算一遍：结果都没变就不改文档（不算修改、不进撤销）
    const dry = updateFields(structuredClone(doc.current), env, new Set(ids));
    if (dry.updated) mutate('field', (d) => void updateFields(d, env, new Set(ids)));
    fieldReport(dry);
  };
  const toggleFieldCodes = (all: boolean) => {
    const [a, b] = ordered(sel);
    const ids = all ? allFieldIds(doc.current) : fieldIdsInRange(doc.current, a, b);
    if (!ids.length) return;
    setShowCode(ids, !ids.some(showsCode));
    setViewTick((t) => t + 1);
  };
  const editFieldCmd = () => {
    const f = fieldHere();
    if (f) setFieldEdit({ id: f.id, instr: f.instr });
  };
  const applyFieldEdit = (instr: string) => {
    const target = fieldEdit;
    setFieldEdit(null);
    if (!target) return;
    const env = fieldEnv();
    let res: UpdateResult | null = null;
    mutate('field', (d) => {
      const hit = findField(d, target.id);
      if (!hit) return null;
      setFieldInstr(hit.p, target.id, instr);
      res = updateFields(d, env, new Set([target.id]));
    });
    const r = res as UpdateResult | null;
    if (r?.kept.length) host().toast('域代码已修改；这个域编辑器暂时算不出结果，保留了原结果');
  };

  const menuItems = (): MenuItem[] => {
    const has = !collapsed(sel);
    const fld = fieldHere();
    const items: MenuItem[] = [
      ...(selectedImage || floatSel?.image ? [{ label: '设置图片格式…', onClick: openPictureDialog }] : []),
      { label: '剪切', keys: 'Ctrl+X', disabled: !has, onClick: () => void clipCmd('cut') },
      { label: '复制', keys: 'Ctrl+C', disabled: !has, onClick: () => void clipCmd('copy') },
      { label: '粘贴', keys: 'Ctrl+V', onClick: () => void clipCmd('paste') },
      { label: '只粘贴文本', keys: 'Ctrl+Shift+V', onClick: () => void clipCmd('plain') },
      ...(fld
        ? [
            { label: '更新域', keys: 'F9', sep: true, onClick: () => updateFieldCmd() },
            { label: '编辑域…', onClick: editFieldCmd },
            { label: '切换域代码', keys: 'Shift+F9', onClick: () => toggleFieldCodes(false) },
          ]
        : []),
      { label: '字体…', keys: 'Ctrl+D', sep: true, onClick: () => setDialog('font') },
      { label: '段落…', onClick: () => setDialog('para') },
      { label: '清除格式', onClick: () => clearFormatCmd(true) },
    ];
    if (inTable) {
      items.push(
        {
          label: '插入',
          sep: true,
          sub: [
            { label: '在上方插入行', onClick: () => tableCmd((d, p) => insertRow(d, p, false)) },
            { label: '在下方插入行', onClick: () => tableCmd((d, p) => insertRow(d, p, true)) },
            { label: '在左侧插入列', onClick: () => tableCmd((d, p) => insertColumn(d, p, true)) },
            { label: '在右侧插入列', onClick: () => tableCmd((d, p) => insertColumn(d, p)) },
          ],
        },
        {
          label: '删除',
          sub: [
            { label: '删除行', onClick: () => tableCmd(deleteRow) },
            { label: '删除列', onClick: () => tableCmd(deleteColumn) },
            { label: '删除表格', onClick: () => mutate('table', (d) => caretSel(deleteTable(d, inTable.bi))) },
          ],
        },
        { label: '合并单元格', disabled: !has, onClick: mergeCmd },
        { label: '拆分单元格…', onClick: () => void splitCmd() },
        {
          label: '单元格对齐方式',
          sub: [
            { label: '靠上', onClick: () => cellFormatCmd({ vAlign: 'top' }) },
            { label: '垂直居中', onClick: () => cellFormatCmd({ vAlign: 'center' }) },
            { label: '靠下', onClick: () => cellFormatCmd({ vAlign: 'bottom' }) },
            { label: '水平垂直居中', sep: true, onClick: () => cellFormatCmd({ vAlign: 'center', align: 'center' }) },
          ],
        },
        { label: '边框…', onClick: () => setDialog('borders') },
      );
    }
    items.push(
      { label: '超链接…', keys: 'Ctrl+K', sep: true, onClick: openLinkDialog },
      { label: '新建批注', disabled: !has || sel.anchor.story !== 'body' || sel.focus.story !== 'body', onClick: () => void newComment() },
    );
    return items;
  };
  const colWidthCmd = async () => {
    const t = inTable;
    if (!t) return;
    const starts = (() => {
      let g = 0;
      return t.table.rows[t.r].cells.map((c) => ((g += c.gridSpan), g - c.gridSpan));
    })();
    const g0 = starts[t.c];
    const g1 = g0 + t.table.rows[t.r].cells[t.c].gridSpan - 1;
    const cur = t.table.cols.slice(g0, g1 + 1).reduce((a, b) => a + b, 0);
    const v = await host().askText('列宽（厘米）', String(Math.round((cur / PT_PER_CM) * 100) / 100));
    if (v === null) return refocus();
    const n = Number(v) * PT_PER_CM;
    if (!(n > 0.3)) return;
    // 当前格的右边线移动到新宽度处（右边一列此消彼长）
    mutate('table', (d) => void moveColumnLine(d.body[t.bi] as Table, g1, n - cur));
  };

  const tableCmd = (fn: (d: Doc, p: Pos) => Pos | null, fail = '无法完成此表格操作') =>
    mutate('table', (d, s) => {
      const p = fn(d, s.focus);
      if (!p) host().toast(fail);
      return p ? caretSel(p) : null;
    });

  const rowHeightCmd = async () => {
    const t = inTable;
    if (!t) return;
    const row = t.table.rows[t.r];
    const v = await host().askText('行高（磅），填 0 为自动；前加“=”表示固定值', row.heightRule === 'exact' ? `=${row.height}` : String(row.height));
    if (v === null) return refocus();
    const exact = v.trim().startsWith('=');
    const n = Number(v.replace('=', ''));
    if (!Number.isFinite(n) || n < 0) return;
    mutate('table', (d) => {
      const r = (d.body[t.bi] as typeof t.table).rows[t.r];
      r.height = n;
      r.heightRule = n === 0 ? 'auto' : exact ? 'exact' : 'atLeast';
    });
  };

  /** 页码写入光标所在页那一节启用的页脚分区（默认、首页、偶数页）；与上一节相同的节一起改变 */
  const pageNumberCmd = (style: string) => {
    mutate('field', (d) => {
      const { owner, footer } = layout.pages[caretPage].hf;
      const at = hfPath(owner, 0);
      const titlePg = sectionList(d).find((x) => x.owner === owner)?.hf.titlePg ?? d.setup.titlePg;
      const stories = (['footer', 'footerFirst', 'footerEven'] as const).filter((k) => k === 'footer' || (k === 'footerFirst' ? titlePg : d.setup.evenAndOdd));
      for (const k of stories) {
        const f = hfList(d, k, at)[0];
        f.props = { ...f.props, align: 'center' };
        const props = { ...(f.runs[0]?.props ?? f.mark) };
        const P = { text: FIELD_CHAR, props, field: 'PAGE' as const };
        const N = { text: FIELD_CHAR, props, field: 'NUMPAGES' as const };
        const t = (text: string) => ({ text, props });
        f.runs = style === 'dash' ? [t('- '), P, t(' -')] : style === 'plain' ? [P] : [t('第 '), P, t(' 页 共 '), N, t(' 页')];
      }
      return caretSel({ story: footer, path: at, offset: paraLength(hfList(d, footer, at)[0]) });
    });
  };

  const breakCmd = (v: string) => {
    if (v === 'page') return typeText('\f');
    if (v === 'column') return typeText('\v');
    mutate('section', (d, s) => {
      const at = clearSelection(d, s);
      const p = at && insertSectionBreak(d, at, v as SectionBreak);
      if (!p) host().toast('分节符只能插入在正文段落中（不能在表格或页眉页脚内）');
      return p ? caretSel(p) : null;
    });
  };

  const numberingCmd = (preset: string) => {
    const [a, b] = ordered(sel);
    mutate('para', (d) => applyNumbering(d, a, b, preset || null));
  };
  const levelCmd = (delta: 1 | -1) => {
    const [a, b] = ordered(sel);
    mutate('para', (d) => (changeNumLevel(d, a, b, delta) ? undefined : null));
  };
  const curPreset = curPara.num ? (doc.current.lists.find((l) => l.id === curPara.num!.id)?.preset ?? '') : '';

  // 多节文档中页眉页脚的“与上一节相同”（第一节为 null）
  const sectionCount = useMemo(() => sectionList(doc.current).length, [version]);
  const hfLinkedNow = story !== 'body' && sectionCount > 1 ? hfLinked(doc.current, story as HFStory, sel.focus.path) : null;
  const toggleHFLink = () => {
    if (story === 'body' || hfLinkedNow === null) return;
    mutate('hf-link', (d) => caretSel(setHFLinked(d, story as HFStory, sel.focus.path, !hfLinkedNow)));
  };

  /** 光标所在的节：正文按段落所在的节，页眉页脚按所在页的节 */
  const caretSection = (): SectionInfo => {
    const list = sectionList(doc.current);
    if (story === 'body') return list.find((x) => x.index === null || x.index >= sel.focus.path[0]) ?? list[list.length - 1];
    return list[layout.pages[caretPage]?.section ?? 0] ?? list[list.length - 1];
  };

  const insertFieldCmd = (kind: 'PAGE' | 'NUMPAGES') =>
    mutate('field', (d, s) => {
      const at = clearSelection(d, s);
      return at ? caretSel(insertField(d, at, kind)) : null;
    });

  const toggleHeaderFooter = () => {
    if (story === 'body') {
      const { header, owner } = layout.pages[caretPage].hf;
      const path = hfPath(owner, 0);
      moveTo({ story: header, path, offset: paraLength(hfList(doc.current, header, path)[0]) }, false, caretPage);
    }
    else moveTo(storyEdge(doc.current, 'body', false), false);
    refocus();
  };

  useEffect(() => {
    if (loaded) refocus();
  }, [loaded]);

  // ---------- 界面 ----------

  // 标尺上的段落：只有正文里的普通段落有缩进标记（表格单元格、页眉页脚里只显示页边距）
  const rulerPage = layout.pages[caretPage] ?? layout.pages[0];
  const rulerHit = story === 'body' && sel.focus.path.length === 1 ? findLine(layout, sel.focus, pageHint) : null;
  const rulerArea = rulerPage && rulerHit ? paragraphArea(rulerPage.setup, rulerHit.pl.x) : null;

  const lineIdx = LINE_OPTIONS.findIndex((o) => o.value.rule === curPara.line.rule && Math.abs(o.value.value - curPara.line.value) < 0.01);
  // 光标所在节的页面设置（状态栏的每页行数等）
  const setup = layout.pages[caretPage]?.setup ?? doc.current.setup;
  const rowEndBox = rowEnd ? rowEnds(layout.pages[rowEnd.page] ?? layout.pages[0]).find((e) => e.bi === rowEnd.bi && e.r === rowEnd.r) : undefined;
  const caretStyle: React.CSSProperties | undefined =
    rowEnd && rowEndBox && pageEls.current[rowEnd.page]
      ? {
          left: pageEls.current[rowEnd.page]!.offsetLeft + (rowEndBox.x + 2) * ppt,
          top: pageEls.current[rowEnd.page]!.offsetTop + (rowEndBox.baseline - rowEndBox.lineH * 0.8) * ppt,
          height: rowEndBox.lineH * 0.95 * ppt,
        }
      : caret && pageEls.current[caret.page]
      ? {
          left: pageEls.current[caret.page]!.offsetLeft + caret.x * ppt,
          top: pageEls.current[caret.page]!.offsetTop + caret.y * ppt,
          height: caret.h * ppt,
        }
      : undefined;

  // ---------- 链接、书签、水印、样式、模板 ----------

  const styleCat = (): StyleCatalog => (origin.current ? origin.current.styles : catalogFromDoc(doc.current));
  const selText = () => {
    const [a, b] = ordered(sel);
    return a.story === b.story && !collapsed(sel) ? selectedText(doc.current, a, b).replace(/[\n\r\f\v]+/g, ' ').replaceAll(FIELD_CHAR, '') : '';
  };
  /** 光标所在的链接（光标在链接里、没有选中文字时才算） */
  const linkHere = () => (collapsed(sel) ? linkAt(doc.current, sel.focus) : null);
  const openLinkDialog = () => {
    if (!collapsed(sel) && sel.anchor.story !== sel.focus.story) return host().toast('选区跨越了正文和页眉页脚，请分开处理');
    setDialog('link');
  };
  const goBookmark = (name: string) => {
    const b = listBookmarks(doc.current, true).find((x) => x.name === name);
    if (!b) return host().toast(`文档里没有书签“${name}”`);
    setSel(caretSel(b.start));
    refocus();
  };
  /** 点链接：网页 / 邮箱用系统程序打开（只放行 http、https、mailto、tel），文档内链接跳到书签 */
  const followLink = (l: LinkTarget) => {
    if (l.anchor) return goBookmark(l.anchor);
    const url = l.url ?? '';
    if (!/^(https?:|mailto:|tel:)/i.test(url)) return host().toast(url ? `不打开这种地址：${url}` : '这个链接的地址没有记录在文档里');
    void host().guard(host().openUrl(url));
  };
  const applyLink = (r: { text: string; link: LinkTarget }) => {
    const hit = linkHere();
    setDialog(null);
    mutate('link', (d, s) => {
      let [a, b] = hit ? [hit.from, hit.to] : ordered(s);
      if (a.story !== b.story) return null;
      if (comparePos(a, b) === 0) b = insertRaw(d, a, r.text);
      else if (hit && r.text !== hit.text) {
        a = deleteRange(d, a, b) ?? a;
        b = insertRaw(d, a, r.text);
      }
      setLink(d, a, b, r.link);
      return caretSel(b);
    });
  };
  const removeLinkCmd = () => {
    const hit = linkHere();
    setDialog(null);
    mutate('link', (d, s) => {
      const [a, b] = hit ? [hit.from, hit.to] : ordered(s);
      clearLink(d, a, b);
    });
  };
  const addBookmarkCmd = (name: string) => {
    const [a, b] = ordered(sel);
    mutate('bookmark', (d) => {
      addBookmark(d, a, b, name);
    });
  };
  const applyWatermark = (w: Watermark | null) => {
    setDialog(null);
    mutate('watermark', (d) => {
      if (w) d.watermark = w;
      else delete d.watermark;
    });
  };
  const applyStyleCmd = (id: string, type: 'paragraph' | 'character') => {
    const [a, b] = ordered(sel);
    mutate('style', (d) => {
      if (type === 'paragraph') applyParagraphStyle(d, styleCat(), a, b, id);
      else applyCharacterStyle(d, styleCat(), a, b, id);
    });
  };
  const createStyleCmd = (name: string, type: 'paragraph' | 'character'): string | null => {
    const [a, b] = ordered(sel);
    const bad = checkStyleName(styleCat(), doc.current, name);
    if (bad) return bad;
    mutate('style', (d) => {
      createStyle(d, styleCat(), a, name, type, b);
    });
    return null;
  };
  const startProtect = async (mode: ProtectMode, password: string): Promise<string | null> => {
    const p = await makeProtection(mode, password || undefined);
    // 设置 / 停止保护不进撤销（撤销不能把保护撤掉）
    history.current.clear();
    doc.current.protection = p;
    bump();
    host().toast(`已启动保护（${protectLabel(p)}）`);
    return null;
  };
  const stopProtect = async (password: string): Promise<string | null> => {
    const p = doc.current.protection;
    if (p?.hash && !(await checkPassword(p, password))) return '密码不对';
    history.current.clear();
    delete doc.current.protection;
    bump();
    host().toast('已停止保护');
    return null;
  };
  const renameStyleCmd = (id: string, name: string): string | null => {
    try {
      renameStyle(structuredClone(doc.current), styleCat(), id, name);
    } catch (e) {
      return e instanceof Error ? e.message : String(e);
    }
    mutate('style', (d) => void renameStyle(d, styleCat(), id, name));
    return null;
  };
  const deleteStyleCmd = (id: string): string | null => {
    try {
      deleteStyle(structuredClone(doc.current), styleCat(), id);
    } catch (e) {
      return e instanceof Error ? e.message : String(e);
    }
    mutate('style', (d) => void deleteStyle(d, styleCat(), id));
    return null;
  };
  const redefineStyleCmd = (id: string) => {
    const [, b] = ordered(sel);
    mutate('style', (d) => {
      redefineStyle(d, styleCat(), id, b);
    });
  };
  /** 用一条数据生成填好的文档：在文档的复制上填值，按原文件就地修补保存（没有原文件时整篇写出） */
  const saveBytes = (d: Doc) => (origin.current ? patchDocx(d, origin.current) : Promise.resolve(writeDocx(d)));
  const generateMerged = async (records: Record<string, string>[], start: number) => {
    const copy = structuredClone(doc.current);
    const r = fillRecords(copy, records, start);
    return { bytes: r.skipped ? null : await saveBytes(copy), consumed: r.consumed };
  };
  const generateCombined = async (records: Record<string, string>[]) => {
    const r = combineRecords(doc.current, records);
    return { bytes: await saveBytes(r.doc), count: r.count, skipped: r.skipped };
  };
  const insertFieldCmd2 = (name: string, asWordField: boolean) => {
    const [a, b] = ordered(sel);
    mutate('field', (d) => {
      const at = a.story === b.story && comparePos(a, b) !== 0 ? (deleteRange(d, a, b) ?? a) : sel.focus;
      return caretSel(asWordField ? insertMergeField(d, at, name) : insertRaw(d, at, `{{${name}}}`));
    });
  };
  const tmplFields = dialog === 'merge' ? placeholders(doc.current) : [];
  const palette = tab.path ? (host().fieldPalette?.(tab.path) ?? null) : null;

  /**
   * “插入字段”面板的“AI 生成”：AI 按工作台的技能文件判断合同里哪些文字是数据、该换成哪个字段，
   * 编辑器先把文档改成审阅状态（修订打开，存进文件），再把建议以修订方式改进文档（一步撤销），等用户审阅后接受或拒绝。
   */
  const aiTemplate = async () => {
    if (!ai?.templateFields || !palette) return;
    if (tplBusy) return tplCtrl.current?.abort();
    if (activeProtection(doc.current)) return host().toast('文档受保护，AI 不能修改；要修改请先“停止保护”');
    const before = templateParagraphs(doc.current);
    if (!before.texts.length) return host().toast('文档里没有文字');
    const fields = paletteFields(palette);
    setTrackOn(true);
    setReviewBar(true);
    setTplReport(null);
    const ctrl = new AbortController();
    tplCtrl.current = ctrl;
    setTplBusy(true);
    let res;
    try {
      res = await ai.templateFields({ title: tab.title, paragraphs: before.texts, fields }, { signal: ctrl.signal });
    } catch (e) {
      if (!ctrl.signal.aborted) host().toast(`AI 生成失败：${e instanceof Error ? e.message : String(e)}`);
      return;
    } finally {
      tplCtrl.current = null;
      setTplBusy(false);
    }
    // 等 AI 的时候文档又被改过：段落序号对不上了，不改
    const now = templateParagraphs(doc.current);
    if (JSON.stringify(now.texts) !== JSON.stringify(before.texts)) return host().toast('等待 AI 期间文档被修改了，请重新点“AI 生成”');
    if (activeProtection(doc.current)) return host().toast('文档受保护，AI 不能修改');
    history.current.record(doc.current, live.current.sel, 'ai');
    tracking.on = true;
    tracking.author = author;
    let r;
    try {
      r = applyTemplateHits(doc.current, now.refs, res.hits, new Set(fields.map((f) => f.key)));
    } catch (e) {
      const snap = history.current.dropLast();
      if (snap) doc.current = snap.doc;
      bump();
      return host().toast(`修改失败：${e instanceof Error ? e.message : String(e)}`);
    }
    if (!r.applied) history.current.dropLast();
    else {
      setFloatSel(null);
      setSel(caretSel(storyEdge(doc.current, 'body', false)));
      bump();
    }
    const lines = [r.applied ? `已把 ${r.applied} 处数据换成字段，都记为修订：请在“审阅”里逐条接受或拒绝，保存后仍是审阅状态。` : 'AI 没有找到可以换成字段的数据。'];
    if (res.notes) lines.push(res.notes);
    if (r.skipped.length) lines.push(`没有处理 ${r.skipped.length} 条：`, ...r.skipped.slice(0, 12).map((x) => `· ${x.hit.text ?? x.hit.after ?? ''} → ${x.hit.field}（${x.reason}）`));
    setTplReport(lines.join('\n'));
  };

  const alignBtn = (a: Align, label: string, hint: string) => (
    <button className={`icon-btn${curPara.align === a ? ' on' : ''}`} aria-pressed={curPara.align === a} title={hint} aria-label={label} onClick={() => setPara({ align: a })}>
      <AlignIcon a={a} />
    </button>
  );

  // ---------- 手机：编辑命令（都是上面现成的命令） ----------
  const mSelectOn = mobile && mMode === 'edit' && mSelect;
  const mTable = (op: TableOp) =>
    tableCmd((d, p) =>
      op === 'rowAbove' ? insertRow(d, p, false) : op === 'rowBelow' ? insertRow(d, p, true) : op === 'colLeft' ? insertColumn(d, p, true) : op === 'colRight' ? insertColumn(d, p) : op === 'delRow' ? deleteRow(d, p) : deleteColumn(d, p),
    );
  const mInsertImage = async () => {
    const s = host();
    const paths = (await s.guard(s.pickFiles?.('插入图片', ['png', 'jpg', 'jpeg', 'gif', 'bmp']) ?? Promise.resolve([]))) ?? [];
    if (!paths[0]) return;
    const bytes = await s.guard(s.readBytes(paths[0]));
    if (bytes) await insertImageFile(new Blob([bytes as BlobPart]));
  };
  const mRev: MobileReviewApi = {
    trackOn,
    setTrack: (on) => (activeProtection(doc.current)?.edit === 'trackedChanges' ? host().toast('文档受保护（只能修订），修订不能关闭') : setTrackOn(on)),
    prev: () => gotoRevision(true),
    next: () => gotoRevision(false),
    accept: () => resolveCmd(true),
    reject: () => resolveCmd(false),
    acceptAll: () => resolveCmd(true, true),
    rejectAll: () => resolveCmd(false, true),
    comments: doc.current.comments?.filter((c) => !c.parent).length ?? 0,
    commentsOpen: mComments,
    toggleComments: () => setMComments(!mComments),
    newComment: () => void newComment(),
    final: mFinal,
    setFinal: (on) => {
      setSel(caretSel(storyEdge(doc.current, 'body', false)));
      setMComments(false);
      setMFinal(on);
    },
  };
  const mEd: MobileEditApi = {
    run: curRun,
    para: curPara,
    preset: curPreset,
    style: curStyle ?? 'Normal',
    inTable: !!inTable,
    canUndo: history.current.canUndo,
    canRedo: history.current.canRedo,
    hasSelection: !collapsed(sel),
    selecting: mSelect,
    setSelecting: setMSelect,
    setRun,
    setPara,
    grow: (dir) => growFont(dir, true),
    numbering: numberingCmd,
    paraStyle: (id) => mutate('style', (d, s) => void applyParaStyle(d, s.anchor, s.focus, id as ParaStyle)),
    table: mTable,
    undo,
    redo,
    insertImage: () => void mInsertImage(),
    clip: (k) => void clipCmd(k),
  };

  return (
    <div className={`word${mobile ? ' mobile' : ''}${mSelectOn ? ' selecting' : ''}`}>
      {!mobile && (<>
      <div className="word-tabbar">
        <div className="word-tabs" role="tablist">
          {(['home', 'page', 'tools'] as const).map((t) => (
            <button key={t} role="tab" aria-selected={barTab === t} className={barTab === t ? 'on' : ''} onClick={() => setBarTab(t)}>
              {t === 'home' ? '常用' : t === 'page' ? '页面' : '链接与模板'}
            </button>
          ))}
        </div>
        <div className="word-tabbar-quick">
          {host().newDocument && (
            <button className="icon-btn raised" title="新建 Word 文档" aria-label="新建" onClick={() => host().newDocument?.('word')}>
              <Icon name="newDocColor" size={20} />
            </button>
          )}
          {host().openDocument && (
            <button className="icon-btn raised" title="打开文件…" aria-label="打开" onClick={() => void host().openDocument?.('word')}>
              <Icon name="openColor" size={20} />
            </button>
          )}
        <div className="split-btn" ref={saveMenuRef}>
          <button className="icon-btn primary" onClick={() => void save(false)} title="保存（Ctrl+S）" aria-label="保存">
            <Icon name="saveColor" size={20} />
          </button>
          <button className="icon-btn split-arrow" aria-label="保存选项" aria-haspopup="menu" aria-expanded={saveMenu} onClick={() => setSaveMenu(!saveMenu)}>
            <Icon name="chevron" size={16} className="split-chev" />
          </button>
          {saveMenu && (
            <div className="split-menu" role="menu">
              <button role="menuitem" onClick={() => { setSaveMenu(false); void save(false); }}>保存</button>
              <button role="menuitem" onClick={() => { setSaveMenu(false); void save(true); }}>另存为</button>
              <button role="menuitem" title="按当前排版导出 PDF" onClick={() => { setSaveMenu(false); void exportPdf(); }}>另存为 PDF</button>
            </div>
          )}
        </div>
        <button className="icon-btn" title="撤销（Ctrl+Z）" aria-label="撤销" disabled={!history.current.canUndo} onMouseDown={(e) => e.preventDefault()} onClick={undo}>
          <Icon name="undo" size={18} />
        </button>
        <button className="icon-btn" title="重做（Ctrl+Y）" aria-label="重做" disabled={!history.current.canRedo} onMouseDown={(e) => e.preventDefault()} onClick={redo}>
          <Icon name="redo" size={18} />
        </button>
        <button
          className={`icon-btn${painter ? ' on' : ''}`}
          aria-label="格式刷"
          aria-pressed={!!painter}
          title="格式刷：单击用一次，双击可连续刷（Esc 结束）；选中整段时连段落格式一起刷"
          onClick={() => pickFormat(false)}
          onDoubleClick={() => {
            setPainter(null);
            pickFormat(true);
          }}
        >
          <Icon name="brush" size={20} />
        </button>
        </div>
        <span className="grow" />
        <div className="word-tabbar-right">
        <div className="word-search" title="查找（Ctrl+F）；Enter 开始查找，替换请按 Ctrl+H">
          <input
            aria-label="搜索文档"
            placeholder="搜索文档"
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && runSearch()}
          />
          <button className="icon-btn" aria-label="查找" onClick={runSearch}>
            <Icon name="search" size={14} />
          </button>
        </div>
        <button className={`btn${reviewBar ? ' soft' : ''}`} title="修订和批注" onClick={() => setReviewBar(!reviewBar)}>
          审阅{trackOn ? '（修订中）' : ''}
        </button>
        <button className="icon-btn" title="设置：标尺、页面设置等" aria-label="设置" onClick={() => setSettingsOpen(true)}>
          <Icon name="gear" />
        </button>
        <button className="icon-btn" title="打印（Ctrl+P）" aria-label="打印" onClick={() => void print()}>
          <Icon name="printer" />
        </button>
        <FloatViewButton tabId={tab.id} path={tab.path} />
        </div>
      </div>
      <div className="toolbar word-bar">
        {barTab === 'home' ? (
          <>
        {ai && aiActions.length > 0 && (
          <select
            className="ed-select"
            style={{ width: 64 }}
            aria-label="AI"
            title="AI：对选中的文字润色、改写、翻译、总结（结果先预览，确认后再写入）"
            value=""
            onChange={(e) => void runAi(e.target.value)}
          >
            <option value="" disabled>
              AI
            </option>
            {aiActions.map((x) => (
              <option key={x.id} value={x.id} title={x.hint}>
                {x.label}
              </option>
            ))}
          </select>
        )}
        {ai?.openAssistant && (
          <button className="btn" title="AI 助手：用对话让 AI 阅读和修改这份文档（每次修改都可以撤销）" onClick={openAssistant}>
            AI 助手
          </button>
        )}
        <span className="sep" />
        <select
          className="ed-select"
          style={{ width: 78 }}
          aria-label="样式"
          title="段落样式（标题会出现在目录和 Word 的导航窗格里）"
          value={curStyle ?? ''}
          onChange={(e) => {
            const v = e.target.value as ParaStyle;
            if (v) mutate('style', (d, s) => void applyParaStyle(d, s.anchor, s.focus, v));
          }}
        >
          {!curStyle && <option value="">其他</option>}
          {PARA_STYLES.map((s) => (
            <option key={s.id} value={s.id}>
              {s.label}
            </option>
          ))}
        </select>
        <select
          className="ed-select"
          style={{ width: 150 }}
          aria-label="字体"
          title="字体：中文字体作用于汉字，西文字体作用于英文和数字"
          value={fontShown}
          onChange={(e) => {
            const f = e.target.value;
            const latin = EDITOR_FONTS_LATIN.includes(f) || (!EDITOR_FONTS_EA.includes(f) && f === curRun.fontLatin);
            setFontPick(latin ? 'latin' : 'ea');
            setRun(latin ? { fontLatin: f } : { fontEA: f });
          }}
        >
          <optgroup label="中文字体">
            {[...new Set([...EDITOR_FONTS_EA.filter((f) => f !== '等线'), ...(curRun.fontEA === '等线' ? [] : [curRun.fontEA]), '等线'])].map((f) => (
              <option key={`ea-${f}`} value={f}>{f}</option>
            ))}
          </optgroup>
          <optgroup label="西文字体">
            {[...new Set([curRun.fontLatin, ...EDITOR_FONTS_LATIN])]
              .filter((f) => !EDITOR_FONTS_EA.includes(f) && f !== curRun.fontEA)
              .map((f) => (
                <option key={`la-${f}`} value={f}>{f}</option>
              ))}
          </optgroup>
        </select>
        <FontSizeBox size={curRun.size} onPick={(size) => setRun({ size })} />
        <button className="icon-btn fmt" title="增大字号（Ctrl+Shift+>，Ctrl+] 每次 1 磅）" aria-label="增大字号" onMouseDown={(e) => e.preventDefault()} onClick={() => growFont(1, true)}>
          <span className="grow-ico">A<sup>▲</sup></span>
        </button>
        <button className="icon-btn fmt" title="减小字号（Ctrl+Shift+<，Ctrl+[ 每次 1 磅）" aria-label="减小字号" onMouseDown={(e) => e.preventDefault()} onClick={() => growFont(-1, true)}>
          <span className="grow-ico small">A<sup>▼</sup></span>
        </button>
        <button className={`icon-btn fmt${curRun.bold ? ' on' : ''}`} aria-pressed={curRun.bold} title="加粗（Ctrl+B）" onClick={() => setRun({ bold: !curRun.bold })} style={{ fontWeight: 800 }}>
          B
        </button>
        <button className={`icon-btn fmt${curRun.italic ? ' on' : ''}`} aria-pressed={curRun.italic} title="倾斜（Ctrl+I）" onClick={() => setRun({ italic: !curRun.italic })} style={{ fontStyle: 'italic', fontFamily: 'serif' }}>
          I
        </button>
        <button className={`icon-btn fmt${curRun.underline ? ' on' : ''}`} aria-pressed={curRun.underline} title="下划线（Ctrl+U）" onClick={() => setRun({ underline: !curRun.underline })} style={{ textDecoration: 'underline' }}>
          U
        </button>
        <button className={`icon-btn fmt${curRun.strike ? ' on' : ''}`} aria-pressed={!!curRun.strike} title="删除线" onClick={() => setRun({ strike: !curRun.strike })}>
          <span className="strike-ico">abc</span>
        </button>
        <button
          className={`icon-btn fmt${curRun.vertAlign === 'superscript' ? ' on' : ''}`}
          aria-pressed={curRun.vertAlign === 'superscript'}
          title="上标（Ctrl+Shift+=）"
          onClick={() => setRun({ vertAlign: curRun.vertAlign === 'superscript' ? undefined : 'superscript' })}
        >
          <span className="vert-ico">x<sup>2</sup></span>
        </button>
        <button
          className={`icon-btn fmt${curRun.vertAlign === 'subscript' ? ' on' : ''}`}
          aria-pressed={curRun.vertAlign === 'subscript'}
          title="下标（Ctrl+=）"
          onClick={() => setRun({ vertAlign: curRun.vertAlign === 'subscript' ? undefined : 'subscript' })}
        >
          <span className="vert-ico">x<sub>2</sub></span>
        </button>
        <span className="icon-btn color-btn" title="突出显示（文字底色）" style={{ position: 'relative' }}>
          <span className="color-ico">
            <span className="color-glyph hl-glyph">ab</span>
            <span className="color-bar" style={{ background: curRun.highlight ? `#${curRun.highlight}` : '#FFFF00' }} />
          </span>
          <Icon name="chevron" size={10} className="split-chev" />
          <select
            aria-label="突出显示"
            value={curRun.highlight ?? ''}
            onChange={(e) => setRun({ highlight: e.target.value || undefined })}
            style={{ position: 'absolute', inset: 0, opacity: 0, width: '100%', cursor: 'pointer' }}
          >
            <option value="">无底色</option>
            {curRun.highlight && !HIGHLIGHT_OPTIONS.some(([, v]) => v === curRun.highlight) && <option value={curRun.highlight}>文件里的</option>}
            {HIGHLIGHT_OPTIONS.map(([label, v]) => (
              <option key={v} value={v}>
                {label}
              </option>
            ))}
          </select>
        </span>
        <label className="icon-btn color-btn" title="字体颜色" style={{ position: 'relative' }}>
          <span className="color-ico">
            <span className="color-glyph">A</span>
            <span className="color-bar" style={{ background: curRun.color === 'auto' ? '#FF0000' : `#${curRun.color}` }} />
          </span>
          <Icon name="chevron" size={10} className="split-chev" />
          <input type="color" aria-label="字体颜色" value={curRun.color === 'auto' ? '#000000' : `#${curRun.color}`} onChange={(e) => setRun({ color: e.target.value.slice(1).toUpperCase() })} style={{ position: 'absolute', inset: 0, opacity: 0, width: '100%', cursor: 'pointer' }} />
        </label>
        <button className="icon-btn fmt" title="清除所有格式：文字和段落回到“正文”（Ctrl+Space 只清除文字格式）" aria-label="清除格式" onMouseDown={(e) => e.preventDefault()} onClick={() => clearFormatCmd(true)}>
          <span className="clear-ico">A<i>✕</i></span>
        </button>
        <button className="btn" title="字体（Ctrl+D）：字号、字形、颜色、字符间距等" onClick={() => setDialog('font')}>
          字体…
        </button>
        <span className="sep" />
        {alignBtn('left', '左对齐', '左对齐（Ctrl+L）')}
        {alignBtn('center', '居中', '居中（Ctrl+E）')}
        {alignBtn('right', '右对齐', '右对齐（Ctrl+R）')}
        {alignBtn('both', '两端对齐', '两端对齐（Ctrl+J）')}
        {alignBtn('distribute', '分散对齐', '分散对齐')}
        <button className="icon-btn fmt" title="减少缩进量（Ctrl+Shift+M）" aria-label="减少缩进量" onMouseDown={(e) => e.preventDefault()} onClick={() => indentCmd(-1)}>
          <IndentIcon dir={-1} />
        </button>
        <button className="icon-btn fmt" title="增加缩进量（Ctrl+M）" aria-label="增加缩进量" onMouseDown={(e) => e.preventDefault()} onClick={() => indentCmd(1)}>
          <IndentIcon dir={1} />
        </button>
        <button className="btn" title="段落：缩进、间距、行距、换行和分页" onClick={() => setDialog('para')}>
          段落…
        </button>
          </>
        ) : barTab === 'tools' ? (
          <>
        <button className="btn" title="插入超链接（Ctrl+K）；光标在链接里时编辑它。按住 Ctrl 点击链接可以打开" onClick={openLinkDialog}>
          超链接
        </button>
        <button className="btn" title="给选中的内容加书签，或跳到已有的书签" onClick={() => setDialog('bookmark')}>
          书签
        </button>
        <button className="btn" title="插入图片（页眉页脚里也可以，也可以直接粘贴截图）" onClick={() => imageInput.current?.click()}>
          图片
        </button>
        <input
          ref={imageInput}
          type="file"
          accept="image/png,image/jpeg,image/gif,image/bmp,image/webp"
          hidden
          onChange={(e) => {
            const f = e.target.files?.[0];
            e.target.value = '';
            if (f) void insertImageFile(f);
          }}
        />
        <button className={`btn${doc.current.watermark ? ' soft' : ''}`} title="文字水印（机密、草稿…）" onClick={() => setDialog('watermark')}>
          水印
        </button>
        <span className="sep" />
        <button className={`btn${doc.current.protection?.enforced ? ' soft' : ''}`} title="只读、只能批注、只能修订，可设密码；可以只放开要填的区域" onClick={() => setDialog('protect')}>
          限制编辑
        </button>
        <button className="btn" title="套用、新建、更新样式" onClick={() => setDialog('styles')}>
          样式…
        </button>
        <span className="sep" />
        <button className="btn" title="在文档里放 {{字段}}，用 Excel / csv 数据批量生成文档" onClick={() => setDialog('merge')}>
          模板与批量生成
        </button>
        {story !== 'body' && <span style={{ color: 'var(--text-dim, #888)', marginLeft: 8 }}>正在编辑{STORY_LABEL[story]}：图片、链接会放进页眉页脚</span>}
          </>
        ) : (
          <>
        <button className={`btn${navOpen ? ' soft' : ''}`} title="导航窗格：按标题跳转" onClick={() => setNavOpen(!navOpen)}>
          导航
        </button>
        <button className="btn" title="按标题 1–3 生成目录；已有目录时更新标题和页码" onClick={tocCmd}>
          目录
        </button>
        <button className="btn" onClick={() => void insertTableCmd()}>
          表格
        </button>
        <button className="btn" title="插入图片（也可以直接粘贴截图）" onClick={() => imageInput.current?.click()}>
          图片
        </button>
        <input
          ref={imageInput}
          type="file"
          accept="image/png,image/jpeg,image/gif,image/bmp,image/webp"
          hidden
          onChange={(e) => {
            const f = e.target.files?.[0];
            e.target.value = '';
            if (f) void insertImageFile(f);
          }}
        />
        <select className="ed-select" style={{ width: 80 }} aria-label="分隔符" value="" onChange={(e) => breakCmd(e.target.value)}>
          <option value="" disabled>
            分隔符
          </option>
          <option value="page">分页符（Ctrl+Enter）</option>
          <option value="column">分栏符</option>
          {(Object.keys(SECTION_BREAK_LABEL) as SectionBreak[]).map((k) => (
            <option key={k} value={k}>
              {SECTION_BREAK_LABEL[k]}
            </option>
          ))}
        </select>
        <button className={`btn${story !== 'body' ? ' soft' : ''}`} onClick={toggleHeaderFooter} title="编辑页眉页脚（也可双击页面顶部/底部）">
          页眉页脚
        </button>
        <select className="ed-select" style={{ width: 80 }} aria-label="页码" value="" onChange={(e) => pageNumberCmd(e.target.value)}>
          <option value="" disabled>
            页码
          </option>
          <option value="full">第 X 页 共 Y 页</option>
          <option value="dash">- X -</option>
          <option value="plain">X</option>
        </select>
        {story !== 'body' && (
          <button className="btn" onClick={() => insertFieldCmd('PAGE')} title="在光标处插入当前页码">
            插入页码域
          </button>
        )}
        {hfLinkedNow !== null && (
          <button
            className={`btn${hfLinkedNow ? ' soft' : ''}`}
            aria-pressed={hfLinkedNow}
            onClick={toggleHFLink}
            title={hfLinkedNow ? '本节与上一节使用相同的页眉页脚，点击后本节可单独编辑' : '点击后本节改用上一节的页眉页脚'}
          >
            与上一节相同
          </button>
        )}
        <span className="sep" />
        <select
          className="ed-select"
          style={{ width: 118 }}
          aria-label="行距"
          value={lineIdx}
          onChange={(e) => setPara({ line: LINE_OPTIONS[Number(e.target.value)].value })}
        >
          {lineIdx < 0 && <option value={-1}>{curPara.line.rule === 'auto' ? `${curPara.line.value.toFixed(2)} 倍` : `${curPara.line.rule === 'exact' ? '固定' : '最小'} ${curPara.line.value} 磅`}</option>}
          {LINE_OPTIONS.map((o, i) => (
            <option key={o.label} value={i}>
              {o.label}
            </option>
          ))}
        </select>
        <button
          className={`btn${curPara.firstLineChars === 2 ? ' soft' : ''}`}
          title="首行缩进 2 字符"
          onClick={() => setPara(curPara.firstLineChars === 2 ? { firstLineChars: null, firstLine: 0 } : { firstLineChars: 2, firstLine: 0 })}
        >
          首行缩进
        </button>
        <select
          className="ed-select"
          style={{ width: 92 }}
          aria-label="段前段后"
          value=""
          onChange={(e) => {
            const [b, a] = e.target.value.split(',').map(Number);
            setPara({ spaceBeforeLines: b, spaceAfterLines: a, spaceBefore: 0, spaceAfter: 0 });
          }}
        >
          <option value="" disabled>
            段间距
          </option>
          <option value="0,0">无</option>
          <option value="0,0.5">段后 0.5 行</option>
          <option value="0.5,0.5">段前后 0.5 行</option>
          <option value="1,1">段前后 1 行</option>
        </select>
        <select className="ed-select" style={{ width: 118 }} aria-label="自动编号" value={curPreset} onChange={(e) => numberingCmd(e.target.value)}>
          <option value="">无编号</option>
          {curPreset && !NUM_PRESETS.some((p) => p.key === curPreset) && <option value={curPreset}>文档中的编号</option>}
          {NUM_PRESETS.map((p) => (
            <option key={p.key} value={p.key}>
              {p.label}
            </option>
          ))}
        </select>
        {curPara.num && (
          <>
            <button className="icon-btn fmt" title="编号升级（Shift+Tab）" aria-label="编号升级" onClick={() => levelCmd(-1)}>
              ←
            </button>
            <button className="icon-btn fmt" title="编号降级（段首按 Tab）" aria-label="编号降级" onClick={() => levelCmd(1)}>
              →
            </button>
          </>
        )}
        <button className="btn" title="纸张、页边距、文档网格（作用于光标所在的节）" onClick={() => setSetupOpen(true)}>
          页面
        </button>
          </>
        )}
      </div>
      </>)}
      {inTable && !mobile && (
        <div className="word-subbar">
          <span className="sub">表格：</span>
          <button className="btn" onClick={() => tableCmd((d, p) => insertRow(d, p, false))}>上方插入行</button>
          <button className="btn" onClick={() => tableCmd((d, p) => insertRow(d, p, true))}>下方插入行</button>
          <button className="btn" onClick={() => tableCmd((d, p) => insertColumn(d, p, true))}>左侧插入列</button>
          <button className="btn" onClick={() => tableCmd((d, p) => insertColumn(d, p))}>右侧插入列</button>
          <button className="btn" onClick={() => tableCmd(deleteRow)}>删除行</button>
          <button className="btn" onClick={() => tableCmd(deleteColumn)}>删除列</button>
          <button className="btn" onClick={() => void rowHeightCmd()}>行高</button>
          <button className="btn" onClick={() => void colWidthCmd()}>列宽</button>
          <button className="btn" title="各列宽度相同，表格总宽不变" onClick={() => mutate('table', (d) => void distributeColumns(d.body[inTable.bi] as Table))}>
            平均分布各列
          </button>
          <span className="sep" />
          <button className="btn" title="先拖动选中几个单元格" onClick={mergeCmd}>合并单元格</button>
          <button className="btn" onClick={() => tableCmd(unmergeCell, '当前单元格没有合并')}>取消合并</button>
          <button className="btn" onClick={() => void splitCmd()}>拆分单元格</button>
          <select
            className="select"
            title="单元格底纹"
            value=""
            onChange={(e) => {
              const v = e.target.value;
              if (v) cellFormatCmd({ shading: v === 'none' ? null : v });
            }}
          >
            <option value="">底纹</option>
            <option value="none">无底纹</option>
            {SHADING_OPTIONS.map(([label, v]) => (
              <option key={v} value={v}>
                {label}
              </option>
            ))}
          </select>
          <select
            className="select"
            style={{ width: 104 }}
            title="单元格对齐方式（水平对齐也可以用工具条上的对齐按钮）"
            value=""
            onChange={(e) => {
              const v = e.target.value;
              if (v === 'mid') cellFormatCmd({ vAlign: 'center', align: 'center' });
              else if (v) cellFormatCmd({ vAlign: v as 'top' | 'center' | 'bottom' });
            }}
          >
            <option value="">对齐（{inTable.table.rows[inTable.r].cells[inTable.c]?.vAlign === 'center' ? '居中' : inTable.table.rows[inTable.r].cells[inTable.c]?.vAlign === 'bottom' ? '靠下' : '靠上'}）</option>
            <option value="top">靠上</option>
            <option value="center">垂直居中</option>
            <option value="bottom">靠下</option>
            <option value="mid">水平垂直居中</option>
          </select>
          <select
            className="select"
            style={{ width: 76 }}
            title={`单元格框线：用当前线型（${pen.style === 'dashed' ? '虚线' : pen.style === 'dotted' ? '点线' : pen.style === 'double' ? '双线' : '实线'} ${pen.w} 磅）画到选中格子的哪些边；线型、粗细、颜色在“边框…”里改`}
            value=""
            onChange={(e) => {
              const v = e.target.value;
              if (v === 'inherit') cellFormatCmd({ borders: null });
              else if (v) bordersCmd(v as BorderTarget);
            }}
          >
            <option value="">框线</option>
            {BORDER_TARGETS.map(([v, l]) => (
              <option key={v} value={v}>
                {l}
              </option>
            ))}
            <option value="inherit">跟随表格（清除单元格框线）</option>
          </select>
          <button className="btn" title="框线的线型、粗细、颜色" onClick={() => setDialog('borders')}>
            边框…
          </button>
          <span className="sep" />
          <button
            className={`btn${inTable.table.rows[inTable.r].header ? ' soft' : ''}`}
            title="跨页时在每页顶端重复此行"
            onClick={() => mutate('table', (d) => void ((d.body[inTable.bi] as typeof inTable.table).rows[inTable.r].header = !inTable.table.rows[inTable.r].header))}
          >
            标题行重复
          </button>
          <button className="btn" onClick={() => mutate('table', (d) => caretSel(deleteTable(d, inTable.bi)))}>删除表格</button>
        </div>
      )}
      {(selectedImage || floatSel?.image) && !mobile && (
        <div className="word-subbar">
          <span className="sub">
            图片：{Math.round((floatSel?.image ?? selectedImage)!.w / PT_PER_CM * 10) / 10} × {Math.round((floatSel?.image ?? selectedImage)!.h / PT_PER_CM * 10) / 10} 厘米
          </span>
          <button className="btn" onClick={openPictureDialog}>设置图片格式…</button>
          {selectedImage && !floatSel && <>
          <button className="btn" onClick={() => resizeImage(0.9)}>缩小</button>
          <button className="btn" onClick={() => resizeImage(1 / 0.9)}>放大</button>
          <button className="btn" onClick={() => resizeImage('fit')}>宽度适合页面</button>
          </>}
        </div>
      )}
      {reviewBar && !mobile && (
        <div className="word-subbar">
          <span className="sub">审阅：</span>
          <button className={`btn${trackOn ? ' soft on' : ''}`} aria-pressed={trackOn} title={trackOn ? `正在记录修订（修订人：${author}），点击停止` : '打开后，输入和删除都会记为修订'} onClick={() => (activeProtection(doc.current)?.edit === 'trackedChanges' ? host().toast('文档受保护（只能修订），修订不能关闭') : setTrackOn(!trackOn), refocus())}>
            {trackOn ? '修订：开' : '修订：关'}
          </button>
          <span className="sep" />
          <button className="btn" onClick={() => void newComment()}>新建批注</button>
          <button className={`btn${commentsOpen ? ' soft' : ''}`} onClick={() => setCommentsOpen(!commentsOpen)}>
            批注栏{doc.current.comments?.length ? `（${doc.current.comments.length}）` : ''}
          </button>
          <span className="rev-actions" onMouseDown={(e) => e.preventDefault()}>
            <button className="btn" onClick={() => gotoRevision(true)}>上一处</button>
            <button className="btn" onClick={() => gotoRevision(false)}>下一处</button>
            <button className="btn" onClick={() => resolveCmd(true)}>接受</button>
            <button className="btn" onClick={() => resolveCmd(false)}>拒绝</button>
            <button className="btn" onClick={() => resolveCmd(true, true)}>全部接受</button>
            <button className="btn" onClick={() => resolveCmd(false, true)}>全部拒绝</button>
          </span>
        </div>
      )}
      {doc.current.protection?.enforced && (
        <div className="ed-warn" role="status">
          此文档受保护（{protectLabel(doc.current.protection)}）{doc.current.protection.edit === 'trackedChanges' ? '：所有修改都记为修订。' : permRanges(doc.current).length ? '：只有浅黄色的区域可以修改。' : '：不能修改。'}
          <button className="btn ghost" onClick={() => setDialog('protect')}>停止保护…</button>
        </div>
      )}
      {lossy.length > 0 && (
        <div className="ed-warn" role="status">
          此文档包含暂不支持的内容：{lossy.join('、')}。编辑后保存会丢失这些内容，建议使用“另存”。
          <button className="btn ghost" onClick={() => setLossy([])}>知道了</button>
        </div>
      )}
      {findMode && (
        <WordFind
          key={searchSeq}
          initialQuery={searchText}
          doc={doc.current}
          at={sel.focus}
          current={
            !collapsed(sel) && sel.anchor.story === sel.focus.story && sel.anchor.path.join() === sel.focus.path.join()
              ? { story: sel.focus.story, path: sel.focus.path, offset: ordered(sel)[0].offset, len: Math.abs(sel.focus.offset - sel.anchor.offset) }
              : null
          }
          mode={findMode}
          version={version}
          onMode={setFindMode}
          onSelect={(m: Match) => {
            setPending(null);
            setSel({ anchor: { story: m.story, path: m.path, offset: m.offset }, focus: { story: m.story, path: m.path, offset: m.offset + m.len } });
          }}
          onReplace={(m: Match, repl: string) => mutate('replace', (d) => caretSel(replaceMatch(d, m, repl)))}
          onReplaceAll={(q, repl, o) => {
            let n = 0;
            mutate('replace', (d) => {
              n = replaceAllText(d, q, repl, o);
              return n ? undefined : null;
            });
            return n;
          }}
          onClose={() => {
            setFindMode(null);
            refocus();
          }}
        />
      )}
      <div className="word-main">
      {navOpen && (
        <nav className={`word-nav${mobile ? ' word-nav-mobile' : ''}`} aria-label="导航" onClick={mobile ? () => setNavOpen(false) : undefined}>
          <div className="sub">标题</div>
          {headings(doc.current, 9).map((h) => (
            <button
              key={h.bi}
              className={sel.focus.story === 'body' && sel.focus.path[0] >= h.bi && !headings(doc.current, 9).some((x) => x.bi > h.bi && x.bi <= sel.focus.path[0]) ? 'on' : undefined}
              style={{ paddingLeft: 8 + (h.level - 1) * 14 }}
              title={h.text}
              onClick={() => {
                setSel(caretSel({ story: 'body', path: [h.bi], offset: 0 }));
                refocus();
              }}
            >
              {h.text}
            </button>
          ))}
          {!headings(doc.current, 9).length && <div className="sub">文档里没有标题。把章节标题设为“标题 1 / 2 / 3”样式后会显示在这里。</div>}
        </nav>
      )}
      <div className="word-center" ref={centerEl}>
      {rulerOn && !mobile && rulerPage && (
        <WordRuler
          ppt={ppt}
          pageLeft={rulerLeft}
          setup={rulerPage.setup}
          area={rulerArea}
          para={rulerArea ? curPara : null}
          fontSize={curRun.size}
          onIndent={(patch) => setPara(patch)}
          onMargin={(side, x) => {
            const geom = marginPatch(side, x, rulerPage.setup);
            if (geom) mutate('page', (d) => void setSectionPage(d, caretSection().index, geom, false));
          }}
          onTabs={(tabs) => setPara({ tabs })}
        />
      )}
      <div className="word-body">
      {rulerOn && !mobile && rulerPage && (
        <WordVRuler
          ppt={ppt}
          pageTop={rulerTop}
          setup={rulerPage.setup}
          onMargin={(side, y) => {
            const geom = marginPatchV(side, y, rulerPage.setup);
            if (geom) mutate('page', (d) => void setSectionPage(d, caretSection().index, geom, false));
          }}
        />
      )}
      <div className="word-scroll" ref={scroller}>
        <div className="word-pages" ref={pagesEl} style={{ gap: mobile ? 8 : PAGE_GAP }}>
          {layout.pages.map((pg) => (
            <div
              key={pg.index}
              ref={(el) => (pageEls.current[pg.index] = el)}
              className="word-page"
              style={{ width: pg.setup.width * ppt, height: pg.setup.height * ppt, cursor: painter ? 'copy' : 'text' }}
              onPointerDown={(e) => (mobile && !mSelectOn ? mobileDown(e, pg.index) : onPagePointerDown(e, pg.index))}
              onPointerMove={(e) => (!mobile || mSelectOn) && onPagePointerMove(e, pg.index)}
              onPointerCancel={mobile && !mSelectOn ? mobileCancel : undefined}
              onPointerUp={(e) => {
                if (mobile && !mSelectOn) return mobileUp(e, pg.index);
                if (dragging.current && painter) applyPainter(sel);
                dragging.current = false;
                const cd = colDrag.current;
                colDrag.current = null;
                if (cd && colGhost?.dx) {
                  const dx = colGhost.dx;
                  mutate('table', (d) => void moveColumnLine(d.body[cd.bi] as Table, cd.g, dx));
                }
                setColGhost(null);
                const fd = floatDrag.current;
                floatDrag.current = null;
                if (fd && floatGhost && floatSel) {
                  if (fd.mode === 'move') moveFloat(floatSel, { x: floatGhost.x, y: floatGhost.y });
                  else moveFloat(floatSel, { x: floatGhost.x, y: floatGhost.y, w: Math.round(floatGhost.w * 10) / 10, h: Math.round(floatGhost.h * 10) / 10 });
                }
                setFloatGhost(null);
              }}
              onDoubleClick={(e) => !mobile && onPageDoubleClick(e, pg.index)}
              onContextMenu={(e) => {
                if (!mobile) return onPageContextMenu(e, pg.index);
                // 安卓的长按会发 contextmenu：编辑模式下用它选词（计时器先到时已选过，这里不重复）
                e.preventDefault();
                if (mTouch.current && !mSelectOn) mSelectWordAt(e, pg.index);
              }}
            >
              <PageCanvas page={pg} ppt={ppt} story={story} marks={mobile ? (mMode === 'edit' ? 'para' : 'none') : marks} corners={!mobile || mMode === 'edit'} selection={selRectsByPage.get(pg.index) ?? NO_SEL} highlights={[...(commentRectsByPage.get(pg.index) ?? NO_SEL), ...(permRectsByPage.get(pg.index) ?? NO_SEL)]} />
            </div>
          ))}
          {focused && collapsed(sel) && caretStyle && !floatSel && <div className="word-caret" style={caretStyle} />}
          {floatRev && pageEls.current[floatRev.rect.page] && (
            <div
              className="word-rev-pop"
              style={{ left: pageEls.current[floatRev.rect.page]!.offsetLeft + floatRev.rect.x * ppt, top: pageEls.current[floatRev.rect.page]!.offsetTop + (floatRev.rect.y + floatRev.rect.h) * ppt + 4 }}
              onMouseDown={(e) => e.preventDefault()}
            >
              <button className="btn" onClick={() => resolveCmd(true, false, floatRev.range)}>接受</button>
              <button className="btn reject" onClick={() => resolveCmd(false, false, floatRev.range)}>拒绝</button>
            </div>
          )}
          {colGhost && pageEls.current[colGhost.page] && (
            <div
              className="word-col-ghost"
              style={{
                left: pageEls.current[colGhost.page]!.offsetLeft + (colGhost.x + (colGhost.dx ?? 0)) * ppt,
                top: pageEls.current[colGhost.page]!.offsetTop + colGhost.y * ppt,
                height: colGhost.h * ppt,
              }}
            />
          )}
          {(() => {
            // 选中的浮动图片：边框 + 右下角缩放手柄；拖动时显示新位置的虚线框
            const f = floatSel && layout.pages.flatMap((pg) => pg.floats.map((x) => ({ ...x, page: pg.index }))).find((x) => x.box === floatSel);
            const g = floatGhost ?? (f ? { page: f.page, x: f.x, y: f.y, w: f.box.w, h: f.box.h } : null);
            const el = g && pageEls.current[g.page];
            if (!g || !el) return null;
            return (
              <div
                className={`word-float-sel${floatGhost ? ' ghost' : ''}`}
                style={{ left: el.offsetLeft + g.x * ppt, top: el.offsetTop + g.y * ppt, width: g.w * ppt, height: g.h * ppt, transform: `rotate(${floatSel?.image?.rot ?? 0}deg)` }}
              >
                <span className="word-float-handle" />
              </div>
            );
          })()}
          <textarea
            ref={input}
            className="word-input"
            inputMode={mobile && mMode !== 'edit' ? 'none' : undefined}
            style={caretStyle ? { left: caretStyle.left, top: caretStyle.top, fontSize: curRun.size * ppt } : undefined}
            aria-label="文档编辑区"
            autoCapitalize="off"
            autoComplete="off"
            spellCheck={false}
            onFocus={() => setFocused(true)}
            onBlur={() => setFocused(false)}
            onKeyDown={onKeyDown}
            onCompositionStart={() => (composing.current = true)}
            onCompositionEnd={(e) => {
              composing.current = false;
              flushInput(e.data);
            }}
            onInput={(e) => {
              if (composing.current || (e.nativeEvent as InputEvent).isComposing) return;
              flushInput(e.currentTarget.value);
            }}
            onCopy={(e) => {
              e.preventDefault();
              if (!collapsed(sel)) writeClip(e.clipboardData);
            }}
            onCut={(e) => {
              e.preventDefault();
              if (collapsed(sel)) return;
              writeClip(e.clipboardData);
              mutate('cut', (d, s) => {
                const p = clearSelection(d, s, false);
                return p ? caretSel(p) : null;
              });
            }}
            onPaste={(e) => {
              e.preventDefault();
              const plain = plainPaste.current;
              plainPaste.current = false;
              // 截图等剪贴板里的图片
              const img = Array.from(e.clipboardData.files).find((f) => f.type.startsWith('image/'));
              if (img && !e.clipboardData.getData('text/plain')) return void insertImageFile(img);
              pasteFrom(e.clipboardData.getData('text/plain'), plain ? '' : e.clipboardData.getData('text/html'));
            }}
          />
        </div>
      </div>
      </div>
      </div>
      {palette && !mobile && (
        <aside className="word-palette" aria-label={palette.title ?? '插入字段'}>
          <div className="word-palette-head">
            <strong>{palette.title ?? '插入字段'}</strong>
            <span className="sub">点击在光标处插入 {'{{名称}}'}</span>
            {ai?.templateFields && (
              <button
                className={`btn word-palette-ai-btn${tplBusy ? ' soft on' : ''}`}
                title={tplBusy ? '点击取消' : 'AI 识别合同里的公司、合同号、金额等数据，换成字段（记为修订，审阅后确认）'}
                onClick={() => void aiTemplate()}
              >
                {tplBusy ? 'AI 分析中…（取消）' : 'AI 生成'}
              </button>
            )}
          </div>
          {tplReport && (
            <div className="word-palette-ai">
              <button className="word-palette-ai-close" aria-label="关闭" onClick={() => setTplReport(null)}>
                ×
              </button>
              {tplReport}
            </div>
          )}
          <div className="word-palette-cols word-palette-titles">
            {palette.columns.map((col) => (
              <div key={col.title} className="word-palette-coltitle">
                {col.title}
              </div>
            ))}
          </div>
          <div className="word-palette-scroll">
            <div className="word-palette-cols">
              {palette.columns.map((col) => (
                <div key={col.title} className="word-palette-col">
                  {col.groups.map((g) => (
                    <div key={g.title} className="word-palette-group">
                      <div className="sub">{g.title}</div>
                      <div className="word-palette-chips">
                        {g.fields.map((f) => (
                          <button
                            key={f.key}
                            className="word-chip"
                            title={`{{${f.key}}}`}
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => {
                              insertFieldCmd2(f.key, false);
                              refocus();
                            }}
                          >
                            {f.label ?? f.key}
                          </button>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>
        </aside>
      )}
      {reviewBar && commentsOpen && !mobile && (doc.current.comments?.length ?? 0) > 0 && (
        <CommentsPanel
          doc={doc.current}
          ranges={ranges}
          onPick={(id) => {
            const r = ranges.get(id);
            if (r) setSel({ anchor: r[0], focus: r[1] });
            refocus();
          }}
          onReply={(id) => void replyCmd(id)}
          onDelete={(id) => mutate('comment', (d) => void deleteComment(d, id))}
        />
      )}
      </div>
      {mobile && mMode === 'review' && mComments && !mFinal && (
        <div className="word-m-comments" role="dialog" aria-label="批注">
          <div className="word-m-comments-head">
            <strong>批注</strong>
            <button className="icon-btn" aria-label="关闭批注" onClick={() => setMComments(false)}>
              <Icon name="close" size={18} />
            </button>
          </div>
          {(doc.current.comments?.length ?? 0) > 0 ? (
            <CommentsPanel
              doc={doc.current}
              ranges={ranges}
              onPick={(id) => {
                const r = ranges.get(id);
                if (r) setSel({ anchor: r[0], focus: r[1] });
                setMComments(false);
              }}
              onReply={(id) => void replyCmd(id)}
              onDelete={(id) => mutate('comment', (d) => void deleteComment(d, id))}
            />
          ) : (
            <div className="sub word-m-empty">还没有批注。长按选中文字，再点“新批注”。</div>
          )}
        </div>
      )}
      {mobile && (
        <WordMobileBar
          mode={mMode}
          onMode={setMMode}
          page={mPage}
          pages={layout.pages.length}
          fit={mFit}
          navOpen={navOpen}
          onNav={() => setNavOpen(!navOpen)}
          onFind={() => setFindMode(findMode ? null : 'find')}
          onFit={() => setMFit(true)}
          ed={mEd}
          rev={mRev}
          onMerge={
            host().startMerge && tab.path
              ? () => (dirty ? host().toast('请先保存修改，再用它套打') : host().startMerge?.(tab.path!))
              : undefined
          }
          readOnly={activeProtection(doc.current) && activeProtection(doc.current)?.edit !== 'trackedChanges' ? protectLabel(doc.current.protection!) : null}
        />
      )}
      {!mobile && <div className="word-status">
        <span>
          第 {caretPage + 1} 页，共 {layout.pages.length} 页
        </span>
        <span>字数 {words}</span>
        <span>
          {story !== 'body' && sectionCount > 1 && `第 ${(layout.pages[caretPage]?.section ?? 0) + 1} 节 `}
          {STORY_LABEL[story]}
          {story !== 'body' && '（Esc 返回正文）'}
        </span>
        <span>
          {setup.grid.type === 'none'
            ? '无网格'
            : `每页 ${Math.floor(textHeight(setup) / setup.grid.linePitch + 0.01)} 行${setup.grid.type !== 'lines' ? `，每行 ${Math.floor(textWidth(setup) / charPitch(doc.current) + 0.01)} 字` : ''}`}
        </span>
        <span className="grow" />
        <button className="icon-btn" aria-label="缩小" onClick={() => setZoom((z) => Math.max(0.5, +(z - 0.1).toFixed(2)))}>
          <Icon name="zoomOut" size={14} />
        </button>
        <span className="mono">{Math.round(zoom * 100)}%</span>
        <button className="icon-btn" aria-label="放大" onClick={() => setZoom((z) => Math.min(3, +(z + 0.1).toFixed(2)))}>
          <Icon name="zoomIn" size={14} />
        </button>
      </div>}
      {settingsOpen && (
        <WordSettings
          rulerOn={rulerOn}
          setRulerOn={setRulerOn}
          marks={marks}
          setMarks={changeMarks}
          onPageSetup={() => {
            setSettingsOpen(false);
            setSetupOpen(true);
          }}
          onClose={() => {
            setSettingsOpen(false);
            refocus();
          }}
        />
      )}
      {fieldEdit && (
        <FieldDialog
          instr={fieldEdit.instr}
          env={{ ...fieldEnv(), bookmark: (n) => bookmarkText(doc.current, n) }}
          bookmarks={listBookmarks(doc.current).map((b) => b.name)}
          onOk={applyFieldEdit}
          onClose={() => {
            setFieldEdit(null);
            refocus();
          }}
        />
      )}
      {pictureEdit && <PictureDialog value={pictureEdit.value} onApply={savePictureDialog} onClose={() => { setPictureEdit(null); refocus(); }} />}
      {dialog === 'link' && (
        <LinkDialog
          text={linkHere()?.text ?? selText()}
          link={linkHere()?.link}
          bookmarks={listBookmarks(doc.current).map((b) => b.name)}
          onOk={applyLink}
          onRemove={linkHere() || !collapsed(sel) ? removeLinkCmd : undefined}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
        />
      )}
      {dialog === 'bookmark' && (
        <BookmarkDialog
          bookmarks={listBookmarks(doc.current)}
          canAdd={sel.anchor.story === 'body' && sel.focus.story === 'body'}
          suggest=""
          onAdd={addBookmarkCmd}
          onGo={(name) => {
            setDialog(null);
            goBookmark(name);
          }}
          onDelete={(name) => mutate('bookmark', (d) => void removeBookmark(d, name))}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
        />
      )}
      {dialog === 'watermark' && (
        <WatermarkDialog
          value={doc.current.watermark}
          onOk={applyWatermark}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
        />
      )}
      {dialog === 'styles' && (
        <StylesDialog
          styles={withEdits(styleCat(), doc.current.styleEdits).list()}
          currentPara={getPara(doc.current, sel.focus).props.styleId}
          currentChar={runPropsAt(doc.current, sel.focus).rStyle}
          hasSelection={!collapsed(sel)}
          onApply={applyStyleCmd}
          onCreate={createStyleCmd}
          onRedefine={redefineStyleCmd}
          onRename={renameStyleCmd}
          onDelete={deleteStyleCmd}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
        />
      )}
      {dialog === 'font' && (
        <FontDialog
          run={curRun}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
          onOk={(patch) => {
            setDialog(null);
            if (Object.keys(patch).length) setRun(patch);
            else refocus();
          }}
        />
      )}
      {dialog === 'para' && (
        <ParagraphDialog
          props={curPara}
          charPt={charPtOf(getPara(doc.current, sel.focus))}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
          onOk={(patch) => {
            setDialog(null);
            if (Object.keys(patch).length) setPara(patch);
            else refocus();
          }}
        />
      )}
      {dialog === 'borders' && (
        <BordersDialog
          pen={pen}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
          onOk={(target, line) => {
            setDialog(null);
            if (target !== 'none') setPen(line);
            bordersCmd(target, line);
          }}
        />
      )}
      {menu && <WordMenu x={menu.x} y={menu.y} items={menuItems()} onClose={() => setMenu(null)} />}
      {dialog === 'protect' && (
        <ProtectDialog
          protection={doc.current.protection}
          hasSelection={!collapsed(sel) && sel.anchor.story === 'body' && sel.focus.story === 'body'}
          permCount={permRanges(doc.current).length}
          onStart={startProtect}
          onStop={stopProtect}
          onAddPerm={() => {
            const [a, b] = ordered(sel);
            mutate('perm', (d) => void addPermRange(d, a, b));
          }}
          onClearPerms={() => mutate('perm', (d) => void clearPermRanges(d))}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
        />
      )}
      {dialog === 'merge' && (
        <MergeDialog
          fields={tmplFields}
          tabPath={tab.path}
          generate={generateMerged}
          generateAll={generateCombined}
          onInsertField={insertFieldCmd2}
          onInsertNext={() => mutate('field', (d) => caretSel(insertNextRecord(d, sel.focus)))}
          onClose={() => {
            setDialog(null);
            refocus();
          }}
        />
      )}
      {setupOpen && (
        <PageSetupDialog
          doc={doc.current}
          section={caretSection()}
          sectionCount={sectionCount}
          onClose={() => {
            setSetupOpen(false);
            refocus();
          }}
          onApply={(fn) => {
            mutate('setup', (d) => fn(d));
            setSetupOpen(false);
          }}
        />
      )}
    </div>
  );
}

function AlignIcon({ a }: { a: Align }) {
  const rows: [number, number][] =
    a === 'left' ? [[2, 12], [2, 8], [2, 12], [2, 6]] :
    a === 'center' ? [[2, 12], [4, 8], [2, 12], [5, 6]] :
    a === 'right' ? [[2, 12], [6, 8], [2, 12], [8, 6]] :
    a === 'both' ? [[2, 12], [2, 12], [2, 12], [2, 7]] :
    [[2, 12], [2, 12], [2, 12], [2, 12]];
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round">
      {rows.map(([x, w], i) => (
        <path key={i} d={`M${x} ${3 + i * 3.4}h${w}`} />
      ))}
    </svg>
  );
}

/** 增加 / 减少缩进量的图标：四条横线，前面一个向右 / 向左的小三角 */
function IndentIcon({ dir }: { dir: 1 | -1 }) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round">
      <path d="M2 3h12M7 6.4h7M7 9.8h7M2 13.2h12" />
      <path d={dir > 0 ? 'M2 5.6l2.6 2.4L2 10.4z' : 'M4.6 5.6L2 8l2.6 2.4z'} fill="currentColor" strokeWidth="1" strokeLinejoin="round" />
    </svg>
  );
}

/** 单页画布：进入可视区域才绘制；内容、缩放、选区都没变时不重绘。 */
function PageCanvas({ page, ppt, story, marks, corners = true, selection, highlights }: { page: PageLayout; ppt: number; story: Story; marks: MarksMode; /** 版心角标（手机查看时不画） */ corners?: boolean; selection: SelRect[]; highlights: SelRect[] }) {
  const ref = useRef<HTMLCanvasElement>(null);
  const [near, setNear] = useState(false);
  const drawn = useRef('');

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const io = new IntersectionObserver(([e]) => setNear(e.isIntersecting), { rootMargin: '600px 0px' });
    io.observe(el);
    return () => io.disconnect();
  }, []);

  useEffect(() => {
    const c = ref.current;
    if (!c || !near) return;
    const dpr = window.devicePixelRatio || 1;
    const s = page.setup;
    const key = `${pageSignature(page)}§${ppt}§${dpr}§${story}§${marks}§${corners}§${selection.map((r) => `${r.x},${r.y},${r.w}`).join(';')}§${highlights.map((r) => `${r.x},${r.y},${r.w}`).join(';')}§${s.width}x${s.height}§${s.marginLeft},${s.marginTop},${s.marginRight},${s.marginBottom}`;
    if (key === drawn.current) return;
    drawn.current = key;
    const w = Math.round(s.width * ppt * dpr);
    const h = Math.round(s.height * ppt * dpr);
    if (c.width !== w || c.height !== h) {
      c.width = w;
      c.height = h;
    }
    drawPage(c.getContext('2d')!, page, ppt * dpr, { story, selection, highlights, showMarks: corners, formatMarks: marks === 'none' ? null : marks });
  }, [near, page, ppt, story, marks, corners, selection, highlights]);

  return <canvas ref={ref} style={{ width: '100%', height: '100%', display: 'block' }} />;
}

const SETTING_GROUPS = ['视图', '页面'] as const;

/** 工具条上不常用的功能：左边分类，右边具体设置 */
function WordSettings({
  rulerOn,
  setRulerOn,
  marks,
  setMarks,
  onPageSetup,
  onClose,
}: {
  rulerOn: boolean;
  setRulerOn: (on: boolean) => void;
  marks: MarksMode;
  setMarks: (m: MarksMode) => void;
  onPageSetup: () => void;
  onClose: () => void;
}) {
  const [group, setGroup] = useState<(typeof SETTING_GROUPS)[number]>('视图');
  return (
    <div className="backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="dialog word-settings" role="dialog" aria-modal="true" aria-label="设置" style={{ width: 560 }}>
        <h3>设置</h3>
        <div className="ws-body">
          <div className="ws-nav" role="tablist">
            {SETTING_GROUPS.map((g) => (
              <button key={g} role="tab" aria-selected={group === g} className={group === g ? 'on' : ''} onClick={() => setGroup(g)}>
                {g}
              </button>
            ))}
          </div>
          <div className="ws-pane">
            {group === '视图' && (
              <>
              <div className="ws-row">
                <span>
                  编辑标记
                  <small>回车（↵）、分页符等不会打印出来的标记</small>
                </span>
                <select className="field" style={{ width: 200 }} value={marks} onChange={(e) => setMarks(e.target.value as MarksMode)}>
                  {(Object.keys(MARKS_LABEL) as MarksMode[]).map((m) => (
                    <option key={m} value={m}>
                      {MARKS_LABEL[m]}
                    </option>
                  ))}
                </select>
              </div>
              <label className="ws-row">
                <input type="checkbox" checked={rulerOn} onChange={(e) => setRulerOn(e.target.checked)} />
                <span>
                  显示标尺
                  <small>拖动缩进标记、制表位和页边距来排版</small>
                </span>
              </label>
              </>
            )}
            {group === '页面' && (
              <div className="ws-row">
                <span>
                  页面设置
                  <small>纸张、页边距、文档网格（作用于光标所在的节）</small>
                </span>
                <button className="btn" onClick={onPageSetup}>
                  打开…
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="dialog-foot">
          <button className="btn primary" onClick={onClose}>
            关闭
          </button>
        </div>
      </div>
    </div>
  );
}

function PageSetupDialog({
  doc,
  section,
  sectionCount,
  onClose,
  onApply,
}: {
  doc: Doc;
  /** 光标所在的节，对话框显示它的纸张与页边距 */
  section: SectionInfo;
  sectionCount: number;
  onClose: () => void;
  onApply: (fn: (d: Doc) => void) => void;
}) {
  const s = section.setup;
  // 多节文档：纸张与页边距可只应用于本节
  const [wholeDoc, setWholeDoc] = useState(sectionCount < 2);
  const cm = (pt: number) => +(pt / PT_PER_CM).toFixed(2);
  const [paper, setPaper] = useState(PAPERS.find((p) => Math.abs(p.width - Math.min(s.width, s.height)) < 2 && Math.abs(p.height - Math.max(s.width, s.height)) < 2)?.name ?? '自定义');
  const [landscape, setLandscape] = useState(s.width > s.height);
  const [m, setM] = useState({ top: cm(s.marginTop), bottom: cm(s.marginBottom), left: cm(s.marginLeft), right: cm(s.marginRight), header: cm(s.headerDistance), footer: cm(s.footerDistance) });
  const [grid, setGrid] = useState<GridType>(s.grid.type === 'snapToChars' ? 'linesAndChars' : s.grid.type);
  const [lines, setLines] = useState(Math.floor(textHeight(s) / s.grid.linePitch + 0.01));
  const [chars, setChars] = useState(Math.floor(textWidth(s) / charPitch(doc) + 0.01));
  const [gw, setGw] = useState<{ size: number; font: string } | null>(null);
  const [titlePg, setTitlePg] = useState(s.titlePg);
  const [evenAndOdd, setEvenAndOdd] = useState(s.evenAndOdd);

  const apply = () =>
    onApply((d) => {
      const p = PAPERS.find((x) => x.name === paper);
      const geom: PageGeom = {
        width: p ? (landscape ? p.height : p.width) : s.width,
        height: p ? (landscape ? p.width : p.height) : s.height,
        marginTop: m.top * PT_PER_CM,
        marginBottom: m.bottom * PT_PER_CM,
        marginLeft: m.left * PT_PER_CM,
        marginRight: m.right * PT_PER_CM,
        headerDistance: m.header * PT_PER_CM,
        footerDistance: m.footer * PT_PER_CM,
      };
      setSectionPage(d, section.index, geom, wholeDoc);
      if (gw) d.defaults = { ...d.defaults, size: gw.size, fontEA: gw.font };
      // 首页不同在各节分别设置；这里修改时应用到所有节
      if (titlePg !== s.titlePg) for (const b of d.body) if (b.kind === 'p' && b.sectionHF) b.sectionHF.titlePg = titlePg;
      d.setup.titlePg = titlePg;
      d.setup.evenAndOdd = evenAndOdd;
      // 文档网格全文统一，按本对话框中的版心计算
      const L = Math.max(1, lines);
      const C = Math.max(1, chars);
      const g = { ...d.setup, ...geom };
      d.setup.grid = {
        type: grid,
        linePitch: grid === 'none' ? d.setup.grid.linePitch : +(textHeight(g) / L).toFixed(2),
        charSpace: grid === 'linesAndChars' ? Math.round((textWidth(g) / C - d.defaults.size) * 4096) : 0,
      };
    });

  // 党政机关公文格式（GB/T 9704-2012）：A4，天头 37mm，订口 28mm，版心 156×225mm，每面 22 行、每行 28 字，三号仿宋
  const gongwen = () => {
    setPaper('A4');
    setLandscape(false);
    setM({ top: 3.7, bottom: 3.5, left: 2.8, right: 2.6, header: 1.5, footer: 2.8 });
    setGrid('linesAndChars');
    setLines(22);
    setChars(28);
    setGw({ size: 16, font: '仿宋_GB2312' });
  };
  const wordDefault = () => {
    setPaper('A4');
    setLandscape(false);
    setM({ top: 2.54, bottom: 2.54, left: 3.18, right: 3.18, header: 1.5, footer: 1.75 });
    setGrid('lines');
    setLines(44);
    setChars(39);
    setGw({ size: 10.5, font: '宋体' });
  };

  const num = (k: keyof typeof m, label: string) => (
    <label className="label">
      {label}（厘米）
      <input className="field" type="number" step="0.01" min="0" value={m[k]} onChange={(e) => setM({ ...m, [k]: Number(e.target.value) })} />
    </label>
  );

  return (
    <div className="backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="dialog" role="dialog" aria-modal="true" aria-label="页面设置" style={{ width: 560 }}>
        <h3>页面设置</h3>
        <div style={{ display: 'flex', gap: 8 }}>
          <button className="btn" onClick={wordDefault}>Word 中文默认</button>
          <button className="btn" onClick={gongwen}>公文格式（GB/T 9704）</button>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 12 }}>
          <label className="label">
            纸张
            <select className="field" value={paper} onChange={(e) => setPaper(e.target.value)}>
              {PAPERS.map((p) => (
                <option key={p.name}>{p.name}</option>
              ))}
              {paper === '自定义' && <option>自定义</option>}
            </select>
          </label>
          <label className="label">
            方向
            <select className="field" value={landscape ? 'l' : 'p'} onChange={(e) => setLandscape(e.target.value === 'l')}>
              <option value="p">纵向</option>
              <option value="l">横向</option>
            </select>
          </label>
          {sectionCount > 1 && (
            <label className="label" style={{ gridColumn: '1 / -1' }}>
              纸张与页边距应用于
              <select className="field" value={wholeDoc ? 'doc' : 'sect'} onChange={(e) => setWholeDoc(e.target.value === 'doc')}>
                <option value="sect">{`本节（第 ${sectionList(doc).findIndex((x) => x.index === section.index) + 1} 节）`}</option>
                <option value="doc">整篇文档</option>
              </select>
            </label>
          )}
          {num('top', '上边距')}
          {num('bottom', '下边距')}
          {num('left', '左边距')}
          {num('right', '右边距')}
          {num('header', '页眉距边界')}
          {num('footer', '页脚距边界')}
          <label className="label">
            文档网格
            <select className="field" value={grid} onChange={(e) => setGrid(e.target.value as GridType)}>
              <option value="none">无网格</option>
              <option value="lines">只指定行网格</option>
              <option value="linesAndChars">指定行和字符网格</option>
            </select>
          </label>
          <label className="label">
            每页行数
            <input className="field" type="number" min="1" max="100" value={lines} disabled={grid === 'none'} onChange={(e) => setLines(Number(e.target.value))} />
          </label>
          <label className="label">
            每行字数
            <input className="field" type="number" min="1" max="100" value={chars} disabled={grid !== 'linesAndChars'} onChange={(e) => setChars(Number(e.target.value))} />
          </label>
          {gw && (
            <div className="label">
              默认文字
              <span style={{ color: 'var(--text)', paddingTop: 10 }}>
                {gw.font} {CN_SIZES.find(([, v]) => v === gw.size)?.[0] ?? gw.size}
              </span>
            </div>
          )}
        </div>
        <div className="label">
          页眉页脚
          <div style={{ display: 'flex', gap: 18, color: 'var(--text)' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="checkbox" checked={titlePg} onChange={(e) => setTitlePg(e.target.checked)} />
              首页不同
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="checkbox" checked={evenAndOdd} onChange={(e) => setEvenAndOdd(e.target.checked)} />
              奇偶页不同
            </label>
          </div>
        </div>
        <div className="dialog-foot">
          <button className="btn" onClick={onClose}>取消</button>
          <button className="btn primary" onClick={apply}>确定</button>
        </div>
      </div>
    </div>
  );
}

/** 批注栏：按在文档里的位置排列，每条带被批注的文字和回复 */
function CommentsPanel({
  doc,
  ranges,
  onPick,
  onReply,
  onDelete,
}: {
  doc: Doc;
  ranges: Map<string, [Pos, Pos]>;
  onPick: (id: string) => void;
  onReply: (id: string) => void;
  onDelete: (id: string) => void;
}) {
  const all = doc.comments ?? [];
  const roots = all
    .filter((c) => !c.parent)
    .sort((a, b) => {
      const ra = ranges.get(a.id);
      const rb = ranges.get(b.id);
      if (!ra || !rb) return ra ? -1 : rb ? 1 : 0;
      return comparePos(ra[0], rb[0]);
    });
  const when = (d: string) => (d ? d.slice(0, 16).replace('T', ' ') : '');
  const quote = (id: string) => {
    const r = ranges.get(id);
    if (!r) return '';
    const t = selectedText(doc, r[0], r[1]).replace(/\s+/g, ' ');
    return t.length > 40 ? `${t.slice(0, 40)}…` : t;
  };
  return (
    <aside className="word-comments" aria-label="批注">
      {roots.map((c) => (
        <div key={c.id} className="word-comment" onClick={() => onPick(c.id)}>
          <div className="word-comment-head">
            <strong>{c.author}</strong>
            <span className="sub">{when(c.date)}</span>
          </div>
          {quote(c.id) && <div className="word-comment-quote">“{quote(c.id)}”</div>}
          <div className="word-comment-text">{c.text}</div>
          {all
            .filter((x) => x.parent === c.id)
            .map((x) => (
              <div key={x.id} className="word-comment-reply">
                <div className="word-comment-head">
                  <strong>{x.author}</strong>
                  <span className="sub">{when(x.date)}</span>
                  <span className="grow" />
                  <button className="link" onClick={(e) => (e.stopPropagation(), onDelete(x.id))}>
                    删除
                  </button>
                </div>
                <div className="word-comment-text">{x.text}</div>
              </div>
            ))}
          <div className="word-comment-actions">
            <button className="link" onClick={(e) => (e.stopPropagation(), onReply(c.id))}>
              回复
            </button>
            <button className="link" onClick={(e) => (e.stopPropagation(), onDelete(c.id))}>
              删除
            </button>
          </div>
        </div>
      ))}
    </aside>
  );
}

/** 面板里的字段（各列同名的合并）→ 给 AI 的字段列表：说明里写明在每一列（采购 / 销售）的含义 */
function paletteFields(palette: FieldPalette): { key: string; label: string }[] {
  const map = new Map<string, string[]>();
  for (const col of palette.columns)
    for (const g of col.groups)
      for (const f of g.fields) {
        const list = map.get(f.key) ?? map.set(f.key, []).get(f.key)!;
        list.push(`${col.title}·${g.title}${f.label && f.label !== f.key ? `·${f.label}` : ''}`);
      }
  return [...map].map(([key, where]) => ({ key, label: where.join('；') }));
}

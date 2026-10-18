/**
 * 宿主接口：编辑器只通过它使用工作台的能力（读写文件、提示、标签页状态等），不直接引用工作台的 store / api。
 * 工作台启动时调用 setEditorHost 提供实现；编辑器里用 host() 取得。测试环境可以传一个假实现。
 */

/** 编辑器打开的文件（标签页） */
export interface EditorTab {
  id: string;
  title: string;
  /** 新建、还没保存过的文件没有路径 */
  path?: string;
}

import type { BundleItem } from './container';
import type { ColumnsFile, FormDef, SheetColumns } from './schema';

/**
 * 当前文件所在的工作包（.wbp）：只有包内的文件才有（`host().bundleOf(path)`），包外为 null。
 * 列属性、表单的界面只在拿到它时才出现。列属性、表单存在包里，不写进 xlsx。
 */
export interface BundleContext {
  /** 当前登录身份对这个条目能不能写（工作包设了权限且只给了“可读”时为 false） */
  readonly writable?: boolean;
  /** 工作包文件的路径 */
  readonly pack: string;
  readonly itemId: string;
  readonly itemName: string;
  /** 工作包的 id（UUID）和名称 */
  readonly packId: string;
  readonly packTitle: string;
  /** 本条目的列属性（每次返回新对象，可随意修改） */
  columns(): ColumnsFile;
  /** 设置某张工作表的列属性（null = 清除）；只改包的内存状态，随下一次保存条目（或 flush）写盘 */
  setColumns(sheet: string, cols: SheetColumns | null): void;
  /** 目标是本条目的表单 */
  forms(): FormDef[];
  /** 保存 / 删除表单，立即写盘 */
  saveForm(form: FormDef): Promise<void>;
  deleteForm(id: string): Promise<void>;
  newFormId(): string;
  /** 记一条表单提交记录（forms/<表单id>.log.jsonl，只追加；随包的下一次保存写盘） */
  logSubmission(formId: string, record: Record<string, unknown>): void;
  /** 包里的所有条目（表单“提交后运行”选脚本等） */
  items(): BundleItem[];
  /** 往包里加一个文件（新条目，同名自动改名），立即写盘；返回新条目 */
  addFile(name: string, data: Uint8Array): Promise<BundleItem>;
  /** 读 `[条目名]工作表!A2:A20` 里的文字（单选 / 多选的选项取自包里别的表格）；引用不对返回空数组 */
  readRange(ref: string): string[];
  /** 文本类条目（md / html / 脚本）的内容；不是文本或不存在返回 undefined */
  readText(itemId: string): string | undefined;
  /** 把包的当前状态写盘 */
  flush(): Promise<void>;
  /** 包的内容变了（别处改了列属性 / 表单）时回调 */
  subscribe(cb: () => void): () => void;
}

/** 旧版 .xls 工作簿（后端 calamine 解析）的一张表 */
export interface XlsSheet {
  name: string;
  /** 第一个单元格的 [行, 列] */
  start: [number, number];
  rows: XlsCell[][];
  /** [r0, c0, r1, c1] */
  merges: [number, number, number, number][];
}

/** `read_xls` 的单元格：null / 数值 / 文字 / 逻辑值 / 日期（Excel 序列号） */
export type XlsCell = null | number | string | boolean | { d: number };

export interface DirEntry {
  name: string;
  path: string;
  size: number;
  /** 修改时间 YYYY-MM-DD HH:MM */
  modified: string | null;
  /** 是文件夹（工作台的 list_dir 总会给出） */
  isDir?: boolean;
}

export interface EditorHost {
  /**
   * 运行在哪里（docs/ANDROID.md §7）：缺省 = 桌面。'mobile' = 安卓手机版：编辑器换手机布局（不显示功能区 / 标尺 / 状态栏，
   * 默认查看模式、按屏宽缩放、底部工具条），文件路径是虚拟路径（saf:… / app:/…），basename 仍是文件名。
   */
  platform?: 'desktop' | 'mobile';
  // ---- 文件 ----
  readBytes(path: string): Promise<Uint8Array>;
  /** 读文本文件：自动识别编码（BOM → UTF-8 → GBK 等），返回内容和编码，保存时按原编码写回 */
  readText(path: string): Promise<{ content: string; encoding: string }>;
  /** 只能写工作文件夹里白名单扩展名的文件 */
  writeBytes(path: string, data: Uint8Array): Promise<void>;
  /** 这个文件能不能原地保存（在工作文件夹里，或是用户自己打开的文件）；没有这个方法时只认工作文件夹 */
  canWrite?(path: string): Promise<boolean>;
  /** 按指定编码写文本（csv 保持原编码） */
  writeText(path: string, text: string, encoding: string): Promise<void>;
  listDir(dir: string): Promise<DirEntry[]>;
  readXls(path: string): Promise<XlsSheet[]>;
  /** 已安装字体的度量表：每个字体 { 表名: base64 }（name / hhea / OS/2 / head）；拿不到时（浏览器里）返回空数组 */
  fontFaces?(): Promise<Record<string, string>[]>;
  openWithSystem(path: string): Promise<void>;
  openUrl(url: string): Promise<void>;
  /** 用户添加的工作文件夹（保存位置只能在这些文件夹里） */
  workFolders(): string[];
  // ---- 界面 ----
  toast(msg: string, ms?: number): void;
  /** 输入框；取消返回 null */
  askText(title: string, value?: string): Promise<string | null>;
  /** 系统“另存为”对话框（可保存到任意位置）；filters 为 [名称, 扩展名列表]，第一个是默认类型；取消返回 null。浏览器开发页没有 */
  pickSavePath?(title: string, name: string, dir: string | null, filters: [string, string[]][]): Promise<string | null>;
  /**
   * 登记标签页的“保存”：关闭有未保存修改的标签时，工作台会问“保存 / 不保存 / 取消”，选“保存”就调用它
   * （没有路径的新文件在这里弹“另存为”）；返回 true = 已保存，false = 没保存（如取消了另存为），标签不关。返回注销函数。
   */
  registerSave?(tabId: string, save: () => Promise<boolean>): () => void;
  /** 确认框（应用内弹窗，WebView 里 window.confirm 不一定弹出）；确定 true，取消 false */
  confirm(message: string, ok?: string): Promise<boolean>;
  /** 执行一个可能失败的操作：失败时提示错误并返回 undefined */
  guard<T>(p: Promise<T>): Promise<T | undefined>;
  /** 弹出层打开 / 关闭（工作台据此暂时隐藏会盖住弹出层的网页子窗口） */
  pushPopup(): void;
  popPopup(): void;
  /** 改标签页：未保存标记、另存后的路径和标题 */
  updateTab(id: string, patch: { dirty?: boolean; path?: string; title?: string }): void;
  /** 打印页面图片（data URL） */
  printImages(urls: string[], pageSize?: string): Promise<void>;
  log(level: 'info' | 'warn' | 'error', source: string, msg: string): void;
  // ---- 随工作台状态变化的值（React hook，只能在组件里调用） ----
  useWorkFolders(): string[];
  /** 外部改动了文件：tick 每次加一，dirs 为受影响的目录 */
  useFsEvent(): { tick: number; dirs: string[] };
  /** 当前登录用户名（修订、批注作者）；未登录为 null */
  useUserName(): string | null;
  /** 设置当前用户名（手机版：审阅修订、批注的作者，第一次审阅时让用户填）；没有这个方法时不能改 */
  setUserName?(name: string): void;
  /** 用这个 Word 文档当模板开始套打（手机版的套打向导，docs/ANDROID.md §6.6）；没有这个方法时不显示入口 */
  startMerge?(templatePath: string): void;
  // ---- 可选能力 ----
  /** 待办关联（白板记着自己属于哪条待办）：该待办的标题，找不到（已删除）返回 null（React hook；id 缺省时返回 null） */
  useTodoTitle?(id: number | undefined): string | null;
  /** 打开这条待办的编辑窗口 */
  openTodo?(id: number): void;
  /** 新建一个同种类的空白文件（在新标签页里打开）；kind 是编辑器的 wz.kind */
  newDocument?(kind: string): void;
  /** 弹出文件选择框，选中的文件在工作台里打开（按该编辑器的扩展名过滤）；取消什么也不做 */
  openDocument?(kind: string): Promise<void>;
  /** 当前文件所在的工作包；不在包里返回 null（编辑器据此决定是否显示列属性 / 表单界面） */
  bundleOf?(path: string): BundleContext | null;
  /** 在工作台里打开一个路径（可指定编辑器种类和标签标题）；包里打开条目用 */
  openPath?(path: string, kind?: string, title?: string): void;
  /** 把一个标签页换成另一种编辑器打开（可同时换路径）：只含白板的工作包直接进入白板状态 */
  retarget?(tabId: string, kind: string, path?: string): void;
  /** 关闭一个标签页（不询问；新建的工作包改在新窗口里打开后，关掉空白标签用） */
  closeTab?(tabId: string): void;
  /** 弹出文件选择框（可多选），返回选中的路径；取消返回空数组。exts 是扩展名过滤（不带点），缺省为全部文件 */
  pickFiles?(title: string, exts?: string[]): Promise<string[]>;
  /** 选一个文件夹（系统对话框）；取消返回 null */
  pickFolder?(title: string): Promise<string | null>;
  /**
   * 把一整页 HTML 放到隔离的页面协议上，返回 iframe 可加载的地址（页面自带“禁止联网”的 CSP，可运行内联脚本）。
   * 没有（浏览器开发页）时，调用方退回 srcdoc。卡片包运行页面用。
   */
  servePage?(html: string): Promise<string>;
  /**
   * 按计划填 Excel 模板（导出规则的“预览生成”用；计划是 @wz/card 的 FillPlan，这里不引用它的类型）：
   * 工作台用表格编辑器的引擎实现（src/lib/xlsxTemplate.ts）。
   */
  fillXlsxTemplate?(template: Uint8Array, plan: unknown): Promise<Uint8Array>;
  /**
   * 最近一次在资源管理器“显示发票”里识别出的发票（导出规则预览用真实数据）；还没有时 null。
   * 发票对象的字段同 @wz/card 的 InvoiceRecord（file number date project amount tax total pages buyer seller remark check）。
   */
  recentInvoices?(): { folder: string; path: string; at: string; invoices: Record<string, unknown>[] } | null;
  /** 名字的拼音首字母（小写，如 张林雄 → zlx） */
  initialsOf?(name: string): string;
  /**
   * 由工作台后端发一个 HTTP 请求（只允许 http / https，GET / POST，有超时和大小上限）。卡片包的“网络接口”数据源用：
   * 地址来自卡片包里声明的数据源，页面本身不能联网。没有（浏览器开发页）时网络接口数据源不可用。
   */
  httpRequest?(req: { url: string; method: string; headers: Record<string, string>; params: Record<string, string>; body?: string; timeoutMs?: number }): Promise<{
    status: number;
    ok: boolean;
    contentType: string;
    body: string;
    url: string;
  }>;
  /** 导出：弹出文件夹选择框，返回所选文件夹（取消返回 null）。选到的文件夹之后可以用 exportFile 写入 */
  /** 加密的工作包：用解压密钥读出（返回不加密的 zip 字节） */
  readPack?(path: string, key: string): Promise<Uint8Array>;
  /** 加密的工作包：把不加密的 zip 字节加密后整包写盘 */
  writePack?(path: string, plain: Uint8Array, key: string): Promise<void>;
  pickExportDir?(title: string): Promise<string | null>;
  /** 导出一个文件到 pickExportDir 选过的文件夹；从不覆盖，同名时自动改名，返回实际写入的文件名 */
  exportFile?(dir: string, name: string, data: Uint8Array): Promise<string>;
  /** 运行包内脚本（工作台提供沙箱和权限确认；见 docs/SHEET_COLUMNS.md 脚本一节） */
  /** 复制文字但不进任何剪贴板历史（工作台的剪切板菜单、Windows 剪贴板历史都不记）；没有这个能力时调用方不要退回普通复制 */
  copyPrivate?(text: string): Promise<void>;
  /** 悬浮显示：在这个标签页上框选显示区，浮成置顶的只读浏览小窗口（编辑器工具条里的 FloatViewButton 调用） */
  floatView?(tabId: string): void;
  runScript?(pack: string, itemId: string, source: string, hooks: ScriptHooks, opts?: ScriptOptions): ScriptRun;
  /** 模版字段面板：这个文件是工作台里的模版时，返回可插入的字段（Word 编辑器在右侧显示“插入字段”面板）；否则返回 null */
  fieldPalette?(path: string): FieldPalette | null;
  /** AI 动作（只有 AI 版提供；没有时编辑器不显示 AI 入口，编辑器不需要知道版本） */
  ai?: EditorAi;
}

/** 模版字段面板：按列（如 采购 / 销售）→ 分类（如 公司信息）→ 字段；插入的文字是 {{key}} */
export interface FieldPalette {
  title?: string;
  columns: { title: string; groups: { title: string; fields: { key: string; label?: string }[] }[] }[];
}

/** 脚本运行时的回调：输出、进度 */
export interface ScriptHooks {
  log(level: 'log' | 'warn' | 'error', text: string): void;
}

/** 运行脚本的附加选项（接口 2，见 docs/WBP_FORMAT.md 第 8 节） */
export interface ScriptOptions {
  /** 脚本里用 `wbp.input` 读到的内容（表单校验时是这条记录；流程步骤时是 { flow, args }） */
  input?: Record<string, unknown>;
  /** validate：只读校验，脚本不能改任何内容（返回值按 scriptErrors 解释）；缺省 run */
  mode?: 'run' | 'validate';
  /** 要求包里打开的标签页都已保存（脚本要读到最新数据时用），否则不运行 */
  requireSaved?: boolean;
}

export interface ScriptRun {
  /** 结束：返回值（JSON 文字）或抛出错误；outputs 是脚本新生成的文件名 */
  done: Promise<{ value: string | undefined; ms: number; outputs?: string[] }>;
  stop(): void;
}

/**
 * 校验脚本的返回值 → 错误列表（空 = 通过）：
 * 没有返回值 / true / null / { ok: true } 通过；false 不通过；字符串或字符串数组是错误信息；{ errors: [...] } 或 { ok: false, message } 同理。
 */
export function scriptErrors(value: string | undefined): string[] {
  if (value === undefined) return [];
  let v: unknown;
  try {
    v = JSON.parse(value);
  } catch {
    return value.trim() ? [value.trim()] : [];
  }
  const texts = (x: unknown): string[] => (Array.isArray(x) ? x.map((e) => String(e)).filter(Boolean) : typeof x === 'string' && x ? [x] : []);
  if (v === null || v === true) return [];
  if (v === false) return ['校验没有通过'];
  if (typeof v === 'string' || Array.isArray(v)) return texts(v);
  if (typeof v === 'object') {
    const o = v as { ok?: unknown; errors?: unknown; message?: unknown };
    const errs = [...texts(o.errors), ...texts(o.message)];
    if (errs.length) return errs;
    if (o.ok === false) return ['校验没有通过'];
  }
  return [];
}

/** 编辑器里 AI 动作的作用对象 */
export type AiWhere = 'text-selection';

/** AI 结果写回的方式：替换选中内容 / 插入到选中内容后面（另起一段） */
export type AiApplyMode = 'replace' | 'insert-after';

export interface AiActionInfo {
  id: string;
  label: string;
  /** 鼠标提示 */
  hint?: string;
}

export interface EditorAi {
  /** 某类对象上可用的 AI 动作 */
  actions(where: AiWhere): Promise<AiActionInfo[]>;
  /**
   * 运行动作：工作台弹出预览框（流式显示结果，可修改、重试、复制），用户选了写回方式才返回结果；
   * 取消、出错或只复制时返回 null。modes：编辑器支持的写回方式（预览框只显示这些按钮）。
   */
  run(actionId: string, input: string, modes: AiApplyMode[]): Promise<{ text: string; mode: AiApplyMode } | null>;
  /**
   * 单轮文字生成（系统提示 + 用户内容 → 回复全文）：编辑器自己管界面和流程，
   * 只借工作台的模型调用。onText 每收到一段文字回调一次；取消时 signal 中止并抛出 AbortError。
   */
  complete(system: string, user: string, opts?: { signal?: AbortSignal; onText?: (delta: string) => void }): Promise<string>;
  /**
   * 打开 AI 助手侧栏（对话 + 让 AI 直接修改当前文档）。同一时间只有一个助手；再次打开会换成新的文档。
   * 返回关闭函数：编辑器卸载（关标签页）时调用。没有这个方法的版本不显示“AI 助手”按钮。
   */
  openAssistant?(target: AiAssistantTarget): () => void;
  /**
   * 合同 → 模版（“插入字段”面板的“AI 生成”）：读文档的段落，按工作台的技能文件判断哪些文字是合同数据、该换成哪个字段。
   * 只返回建议，编辑器自己在文档里修改（记为修订）。没有这个方法的版本不显示按钮。
   */
  templateFields?(input: TemplateFieldsInput, opts?: { signal?: AbortSignal; onText?: (delta: string) => void }): Promise<TemplateFieldsResult>;
}

export interface TemplateFieldsInput {
  /** 文件名（帮助判断合同类型） */
  title: string;
  /** 有文字的段落，下标就是建议里的 p */
  paragraphs: string[];
  /** 可用的字段（来自 FieldPalette，同名字段合并：label 里写明在各列的含义） */
  fields: { key: string; label: string }[];
}

export interface TemplateFieldsResult {
  /** p = 段落下标；text = 要换掉的原文；after = 原文没有值时在这段标签文字后面插入 */
  hits: { p?: number; text?: string; after?: string; field: string }[];
  /** AI 判断的合同类型（采购 / 销售）和需要人工核对的地方 */
  notes: string;
}

/**
 * 交给 AI 助手的文档。api 按 kind 约定（工作台里对应的 AI 技能知道怎么用）：
 * - 'word'：WordAiApi
 * - 'sheet'：SheetAiApi
 */
export interface AiAssistantTarget {
  kind: 'word' | 'sheet';
  /** 标签页标题（侧栏显示） */
  title: string;
  api: WordAiApi | SheetAiApi;
}

/** Word 编辑器提供给 AI 助手的接口；doc 是编辑器自己的文档模型（core/model.ts 的 Doc） */
export interface WordAiApi {
  doc(): unknown;
  /** 当前选中的文字（没有选中为空串） */
  selectionText(): string;
  /** 一次修改：fn 原地修改文档；成功记一步撤销并刷新；fn 抛错时恢复到修改前（不留半截）并把错误抛回 */
  edit(label: string, fn: (doc: unknown) => void): void;
}

/** 表格编辑器提供给 AI 助手的接口；book / engine 是编辑器自己的模型（core/workbook.ts 的 Workbook、公式引擎 Engine） */
export interface SheetAiApi {
  book(): unknown;
  engine(): unknown;
  /** 当前工作表名 */
  activeSheet(): string;
  /** 当前选区（A1 引用，如 B2:D10） */
  selectionRef(): string;
  /** 把一次已经执行完的修改（core/ops.ts 的 Entry）记入撤销栈并刷新界面、标记未保存 */
  record(entry: unknown): void;
}

let current: EditorHost | null = null;

export function setEditorHost(h: EditorHost): void {
  current = h;
}

/** 安卓手机版（host().platform === 'mobile'）；还没设置宿主时算桌面 */
export const isMobile = (): boolean => current?.platform === 'mobile';

export function host(): EditorHost {
  if (!current) throw new Error('编辑器宿主还没有设置（setEditorHost）');
  return current;
}

/** 编辑器描述：写在编辑器包 package.json 的 wz 字段里，工作台的注册表据此决定哪些文件交给哪个编辑器 */
export interface EditorDescriptor {
  /** 标签页种类，唯一 */
  kind: string;
  /** 设置里“打开方式”分组显示的名字 */
  label: string;
  /** 能打开的扩展名（小写，不带点） */
  exts: string[];
  /** 排序（设置里“打开方式”的显示顺序；两个编辑器认领同一扩展名时小的优先） */
  order?: number;
  /** 小工具列表里的徽标文字（2–4 个字母），如 DOC */
  tag?: string;
  /** 小工具列表里“打开文件”一项的名字，如“打开 Word 文档” */
  openName?: string;
  /** 能新建空白文件时：小工具里的名字和新标签页的标题 */
  create?: { name: string; title: string };
}

/**
 * “另存为 / 导出”选保存位置：有系统保存对话框就用它（初始文件夹 `dir`，缺省为第一个工作文件夹），
 * 没有（浏览器开发页）就退回输入文件名、保存到 `dir`。`exts` 第一个是默认扩展名，名字没写扩展名时补上。取消返回 null。
 */
export async function pickSave(o: { title: string; name: string; dir?: string; label: string; exts: string[] }): Promise<string | null> {
  const h = host();
  const dir = o.dir ?? h.workFolders()[0];
  if (h.pickSavePath) return h.pickSavePath(o.title, o.name, dir ?? null, [[o.label, o.exts]]);
  if (!dir) {
    h.toast('请先在左侧添加一个工作文件夹，文件将保存在其中');
    return null;
  }
  const name = await h.askText(`${o.title}（保存到 ${basename(dir)}）`, o.name);
  if (!name) return null;
  const fixed = o.exts.includes(extOf(name)) ? name : `${name}.${o.exts[0]}`;
  const sep = dir.includes('\\') ? '\\' : '/';
  return `${dir}${sep}${fixed}`;
}

/** 路径工具（Windows / POSIX 分隔符都认） */
export function basename(p: string): string {
  return p.split(/[\\/]/).pop() ?? p;
}

export function dirname(p: string): string {
  const i = Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\'));
  return i > 0 ? p.slice(0, i) : p;
}

export function extOf(name: string): string {
  const base = basename(name);
  const i = base.lastIndexOf('.');
  return i > 0 ? base.slice(i + 1).toLowerCase() : '';
}

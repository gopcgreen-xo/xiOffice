/**
 * 双色图标：主体用 currentColor（跟随文字色），重点部分用强调色。
 * 强调色取 --ico-accent，默认跟随主题 --accent；按钮不可用时由样式把 --ico-accent 设回 currentColor，整体变灰。
 */
const A = 'var(--ico-accent, var(--accent, #0f6cbd))';

const PATHS = {
  plus: <path d="M8 3v10M3 8h10" />,
  close: <path d="M4 4l8 8M12 4l-8 8" />,
  minimize: <path d="M3 8h10" />,
  maximize: <rect x="3.5" y="3.5" width="9" height="9" rx="0.5" strokeWidth={1.1} />,
  unmaximize: (
    <g strokeWidth={1.1}>
      <rect x="3" y="5.5" width="7.5" height="7.5" rx="0.5" />
      <path d="M5.5 5.5V3.5a.5.5 0 0 1 .5-.5h6.5a.5.5 0 0 1 .5.5V10a.5.5 0 0 1-.5.5h-2" />
    </g>
  ),
  drop: (
    <>
      <path d="M8 1.8C6.4 3.9 3.6 7.3 3.6 10a4.4 4.4 0 0 0 8.8 0c0-2.7-2.8-6.1-4.4-8.2z" fill="currentColor" />
      <path d="M5.9 10.3a2.2 2.2 0 0 0 1.8 2" stroke="#fff" strokeOpacity={0.85} strokeWidth={1.2} />
    </>
  ),
  expand: <path d="M9.5 2.5h4v4M13.5 2.5L9 7M6.5 13.5h-4v-4M2.5 13.5L7 9" />,
  tray: <path d="M2.5 9.5v4h11v-4M8 2v8M5 7l3 3 3-3" />,
  trash: (
    <>
      <path d="M3 4.5h10M6.5 4.5V3h3v1.5M4.5 4.5l.6 8.5h5.8l.6-8.5" />
      <path d="M7 7v3.5M9 7v3.5" stroke={A} />
    </>
  ),
  edit: <path d="M10.5 2.5l3 3L6 13H3v-3z" />,
  folder: (
    <>
      <path d="M1.5 4a1 1 0 0 1 1-1h3.6l1.5 1.6h5.9a1 1 0 0 1 1 1v7.4a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1z" />
      <path d="M1.5 6.8h13" stroke={A} />
    </>
  ),
  plusCircle: (
    <>
      <circle cx="8" cy="8" r="6" />
      <path d="M8 5.2v5.6M5.2 8h5.6" stroke={A} />
    </>
  ),
  sort: (
    <>
      <path d="M4.5 13V3M2.2 5.3 4.5 3l2.3 2.3" />
      <path d="M11.5 3v10M9.2 10.7l2.3 2.3 2.3-2.3" stroke={A} />
    </>
  ),
  viewLines: (
    <>
      <path d="M5.5 4h8M5.5 8h8M5.5 12h8" />
      <g fill={A} stroke="none">
        <circle cx="2.6" cy="4" r="1.1" />
        <circle cx="2.6" cy="8" r="1.1" />
        <circle cx="2.6" cy="12" r="1.1" />
      </g>
    </>
  ),
  more: (
    <g fill="currentColor" stroke="none">
      <circle cx="3.5" cy="8" r="1.1" />
      <circle cx="8" cy="8" r="1.1" />
      <circle cx="12.5" cy="8" r="1.1" />
    </g>
  ),
  chevron: <path d="M6 4l4 4-4 4" />,
  back: (
    <>
      <path d="M13.5 8H3" />
      <path d="M7 3.5 2.5 8 7 12.5" stroke={A} />
    </>
  ),
  forward: (
    <>
      <path d="M2.5 8H13" />
      <path d="M9 3.5 13.5 8 9 12.5" stroke={A} />
    </>
  ),
  history: (
    <>
      <circle cx="8" cy="8" r="5.5" />
      <path d="M8 4.8V8l2.2 1.4" />
    </>
  ),
  refresh: (
    <>
      <path d="M13 8a5 5 0 1 1-1.5-3.5" />
      <path d="M13 2.5v3h-3" stroke={A} />
    </>
  ),
  explorer: (
    <>
      <rect x="2" y="2.5" width="12" height="11" rx="1.5" />
      <path d="M2 5.5h12M5 8h6M5 10.5h4" />
    </>
  ),
  palette: (
    <>
      <path d="M8 1.8a6.2 6.2 0 1 0 0 12.4c1 0 1.4-.6 1.2-1.4-.3-1 .3-1.9 1.3-1.9h1.6a2.3 2.3 0 0 0 2.3-2.4A6.4 6.4 0 0 0 8 1.8z" />
      <circle cx="5" cy="7" r=".8" />
      <circle cx="7.5" cy="4.8" r=".8" />
      <circle cx="10.6" cy="5.6" r=".8" />
    </>
  ),
  info: (
    <>
      <circle cx="8" cy="8" r="6" />
      <path d="M8 7.2V11M8 5v.1" />
    </>
  ),
  search: (
    <>
      <circle cx="7" cy="7" r="4.5" />
      <path d="M10.5 10.5L14 14" />
    </>
  ),
  globe: (
    <>
      <circle cx="8" cy="8" r="6.5" />
      <path d="M1.5 8h13" />
      <path d="M8 1.5c2 2 2 11 0 13M8 1.5c-2 2-2 11 0 13" stroke={A} />
    </>
  ),
  mail: (
    <>
      <rect x="1.5" y="3" width="13" height="10" rx="1.5" />
      <path d="M2 4l6 5 6-5" />
    </>
  ),
  /** 左侧栏“常用”：四宫格，左上一格强调色实心 */
  navCommon: (
    <>
      <rect x="2" y="2" width="5.2" height="5.2" rx="1.3" fill={A} stroke={A} />
      <rect x="8.8" y="2" width="5.2" height="5.2" rx="1.3" />
      <rect x="2" y="8.8" width="5.2" height="5.2" rx="1.3" />
      <rect x="8.8" y="8.8" width="5.2" height="5.2" rx="1.3" />
    </>
  ),
  /** 左侧栏“邮箱”：信封 + 强调色封口 */
  navMail: (
    <>
      <rect x="1.5" y="3.2" width="13" height="9.6" rx="1.6" />
      <path d="M2.3 4.4 8 8.8l5.7-4.4" stroke={A} />
    </>
  ),
  /** 左侧栏“数据库”：数据库，顶面强调色实心 */
  navErp: (
    <>
      <path d="M2.5 4v8c0 1.1 2.5 2 5.5 2s5.5-.9 5.5-2V4M2.5 8c0 1.1 2.5 2 5.5 2s5.5-.9 5.5-2" />
      <ellipse cx="8" cy="4" rx="5.5" ry="2" fill={A} stroke={A} />
    </>
  ),
  /** 左侧栏“其它”：三个点，中间一个强调色 */
  navMore: (
    <g stroke="none">
      <circle cx="3.2" cy="8" r="1.7" fill="currentColor" />
      <circle cx="8" cy="8" r="1.7" fill={A} />
      <circle cx="12.8" cy="8" r="1.7" fill="currentColor" />
    </g>
  ),
  /** 发票：锯齿底边的票据 + 强调色 ¥ */
  invoice: (
    <>
      <path d="M3 14V2.5a.5.5 0 0 1 .5-.5h9a.5.5 0 0 1 .5.5V14l-1.25-1-1.25 1-1.25-1L8 14l-1.25-1-1.25 1-1.25-1z" />
      <path d="M5.8 4.3 8 7.1l2.2-2.8M8 7.1v4.4M6 8.3h4M6 10.3h4" stroke={A} strokeWidth={1.3} />
    </>
  ),
  bookmark: <path d="M4 2h8v12l-4-3-4 3z" />,
  star: <path d="M8 1.8l1.9 4 4.3.5-3.2 2.9.9 4.3L8 11.3l-3.9 2.2.9-4.3-3.2-2.9 4.3-.5z" />,
  /** 首页：强调色屋顶 + 线条墙身和拱门（和 tabHome 同一个图案） */
  home: (
    <>
      <path d="M3.2 6.6V13a1.5 1.5 0 0 0 1.5 1.5h6.6a1.5 1.5 0 0 0 1.5-1.5V6.6M6.5 14.5v-3.3a1.5 1.5 0 0 1 3 0v3.3" />
      <path d="M1.5 7.8 8 2l6.5 5.8" stroke={A} />
    </>
  ),
  clipboard: (
    <>
      <path d="M5.5 2.5h-2v12h9v-12h-2" />
      <rect x="5.5" y="1.5" width="5" height="2.5" rx="0.8" />
      <path d="M5.5 7.5h5M5.5 10.5h3.5" />
    </>
  ),
  todo: (
    <>
      <rect x="2.5" y="2" width="11" height="12.5" rx="1.5" />
      <path d="M5 6l1 1 2-2M5 10.5l1 1 2-2M9.5 6.5h2M9.5 11h2" />
    </>
  ),
  // 经典齿轮：8 个齿（齿顶半径 6.7、齿根半径 5）+ 中心圆孔
  gear: (
    <>
      <path d="M6.62 3.19L6.84 1.4L9.16 1.4L9.38 3.19A5 5 0 0 1 10.42 3.63L11.84 2.51L13.49 4.16L12.37 5.58A5 5 0 0 1 12.81 6.62L14.6 6.84L14.6 9.16L12.81 9.38A5 5 0 0 1 12.37 10.42L13.49 11.84L11.84 13.49L10.42 12.37A5 5 0 0 1 9.38 12.81L9.16 14.6L6.84 14.6L6.62 12.81A5 5 0 0 1 5.58 12.37L4.16 13.49L2.51 11.84L3.63 10.42A5 5 0 0 1 3.19 9.38L1.4 9.16L1.4 6.84L3.19 6.62A5 5 0 0 1 3.63 5.58L2.51 4.16L4.16 2.51L5.58 3.63A5 5 0 0 1 6.62 3.19Z" />
      <circle cx="8" cy="8" r="2.2" />
    </>
  ),
  crop: <path d="M4 1.5V12h10.5M1.5 4H12v10.5" />,
  scanText: (
    <g>
      <path d="M2 5V2.5h2.5M11.5 2.5H14V5M14 11v2.5h-2.5M4.5 13.5H2V11" />
      <path d="M5 6h6M5 8h6M5 10h4" />
    </g>
  ),
  clip: <path d="M13 7.5l-5 5a3 3 0 0 1-4.3-4.2l5.4-5.4a2 2 0 0 1 2.8 2.8L6.6 11a1 1 0 0 1-1.4-1.4L10 4.8" />,
  upload: <path d="M8 11V3M4.5 6.5L8 3l3.5 3.5M2.5 11v2.5h11V11" />,
  download: <path d="M8 2.5v8M4.5 7L8 10.5 11.5 7M2.5 11v2.5h11V11" />,
  /** 手形：拖动移动显示区域 */
  hand: <path d="M5.5 8.5V3.8a1 1 0 0 1 2 0V7.5M7.5 7.2V2.8a1 1 0 0 1 2 0v4.6M9.5 7.6V3.8a1 1 0 0 1 2 0V9.5c0 2.6-1.6 4.5-4 4.5-1.8 0-2.8-.9-3.7-2.4L2.3 9.1a1 1 0 0 1 1.6-1.2l1.6 1.8" />,
  /** 鼠标箭头：正常（选择）模式 */
  cursor: <path d="M4 2.5v10l2.8-2.6 1.9 4.1 1.6-.7-1.9-4.1 3.8-.2z" />,
  send: <path d="M8 13V3M4 7l4-4 4 4" />,
  check: <path d="M3.5 8.5l3 3 6-7" />,
  calendar: (
    <>
      <rect x="2" y="3" width="12" height="11" rx="1.5" />
      <path d="M2 6.5h12M5 1.5v3M11 1.5v3" />
    </>
  ),
  external: <path d="M9 2.5h4.5V7M13.5 2.5L7.5 8.5M11.5 9.5v4h-9v-9h4" />,
  /** 终端：黑色实心窗口 + 白色提示符（固定颜色，像终端程序的图标） */
  terminal: (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1.8" fill="#24292f" stroke="#24292f" />
      <path d="M4.5 6.3l2 1.8-2 1.8" stroke="#fff" strokeWidth={1.4} />
      <path d="M8 10.3h3.5" stroke="#fff" strokeWidth={1.4} />
    </>
  ),
  /** 终端里新建标签页：黑色实心窗口（顶部一个标签）+ 白色加号 */
  terminalTab: (
    <>
      <path d="M1.5 4a1.5 1.5 0 0 1 1.5-1.5h3.2l1.2 2h5.6A1.5 1.5 0 0 1 14.5 6v6a1.5 1.5 0 0 1-1.5 1.5H3A1.5 1.5 0 0 1 1.5 12z" fill="#24292f" stroke="#24292f" />
      <path d="M8 7v4.2M5.9 9.1h4.2" stroke="#fff" strokeWidth={1.4} />
    </>
  ),
  /** 链接（复制路径） */
  link: (
    <>
      <path d="M7 9a2.5 2.5 0 0 0 3.6.1l2.2-2.2a2.5 2.5 0 0 0-3.6-3.6l-.9.9" />
      <path d="M9 7a2.5 2.5 0 0 0-3.6-.1L3.2 9.1a2.5 2.5 0 0 0 3.6 3.6l.9-.9" stroke={A} />
    </>
  ),
  /** 压缩包（解压） */
  archive: (
    <>
      <rect x="2.5" y="1.5" width="11" height="13" rx="1.2" />
      <path d="M7.5 1.5v1.5h1.5v1.5h-1.5v1.5h1.5v1.5h-1.5" stroke={A} />
      <rect x="6.8" y="9" width="2.4" height="2.8" rx="0.6" stroke={A} />
    </>
  ),
  /** 转换：两个方向相反的箭头 */
  convert: (
    <>
      <path d="M2.5 5.5h10M10 3l2.5 2.5L10 8" />
      <path d="M13.5 10.5h-10M6 8l-2.5 2.5L6 13" stroke={A} />
    </>
  ),
  /** 全选：虚线框 + 勾 */
  selectAll: (
    <>
      <rect x="2" y="2" width="12" height="12" rx="1.5" strokeDasharray="2 1.8" />
      <path d="M5 8.2l2 2 4-4.4" stroke={A} />
    </>
  ),
  /** 在文件夹中定位文件：文件夹 + 放大镜 */
  reveal: (
    <>
      <path d="M7 13.5H2.5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h3.6l1.5 1.6h5.9a1 1 0 0 1 1 1V7.5" />
      <circle cx="11" cy="11" r="2.2" stroke={A} />
      <path d="M12.6 12.6l1.6 1.6" stroke={A} />
    </>
  ),
  /** 工作包：盒子 */
  pack: (
    <>
      <path d="M8 1.8l5.8 3v6.4L8 14.2l-5.8-3V4.8z" />
      <path d="M2.2 4.8L8 7.8l5.8-3M8 7.8v6.4" stroke={A} />
    </>
  ),
  /** 打开文件：文档 + 箭头 */
  openFile: (
    <>
      <path d="M9.5 1.5H4a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V5z" />
      <path d="M6 10l4-4M7 6h3v3" stroke={A} />
    </>
  ),
  pip: (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" rx="2" />
      <rect x="8" y="7" width="5" height="5" rx="1" />
    </>
  ),
  window: (
    <>
      <rect x="2" y="3" width="12" height="10" rx="1.5" />
      <path d="M2 6h12" />
    </>
  ),
  rotate: <path d="M13 8a5 5 0 1 1-1.5-3.5M13 2.5v3h-3" />,
  rotateLeft: <path d="M3 8a5 5 0 1 0 1.5-3.5M3 2.5v3h3" />,
  printer: (
    <>
      <path d="M4.5 6V2h7v4M4.5 11.5h-2v-5h11v5h-2" />
      <rect x="4.5" y="9.5" width="7" height="4.5" rx="0.5" />
    </>
  ),
  arrowUp: (
    <>
      <path d="M8 13.5V3" />
      <path d="M3.5 7.5 8 3l4.5 4.5" stroke={A} />
    </>
  ),
  arrowDown: <path d="M8 3v10M4 9l4 4 4-4" />,
  user: (
    <>
      <circle cx="8" cy="5.5" r="2.75" />
      <path d="M2.75 14c.6-2.6 2.7-4.25 5.25-4.25S12.65 11.4 13.25 14" />
    </>
  ),
  contacts: (
    <>
      <rect x="3" y="1.75" width="10.5" height="12.5" rx="1.5" />
      <circle cx="8.25" cy="6.25" r="1.75" />
      <path d="M5.5 11.25c.4-1.4 1.5-2.25 2.75-2.25s2.35.85 2.75 2.25M1.75 4.5h2.5M1.75 8h2.5M1.75 11.5h2.5" />
    </>
  ),
  bell: <path d="M4 11V7.5a4 4 0 0 1 8 0V11l1.5 1.5h-11zM6.5 14h3" />,
  sidebarLeft: (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" rx="2" />
      <path d="M6 2.5v11" />
    </>
  ),
  closeAll: (
    <>
      <rect x="1.5" y="4.5" width="10" height="9" rx="1.5" />
      <path d="M4.5 2.5h8a1.5 1.5 0 0 1 1.5 1.5v7M4.5 7l4 4M8.5 7l-4 4" />
    </>
  ),
  /** 贴边（悬浮待办）：一个列表小窗口 + 箭头 → 屏幕边缘（强调色粗线） */
  dockEdge: (
    <>
      <rect x="1.5" y="3" width="7" height="10" rx="1.5" />
      <path d="M3.8 6.2h2.4M3.8 9h2.4M10 8h2.7M11.3 6.4 12.9 8l-1.6 1.6" />
      <path d="M14.7 1.8v12.4" stroke={A} strokeWidth={2} />
    </>
  ),
  sidebarRight: (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" rx="2" />
      <path d="M10 2.5v11" />
    </>
  ),
  database: (
    <>
      <ellipse cx="8" cy="4" rx="5.5" ry="2" />
      <path d="M2.5 4v8c0 1.1 2.5 2 5.5 2s5.5-.9 5.5-2V4M2.5 8c0 1.1 2.5 2 5.5 2s5.5-.9 5.5-2" />
    </>
  ),
  zoomIn: (
    <>
      <circle cx="7" cy="7" r="4.5" />
      <path d="M10.5 10.5L14 14M5 7h4M7 5v4" />
    </>
  ),
  zoomOut: (
    <>
      <circle cx="7" cy="7" r="4.5" />
      <path d="M10.5 10.5L14 14M5 7h4" />
    </>
  ),
  shapeRect: (
    <>
      <rect x="1.8" y="2.5" width="10" height="8" rx="1" />
      <path d="M12.5 9.8v5M10 12.3h5" stroke={A} />
    </>
  ),
  shapeEllipse: (
    <>
      <ellipse cx="6.8" cy="6.5" rx="5.3" ry="4" />
      <path d="M12.5 9.8v5M10 12.3h5" stroke={A} />
    </>
  ),
  polyline: (
    <>
      <path d="M2 12.5l3.5-6 3.5 4 5-7.5" />
      <g fill={A} stroke="none">
        <circle cx="2" cy="12.5" r="1.3" />
        <circle cx="5.5" cy="6.5" r="1.3" />
        <circle cx="9" cy="10.5" r="1.3" />
        <circle cx="14" cy="3" r="1.3" />
      </g>
    </>
  ),
  arrowTool: (
    <>
      <path d="M3 13 12.5 3.5" />
      <path d="M7 3.5h5.5V9" stroke={A} />
    </>
  ),
  pen: (
    <>
      <path d="M10.5 2l2.5 2.5L6 11.5H3.5V9z" />
      <path d="M8.5 14h5.5" stroke={A} />
    </>
  ),
  marker: (
    <>
      <path d="M9.5 2.5l4 4-5 5-4-4zM4.5 7.5 3 12.5l1 1 5-1.5" />
      <path d="M1.5 14.5h6" stroke={A} />
    </>
  ),
  mosaic: (
    <>
      <path d="M2.5 2.5H8V8H2.5zM8 8h5.5v5.5H8z" fill={A} stroke="none" />
      <rect x="2.5" y="2.5" width="11" height="11" rx="1" />
      <path d="M2.5 8h11M8 2.5v11" />
    </>
  ),
  text: (
    <>
      <path d="M2.5 3h8M6.5 3v10" />
      <path d="M12.8 6v8M11.3 6h3M11.3 14h3" stroke={A} />
    </>
  ),
  eraser: (
    <>
      <path d="M6 6l4 4-3 3H4l-1.5-1.5a1.4 1.4 0 0 1 0-2z" fill={A} stroke="none" />
      <path d="M9.5 2.5l4 4-6.5 6.5H4l-1.5-1.5a1.4 1.4 0 0 1 0-2zM7 13h7" />
    </>
  ),
  undo: (
    <>
      <path d="M2.5 6H10a3.5 3.5 0 0 1 0 7H6.5" />
      <path d="M5.5 3 2.5 6l3 3" stroke={A} />
    </>
  ),
  redo: (
    <>
      <path d="M13.5 6H6a3.5 3.5 0 0 0 0 7h3.5" />
      <path d="M10.5 3l3 3-3 3" stroke={A} />
    </>
  ),
  /** 截图合并：两块选区 + 加号 */
  mergeShots: (
    <>
      <rect x="1.5" y="2.5" width="6" height="7" rx="0.8" />
      <rect x="8.5" y="6.5" width="6" height="7" rx="0.8" strokeDasharray="2 1.5" stroke={A} />
      <path d="M4.5 11.5v3M3 13h3" stroke={A} />
    </>
  ),
  /** 左右拼接 */
  joinH: (
    <>
      <rect x="1.5" y="3.5" width="6" height="9" rx="0.8" />
      <rect x="8.5" y="3.5" width="6" height="6" rx="0.8" stroke={A} />
    </>
  ),
  /** 上下拼接 */
  joinV: (
    <>
      <rect x="3.5" y="1.5" width="9" height="6" rx="0.8" />
      <rect x="3.5" y="8.5" width="6" height="6" rx="0.8" stroke={A} />
    </>
  ),
  pin: <path d="M10 1.8l4.2 4.2-1.4.4-2.6 2.6.2 3-1.4 1.4-2.3-3.1L3 14l3.5-3.7-3.1-2.3L4.8 6.6l3 .2 2.6-2.6z" />,
  save: (
    <>
      <path d="M2.5 2.5h9l2 2v9h-11zM5 2.5V6h5V2.5" />
      <path d="M5 13.5v-4h6v4" stroke={A} />
    </>
  ),
  /** 格式刷：Office 经典样子，斜放的刷子（右上蓝灰刷柄、白色金属箍、左下黄色刷毛），固定颜色 */
  brush: (
    <g stroke="#4B5B73" strokeWidth={0.6} strokeLinejoin="round" transform="rotate(45 8 8) translate(8 8) scale(0.92) translate(-8 -8)">
      <rect x="6.2" y="0.6" width="3.6" height="5.4" rx="1" fill="#5F7FA3" />
      <rect x="3.2" y="6" width="9.6" height="2.8" rx="0.7" fill="#E6EAF0" />
      <rect x="3.2" y="8" width="9.6" height="0.8" fill="#B3BCC9" />
      <path d="M3.2 8.8h9.6l-.5 4.6a1.7 1.7 0 0 1-1.7 1.5H5.4a1.7 1.7 0 0 1-1.7-1.5z" fill="#F7C020" />
      <path d="M3.5 11.5h9l-.2 1.9a1.7 1.7 0 0 1-1.7 1.5H5.4a1.7 1.7 0 0 1-1.7-1.5z" fill="#EDA70F" />
    </g>
  ),
  /** 保存：平涂蓝色软盘（银色滑盖、白色标签），固定颜色 */
  saveColor: (
    <g stroke="#173F82" strokeWidth={0.7} strokeLinejoin="round">
      <path d="M2 2.8A.8.8 0 0 1 2.8 2h8.4L14 4.8v8.4a.8.8 0 0 1-.8.8H2.8a.8.8 0 0 1-.8-.8z" fill="#2B63BE" />
      <rect x="4.5" y="2" width="6.2" height="4.2" rx="0.4" fill="#D9DFE9" stroke="#7B8798" />
      <rect x="8.6" y="2.8" width="1.3" height="2.6" rx="0.3" fill="#2B63BE" stroke="none" />
      <rect x="3.8" y="8.4" width="8.4" height="5.6" rx="0.6" fill="#fff" stroke="#173F82" />
      
    </g>
  ),
  /** 新建：白色页面 + 蓝色加号，固定颜色 */
  newDocColor: (
    <g strokeLinejoin="round" strokeWidth={0.9}>
      <path d="M3.5 1.5h6l3 3v10h-9z" fill="#fff" stroke="#5B6B83" />
      <path d="M9.5 1.5v3h3" fill="#DDE3EC" stroke="#5B6B83" />
      <circle cx="11.5" cy="11.5" r="3.6" fill="#2B63BE" stroke="#173F82" />
      <path d="M11.5 9.6v3.8M9.6 11.5h3.8" stroke="#fff" strokeWidth={1.4} strokeLinecap="round" />
    </g>
  ),
  /** 添加文件：蓝色加号，固定颜色 */
  addFileColor: <path d="M8 2.8v10.4M2.8 8h10.4" stroke="#2B63BE" strokeWidth={2.4} strokeLinecap="round" />,
  /** 导出：绿色箭头从托盘里向上送出，固定颜色 */
  exportColor: (
    <g strokeLinejoin="round" strokeLinecap="round">
      <path d="M2.2 9.5v3.2a1 1 0 0 0 1 1h9.6a1 1 0 0 0 1-1V9.5" fill="none" stroke="#5B6B83" strokeWidth={1.3} />
      <path d="M8 10.2V3" stroke="#1F8F55" strokeWidth={1.8} />
      <path d="M4.9 5.6L8 2.4l3.1 3.2z" fill="#1F8F55" stroke="#1F8F55" strokeWidth={1} />
    </g>
  ),
  /** 新增：空白页面（右上角折角），固定颜色 */
  blankDocColor: (
    <g strokeLinejoin="round" strokeWidth={0.9}>
      <path d="M3.5 1.5h6l3 3v10h-9z" fill="#fff" stroke="#5B6B83" />
      <path d="M9.5 1.5v3h3" fill="#DDE3EC" stroke="#5B6B83" />
    </g>
  ),
  /** 表单录入：白色表单页 + 绿色输入行，固定颜色 */
  formColor: (
    <g strokeLinejoin="round" strokeWidth={0.9}>
      <path d="M3 1.8h10a.8.8 0 0 1 .8.8v10.8a.8.8 0 0 1-.8.8H3a.8.8 0 0 1-.8-.8V2.6a.8.8 0 0 1 .8-.8z" fill="#fff" stroke="#5B6B83" />
      <rect x="4.4" y="4" width="7.2" height="2" rx="0.5" fill="#DDE3EC" stroke="#8793A6" strokeWidth={0.6} />
      <rect x="4.4" y="7.2" width="7.2" height="2" rx="0.5" fill="#DDE3EC" stroke="#8793A6" strokeWidth={0.6} />
      <rect x="4.4" y="10.4" width="4.2" height="2" rx="0.5" fill="#2E9E5B" stroke="#1B7040" strokeWidth={0.6} />
    </g>
  ),
  /** 打开：黄色文件夹，固定颜色 */
  openColor: (
    <g strokeLinejoin="round" strokeWidth={0.9}>
      <path d="M1.5 3.5a1 1 0 0 1 1-1h3.8l1.5 1.8h5.7a1 1 0 0 1 1 1V12.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1z" fill="#E0A21A" stroke="#A8700A" />
      <path d="M1.5 7.2a1 1 0 0 1 1-1h11a1 1 0 0 1 1 1l-.6 5.6a1 1 0 0 1-1 .9H3.1a1 1 0 0 1-1-.9z" fill="#FFD25A" stroke="#A8700A" />
    </g>
  ),
  /** 首页标签：和其它标签图标一样“线条 + 一种强调色”（强调色屋顶） */
  tabHome: (
    <>
      <path d="M3.2 6.6V13a1.5 1.5 0 0 0 1.5 1.5h6.6a1.5 1.5 0 0 0 1.5-1.5V6.6M6.5 14.5v-3.3a1.5 1.5 0 0 1 3 0v3.3" />
      <path d="M1.5 7.8 8 2l6.5 5.8" stroke={A} />
    </>
  ),
  // 其它标签栏图标（双色：线条 + 强调色），与 openBlue 同一风格
  tabSettings: (
    <>
      <path d="M6.62 3.19L6.84 1.4L9.16 1.4L9.38 3.19A5 5 0 0 1 10.42 3.63L11.84 2.51L13.49 4.16L12.37 5.58A5 5 0 0 1 12.81 6.62L14.6 6.84L14.6 9.16L12.81 9.38A5 5 0 0 1 12.37 10.42L13.49 11.84L11.84 13.49L10.42 12.37A5 5 0 0 1 9.38 12.81L9.16 14.6L6.84 14.6L6.62 12.81A5 5 0 0 1 5.58 12.37L4.16 13.49L2.51 11.84L3.63 10.42A5 5 0 0 1 3.19 9.38L1.4 9.16L1.4 6.84L3.19 6.62A5 5 0 0 1 3.63 5.58L2.51 4.16L4.16 2.51L5.58 3.63A5 5 0 0 1 6.62 3.19Z" />
      <circle cx="8" cy="8" r="2.2" stroke={A} />
    </>
  ),
  tabMail: (
    <>
      <rect x="1.5" y="3" width="13" height="10" rx="1.5" />
      <path d="M2 4l6 5 6-5" stroke={A} />
    </>
  ),
  tabCompose: (
    <>
      <path d="M8 13H3a1.5 1.5 0 0 1-1.5-1.5v-7A1.5 1.5 0 0 1 3 3h10a1.5 1.5 0 0 1 1.5 1.5V7" />
      <path d="M2 4l6 5 6-5" />
      <path d="M10 14.5l.4-1.9 3.6-3.6 1.5 1.5-3.6 3.6z" stroke={A} />
    </>
  ),
  tabErp: (
    <>
      <rect x="2.5" y="1.5" width="11" height="13" rx="1.2" />
      <path d="M5.5 1.5v13" />
      <path d="M7.75 5h3.5M7.75 8h3.5M7.75 11h2" stroke={A} />
    </>
  ),
  /** 在工作台中打开：双色打开的文件夹（后半灰、前片强调色） */
  openBlue: (
    <>
      <path d="M1.5 12V4a1 1 0 0 1 1-1h3.6l1.5 1.6h4.4a1 1 0 0 1 1 1V7" />
      <path d="M1.5 12l1.6-4.3a1 1 0 0 1 .9-.7h9.6a.7.7 0 0 1 .7.9l-1.6 4.4a1 1 0 0 1-.9.7H2.5a1 1 0 0 1-1-1z" stroke={A} />
    </>
  ),
  copy: (
    <>
      <path d="M3.5 10.5h-1v-8h8v1" stroke={A} />
      <rect x="5.5" y="5.5" width="8" height="8" rx="1.2" />
    </>
  ),
  cut: (
    <>
      <circle cx="4.5" cy="12" r="1.8" />
      <circle cx="11.5" cy="12" r="1.8" />
      <path d="M5.6 10.6 11 2.5M10.4 10.6 5 2.5" stroke={A} />
    </>
  ),
  paste: (
    <>
      <path d="M5.5 3H4.2A1.2 1.2 0 0 0 3 4.2v8.6A1.2 1.2 0 0 0 4.2 14h7.6a1.2 1.2 0 0 0 1.2-1.2V4.2A1.2 1.2 0 0 0 11.8 3h-1.3M6 8h4M6 11h2.5" />
      <rect x="5.5" y="1.5" width="5" height="3" rx="0.8" stroke={A} />
    </>
  ),
  rename: (
    <>
      <path d="M6.8 4.5H3a1.2 1.2 0 0 0-1.2 1.2v4.6A1.2 1.2 0 0 0 3 11.5h3.8M10.8 4.5H13a1.2 1.2 0 0 1 1.2 1.2v4.6A1.2 1.2 0 0 1 13 11.5h-2.2" />
      <path d="M8.8 2v12M7.3 2h3M7.3 14h3" stroke={A} />
    </>
  ),
  /** 存到待办（截图工具条主按钮）：打勾清单 + 强调色加号 */
  saveTodo: (
    <>
      <path d="M2 4.5l1.2 1.2 2.3-2.2M7.5 4.5h6M2 9.5l1.2 1.2 2.3-2.2M7.5 9.5h1.5" />
      <path d="M12.2 9.5v5M9.7 12h5" stroke={A} />
    </>
  ),
} as const;

export type IconName = keyof typeof PATHS;

export function Icon({ name, size = 16, className }: { name: IconName; size?: number; className?: string }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.6}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      className={className}
    >
      {PATHS[name]}
    </svg>
  );
}

/** 应用图标（简化版，用于标题栏与悬浮窗）。 */
export function AppLogo({ size = 26 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 256 256" aria-hidden="true">
      <rect width="256" height="256" rx="58" fill="#0B4F9C" />
      <path d="M50 34L84 64M122 20V58M194 34L160 64" stroke="#FFC247" strokeWidth="12" strokeLinecap="round" />
      <rect x="80" y="76" width="40" height="34" rx="4" fill="#FFC247" />
      <rect x="124" y="76" width="40" height="34" rx="4" fill="#FFFFFF" />
      <rect x="58" y="114" width="128" height="34" rx="4" fill="#9CC8F5" />
      <path d="M34 152H222L200 192Q196 200 186 200H72Q62 200 58 192Z" fill="#FFFFFF" />
      <path d="M28 222q16-10 32 0t32 0t32 0t32 0t32 0t32 0" fill="none" stroke="#5DA3EE" strokeWidth="10" strokeLinecap="round" />
    </svg>
  );
}

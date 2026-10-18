# xiOffice

Word / Excel 桌面编辑器（Windows / macOS / Linux），基于 Electron。直接打开、编辑、保存本机的 `.docx`、`.xlsx`，文件不上传到任何服务器。

- **Word**：打开 `.docx` / `.doc` / `.dot`，所见即所得分页排版（页眉页脚、表格、文本框、编号、脚注、分栏、公式、修订与批注……），按系统字体的度量计算行高；保存为 `.docx` 时只重写改过的部分，其余内容逐字节保留；导出 PDF、打印。
- **表格**：打开 `.xlsx` / `.xlsm` / `.csv`，公式引擎（含动态数组）、条件格式、图表、筛选排序、分组、合并单元格，保存为 `.xlsx`。

## 下载

到 [Releases](https://github.com/gopcgreen-xo/xiOffice/releases) 下载安装包：Windows 用 `.exe`，macOS 用 `.dmg`，Linux 用 `.AppImage`。安装后可以在资源管理器里直接双击 Word / Excel 文件用 xiOffice 打开（安装包未做代码签名，系统可能提示来源未知）。

## 问题反馈

如果遇到文件打不开、显示和 Office 不一致，或者保存后 Office / WPS 报错，欢迎[提交 Issue](https://github.com/gopcgreen-xo/xiOffice/issues/new/choose)，并附上能复现问题的文件（请先去掉敏感内容）。

## 开发

需要 Node.js 20 以上。

```bash
npm install
npm run dev        # 启动桌面版（Vite 开发服务器 + Electron 窗口，改代码自动刷新）
```

其它命令：

```bash
npm run dist       # 打包安装程序，输出到 release/（在哪个系统上运行就打哪个系统的包）
npm start          # 构建后以生产模式启动桌面版
npm test           # 单元测试（vitest）
npm run typecheck  # 类型检查
npm run dev:web    # 在浏览器里调试（http://localhost:1430；文件放在内存里，保存时下载）
npm run dev:word   # 只跑 Word 编辑器的开发页（端口 1431）
npm run dev:sheet  # 只跑表格编辑器的开发页（端口 1432）
```

发布新版本：改 `package.json` 的 `version`，推送 `v` 开头的标签（如 `v0.2.0`），GitHub Actions 会打出三个平台的安装包，放进一个草稿 Release。

## 目录

| 目录 | 内容 |
| --- | --- |
| `electron/` | 主进程：窗口、菜单、文件读写、系统对话框（`main.cjs`），暴露给页面的接口（`preload.cjs`），读取系统字体度量（`fonts.cjs`） |
| `src/` | 界面外壳：标签页、打开文件、对话框，以及编辑器宿主的实现（`src/host.ts`） |
| `editors/word/` | Word 编辑器（`src/core` 文档模型和排版，`src/ooxml` docx 读写，`src/doc97` 旧版 .doc 读取，`src/ui` 界面） |
| `editors/sheet/` | 表格编辑器（`src/core` 工作簿模型、公式引擎、xlsx 读写，`src/ui` 界面） |
| `editors/shared/` | 编辑器公共件：宿主接口 `EditorHost`、草稿恢复、打印、PDF 导出、图标 |
| `editors/*/samples/` | 测试和演示用的样例文件 |
| `docs/wbp/schema/` | 工作包（.wbp）里列属性、表单等的 JSON Schema |

编辑器不直接访问文件系统，只通过 `EditorHost` 接口（`editors/shared/src/host.ts`）使用宿主能力。想把编辑器嵌进自己的应用，实现这个接口、调用 `setEditorHost()`，然后渲染 `<Editor tab={{ id, title, path }} />` 即可，参考 `src/host.ts`。

## 已知限制

- 旧版 `.xls` 暂不支持（请先另存为 `.xlsx`）；旧版 `.doc` 可以打开，保存为 `.docx`。
- `csv` 一律按 UTF-8（带 BOM）写出。

## 许可证

[Apache License 2.0](LICENSE)。Word 编辑器的字节级保留保存使用了 [GenOffice](https://github.com/genspark-ai/genoffice) 的 docx-engine（Apache-2.0），第三方代码说明见 [NOTICE](NOTICE)。

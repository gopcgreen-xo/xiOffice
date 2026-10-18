# xiOffice

在浏览器里运行的 Word / Excel 编辑器：纯前端读写 `.docx`、`.xlsx`，文件不上传到任何服务器。

- **Word**：打开 `.docx` / `.doc` / `.dot`，所见即所得分页排版（页眉页脚、表格、文本框、编号、脚注、分栏、公式、修订与批注……），保存为 `.docx` 时只重写改过的部分，其余内容逐字节保留，导出 PDF、打印。
- **表格**：打开 `.xlsx` / `.xlsm` / `.csv`，公式引擎（含动态数组）、条件格式、图表、筛选排序、分组、合并单元格，保存为 `.xlsx`。

这两个编辑器提取自一个桌面工作台项目，开源出来是为了**收集真实文件上的问题**。如果你的文件打不开、显示和 Office 不一样、或者保存后 Office / WPS 报错，欢迎[提交 Issue](https://github.com/gopcgreen-xo/xiOffice/issues/new/choose)，附上能复现问题的文件（请先去掉敏感内容）。

## 运行

需要 Node.js 20 以上。

```bash
npm install
npm run dev        # 打开 http://localhost:1430
```

- 首页可以新建文档 / 表格、打开本机文件（也可以直接拖进窗口），或者打开内置示例。
- 保存（Ctrl+S）会把文件**下载到本机**；页面里的文件只在内存中，刷新后需要重新打开。
- 地址栏参数：`?new=word`、`?new=sheet` 直接新建；`?sample=报价单样例.xlsx` 打开内置示例。

其它命令：

```bash
npm test           # 单元测试（vitest）
npm run typecheck  # 类型检查
npm run build      # 生产构建，输出 dist/（纯静态，可部署到任意静态托管）
npm run dev:word   # 只跑 Word 编辑器的开发页（端口 1431）
npm run dev:sheet  # 只跑表格编辑器的开发页（端口 1432，?pack=1 / ?pack=2 演示工作包里的列属性和表单）
```

## 目录

| 目录 | 内容 |
| --- | --- |
| `src/` | 浏览器壳子：标签页、打开 / 下载文件、对话框，以及编辑器宿主的浏览器实现（`src/host.ts`） |
| `editors/word/` | Word 编辑器（`src/core` 文档模型和排版，`src/ooxml` docx 读写，`src/doc97` 旧版 .doc 读取，`src/ui` 界面） |
| `editors/sheet/` | 表格编辑器（`src/core` 工作簿模型、公式引擎、xlsx 读写，`src/ui` 界面） |
| `editors/shared/` | 编辑器公共件：宿主接口 `EditorHost`、草稿恢复、打印、PDF 导出、图标 |
| `editors/*/samples/` | 测试和演示用的样例文件 |
| `docs/wbp/schema/` | 工作包（.wbp）里列属性、表单等的 JSON Schema |

编辑器不直接访问文件系统或界面框架，只通过 `EditorHost` 接口（`editors/shared/src/host.ts`）使用宿主能力。想把编辑器嵌进自己的应用，实现这个接口、调用 `setEditorHost()`，然后渲染 `<Editor tab={{ id, title, path }} />` 即可，参考 `src/host.ts`。

## 已知限制

- 旧版 `.xls` 需要后端解析，浏览器版暂不支持（请先另存为 `.xlsx`）；旧版 `.doc` 可以打开，保存为 `.docx`。
- 浏览器里只能用系统已安装的字体排版，度量和 Word 可能略有差别。
- `csv` 一律按 UTF-8（带 BOM）写出。

## 许可证

[Apache License 2.0](LICENSE)。Word 编辑器的字节级保留保存使用了 [GenOffice](https://github.com/genspark-ai/genoffice) 的 docx-engine（Apache-2.0），第三方代码说明见 [NOTICE](NOTICE)。

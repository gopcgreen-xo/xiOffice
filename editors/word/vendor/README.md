# 第三方代码：GenOffice docx-engine

本目录的代码提取自 [GenOffice](https://github.com/genspark-ai/genoffice)（Apache License 2.0，
Copyright 2026 Mainfunc, Inc.），版本：提交 `db347087628281d49e53d009183d2fb506c3a444`。

| 目录 | 来源 | 修改 |
| --- | --- | --- |
| `docx-engine/src`、`docx-engine/tests` | `packages/docx-engine/src`、`tests` | 对 `@genoffice/zip-gate`、`@genoffice/pptx-engine/custgeom` 的引用改为本目录的相对路径；`chart.ts` 未使用的常量 `_XLSX_CONTENT_TYPE` 加 `export`（本项目开启了 noUnusedLocals）；`style-upsert.ts` 的 `StyleParaProps` 加 `lineSpacingFixed`（固定值 / 最小值行距，样式管理写公文样式用） |
| `zip-gate/src` | `packages/zip-gate/src` | 无 |
| `pptx-engine/custgeom.ts` | `packages/pptx-engine/src/custgeom.ts` | 无 |
| `pptx-engine/types.ts` | `packages/pptx-engine/src/types.ts` | 只保留 `CustomGeometry` 类型 |

许可证与版权声明见 `docx-engine/LICENSE`、`docx-engine/NOTICE`；`docx-engine/src/vendor/emf-converter` 另有其许可证（见该目录）。
GenOffice 的 `ee/` 目录是企业许可，**没有**使用其中任何代码。

用途：Word 编辑器的字节级保留保存（`saveDocx` 只重写改过的块，其余逐字节保留）、块定位（`parseDocx` 的
`docxIndex`）。本项目的封装在 `editors/word/src/ooxml/patch.ts`。升级时整体替换这几个目录，再重做上表的修改。

# Microsoft Word 编辑行为基准

这些文件由本机 Microsoft Word 16.0（Build 16.0.4266）实际生成，不是手写 OOXML 模拟文件。

`scripts/word-editing-baseline.ps1` 可重新生成。每组包含操作前 DOCX、操作后 DOCX 和操作后 PDF；`results.json` 记录操作、范围、Word 版本、结果文本、表格数量和异常。

`body-table-body-type` 使用 Selection.TypeText，其余使用 Range.Delete。单元格选区与局部文本 Range 的行为不能混为一谈。自动书签 `_GoBack` 在 Word 保存时可能移动。

关联测试：`editors/word/src/ooxml/editing-baseline.test.ts`。该测试验证结构、文字、保存回读和部分格式，不能代替 PDF 视觉及交互验收。

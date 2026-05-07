# Changelog

## 2026-05-07

### Added

- 新增 Excel 生成能力：
  - `ExcelGenerator` 作为电子表格生成统一接口。
  - `XlsWorkbookGenerator` 支持 Excel 97-2003 `.xls` 文件生成。
  - `XlsxWorkbookGenerator` 支持 Excel Open XML `.xlsx` 文件生成，并提供 SXSSF 流式写出模式。
  - `SpreadsheetGeneratorFactory` 支持按 `xls` / `xlsx` 类型创建电子表格生成器。
  - `SpreadsheetException` 用于封装电子表格生成和写出异常。
- 新增 playground Excel 演示：
  - `XlsSalesReportDemo` 生成 `.xls` 销售报表。
  - `XlsxLargeSalesReportDemo` 使用 SXSSF 生成流式 `.xlsx` 大数据量销售明细报表。
- 新增 Excel 相关测试：
  - 覆盖 `.xls` / `.xlsx` 表头样式、数据写入、公式、自动列宽和默认工作表。
  - 覆盖 playground 中 XLS / XLSX demo 生成结果。

### Changed

- 重构 playground demo 包结构：
  - `batch`：批量生成演示。
  - `excel`：Excel 生成演示。
  - `model`：演示数据模型。
  - `programmatic`：编程式固定版式 Word 生成演示。
  - `template`：模板渲染演示和演示数据。
  - `web`：Spring Web 下载演示。
- 规范化 demo 类命名：
  - `DemoMain` -> `ProgrammaticCadreDocumentDemo`
  - `TemplateDemo` -> `TemplateDocumentDemo`
  - `TemplateBatchDemo` -> `TemplateBatchDocumentDemo`
  - `TemplateDemoData` -> `CadreTemplateDemoData`
- 为 playground demo 子包和关键生成逻辑补充说明注释。
- 更新 README：
  - 增加支持文档类型和待办文档类型说明。
  - 增加 Excel `.xls` / `.xlsx` API 示例。
  - 增加 playground Excel demo 运行命令和测试命令。

### Verified

- `mvn test`
- `mvn -pl playground -am test`
- `mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.excel.XlsSalesReportDemo"`
- `mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.excel.XlsxLargeSalesReportDemo" "-Dexec.args=1000"`

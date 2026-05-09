# Changelog

## 2026-05-09

### Added

- 新增 Publisher 文档读取能力：
  - `PublisherDocumentReader` 支持 `.pub` 元数据和正文文本提取。
  - `PublisherDocumentInfo` 用于承载 Publisher 标题、主题、作者、关键词、备注和文本结果。
  - `PublisherDocumentReaderFactory` 支持按 `pub` 类型创建 Publisher 读取器。
  - `PublisherDocumentException` 用于封装 Publisher 读取异常。
- 新增 playground Publisher 演示：
  - `PublisherBrochureExtractDemo` 生成最小 `.pub` 宣传册示例并输出读取摘要。
- 新增 Publisher 相关测试：
  - 覆盖 PUB 标题、主题、作者、关键词、备注和文本内容提取。
  - 覆盖 Publisher reader factory 和 playground demo 生成结果。
- 新增 Visio 绘图读取能力：
  - `VisioDrawingReader` 支持 `.vsdx` 页面、形状、连接数和文本摘要信息提取。
  - `VisioDrawingReader` 支持 `.vsd` 文本信息提取入口。
  - `VisioDrawingInfo`、`VisioPageInfo` 和 `VisioShapeInfo` 用于承载 Visio 解析结果。
  - `VisioDrawingReaderFactory` 支持按 `vsd` / `vsdx` 类型创建 Visio 读取器。
  - `VisioDrawingException` 用于封装 Visio 读取异常。
- 新增 playground Visio 演示：
  - `VisioWorkflowExtractDemo` 生成最小 `.vsdx` 工作流示例并输出读取摘要。
- 新增 Visio 相关测试：
  - 覆盖 VSDX 页面、图形、文本、连接数和尺寸提取。
  - 覆盖 Visio reader factory 和 playground demo 生成结果。

### Changed

- 更新 README：
  - 将 Publisher `.pub` 从待办文档类型移入已支持文档类型。
  - 增加 `PublisherDocumentReader` API 示例。
  - 增加 playground Publisher demo 运行命令和测试命令。
  - 将 Visio `.vsd` / `.vsdx` 从待办文档类型移入已支持文档类型。
  - 增加 `VisioDrawingReader` API 示例。
  - 增加 playground Visio demo 运行命令和测试命令。

### Verified

- `mvn -pl word-generator test`
- `mvn -pl playground -am test`
- `mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.publisher.PublisherBrochureExtractDemo"`

## 2026-05-07

### Added

- 新增 PDF 文档生成能力：
  - 引入 Apache PDFBox 作为 PDF 生成依赖。
  - `PdfDocumentGenerator` 作为 PDF 文档生成统一接口。
  - `PdfBoxDocumentGenerator` 支持 `.pdf` 文件生成，覆盖标题、段落、表格和图片。
  - `PdfDocumentGeneratorFactory` 支持按 `pdf` 类型创建 PDF 生成器。
  - `PdfDocumentException` 用于封装 PDF 创建、绘制和写出异常。
- 新增 playground PDF 演示：
  - `PdfProjectReportDemo` 生成包含标题、段落、表格和图片的 `.pdf` 项目汇报。
- 新增 PDF 相关测试：
  - 覆盖 PDF 页数、文本内容、表格文本和图片生成结果。
  - 覆盖 playground 中 PDF demo 生成结果。

### Changed

- 更新 README：
  - 将 PDF `.pdf` 从待办文档类型移入已支持文档类型。
  - 增加 `PdfBoxDocumentGenerator` API 示例。
  - 增加 playground PDF demo 运行命令和测试命令。

### Verified

- `mvn -pl word-generator test`
- `mvn -pl playground -am test`
- `mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.pdf.PdfProjectReportDemo"`

### Added

- 新增 Outlook MSG 邮件读取能力：
  - `OutlookMessageReader` 支持读取 `.msg` 文件并提取主题、发件人、收件人、正文和附件摘要。
  - `OutlookMessageInfo` 和 `OutlookAttachmentInfo` 用于承载邮件解析结果。
  - `OutlookMessageReaderFactory` 支持按 `msg` 类型创建邮件读取器。
  - `OutlookMessageException` 用于封装 MSG 读取和解析异常。
- 新增 playground Outlook 邮件读取演示：
  - `OutlookMessageExtractDemo` 生成最小 MSG 示例文件并输出解析摘要。
- 新增 Outlook 相关测试：
  - 覆盖 MSG 主题、发件人、收件人、正文和附件摘要提取结果。
  - 覆盖 playground 中 Outlook demo 生成结果。

### Changed

- 更新 README：
  - 将 Outlook `.msg` 从待办文档类型移入已支持文档类型。
  - 增加 `OutlookMessageReader` API 示例。
  - 增加 playground Outlook demo 运行命令和测试命令。

### Verified

- `mvn -pl word-generator test`
- `mvn -pl playground -am test`
- `mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.outlook.OutlookMessageExtractDemo"`

### Added

- 新增 PowerPoint Open XML 演示文稿生成能力：
  - `XslfPresentationGenerator` 支持 `.pptx` 文件生成。
  - `PresentationGeneratorFactory` 支持按 `pptx` 类型创建演示文稿生成器。
- 新增 playground PowerPoint Open XML 演示：
  - `PptxProjectReportDemo` 生成包含封面、文本要点、表格和图片的 `.pptx` 项目汇报。
- 新增 PPTX 相关测试：
  - 覆盖 `.pptx` 标题页、文本页、表格页、图片页生成结果。
  - 覆盖 playground 中 PPTX demo 生成结果。

### Changed

- 更新 README：
  - 将 PowerPoint Open XML `.pptx` 从待办文档类型移入已支持文档类型。
  - 增加 `XslfPresentationGenerator` API 示例。
  - 增加 playground PPTX demo 运行命令和测试命令。

### Verified

- `mvn -pl word-generator test`
- `mvn -pl playground -am test`
- `mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.presentation.PptxProjectReportDemo"`

### Added

- 新增 PowerPoint 97-2003 演示文稿生成能力：
  - `PresentationGenerator` 作为演示文稿生成统一接口。
  - `HslfPresentationGenerator` 支持 `.ppt` 文件生成。
  - `PresentationGeneratorFactory` 支持按 `ppt` 类型创建演示文稿生成器。
  - `PresentationException` 用于封装演示文稿生成和写出异常。
- 新增 playground PowerPoint 演示：
  - `PptProjectReportDemo` 生成包含封面、文本要点、表格和图片的 `.ppt` 项目汇报。
- 新增 PowerPoint 相关测试：
  - 覆盖 `.ppt` 标题页、文本页、表格页、图片页生成结果。
  - 覆盖 playground 中 PPT demo 生成结果。

### Changed

- 更新 README：
  - 将 PowerPoint 97-2003 `.ppt` 从待办文档类型移入已支持文档类型。
  - 增加 `HslfPresentationGenerator` API 示例。
  - 增加 playground PPT demo 运行命令和测试命令。

### Verified

- `mvn -pl word-generator test`
- `mvn -pl playground -am test`
- `mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.presentation.PptProjectReportDemo"`

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

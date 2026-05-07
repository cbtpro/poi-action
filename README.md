# poi-action

基于 Apache POI 的文档生成库，支持 Word、Excel 和 PowerPoint 的编程式生成。提供流式 API、模板引擎、工厂模式等设计模式，让 Office 文档生成变得简单高效。

## 功能特性

- **编程式生成**：使用流式 API（Builder 模式）直接构建 Word 文档
- **模板式生成**：基于 Word 模板进行数据填充和循环渲染
- **多格式支持**：支持 Word、Excel 和 PowerPoint 常见格式
- **丰富的元素支持**：段落、标题、表格、图片等
- **表格高级功能**：单元格合并、循环填充等
- **图片灵活处理**：支持本地文件、URL、Base64 等多种图片源
- **固定版式工具**：提供 DOCX 页面设置、固定宽度表格、跨列、纵向合并、表格内图片等公共能力
- **批量生成**：模板读取一次后复用，支持多线程写入和进度统计
- **设计模式应用**：工厂模式、建造者模式、模板方法模式

## 支持的文档类型

当前项目围绕 Word 文档生成能力封装，`WordGeneratorFactory` 已支持以下类型：

| 类型 | 文件扩展名 | POI 组件 | 当前能力 |
|------|------------|----------|----------|
| Word 97-2003 文档 | `.doc` | HWPF | 编程式生成，支持段落、标题、文本表格等基础能力，表格和图片能力有限 |
| Word Open XML 文档 | `.docx` | XWPF | 编程式生成和模板式生成，支持段落、标题、表格、图片、占位符渲染、列表循环、固定版式工具 |
| Excel 97-2003 工作簿 | `.xls` | HSSF | 表格数据写入、表头样式、公式、自动列宽 |
| Excel Open XML 工作簿 | `.xlsx` | XSSF / SXSSF | 表格数据写入、表头样式、公式、自动列宽、流式大数据量写出 |
| PowerPoint 97-2003 演示文稿 | `.ppt` | HSLF | 标题页、文本页、表格页、图片页 |
| PowerPoint Open XML 演示文稿 | `.pptx` | XSLF | 标题页、文本页、表格页、图片页 |

## 待办文档类型

Apache POI 还覆盖多种 Office/OLE2/OOXML 文档格式，当前项目尚未封装这些能力，后续可按优先级扩展：

| 类型 | 文件扩展名 | POI 组件 | 计划能力 |
|------|------------|----------|----------|
| Visio 绘图 | `.vsd` / `.vsdx` | HDGF / XDGF | 图形结构读取和基础信息提取 |
| Outlook 邮件 | `.msg` | HSMF | 邮件主题、正文、附件信息提取 |
| Publisher 文档 | `.pub` | HPBF | 元数据和文本内容提取 |
| PDF 文档 | `.pdf` | 非 POI 原生能力 | 如需支持，建议接入 PDFBox、OpenPDF 或 LibreOffice 转换链路 |

## 项目结构

```
poi-action/
├── word-generator/          # 核心库模块
│   └── src/main/java/com/chenbitao/word/
│       ├── core/            # 核心接口定义
│       ├── builder/         # 建造者实现
│       ├── factory/         # 工厂类
│       ├── doc/             # DOC 格式生成器
│       ├── docx/            # DOCX 格式生成器、模板生成器
│       │   └── util/        # DOCX 页面、固定表格、图片来源等公共工具
│       ├── excel/           # XLS / XLSX 工作簿生成器
│       ├── presentation/    # PPT 演示文稿生成器
│       ├── constant/        # 常量定义
│       └── exception/       # 异常类
├── playground/              # 演示和测试模块
│   └── src/main/java/com/chenbitao/word/playground/
│       └── demo/            # 演示示例
│           ├── batch/       # 批量生成演示
│           ├── excel/       # Excel 生成演示
│           ├── model/       # 演示数据模型
│           ├── presentation/# PowerPoint 生成演示
│           ├── programmatic/# 编程式固定版式生成演示
│           ├── template/    # 模板渲染演示和演示数据
│           └── web/         # Spring Web 下载演示
└── pom.xml                  # Maven 项目配置
```

## 安装与快速开始

### 前置要求

- Java 8+
- Maven 3.6+

### 编译和安装

```shell
# 编译整个项目
mvn clean install

# 运行演示（编程式生成）
mvn -pl playground exec:java -Dexec.mainClass=com.chenbitao.word.playground.demo.programmatic.ProgrammaticCadreDocumentDemo

# 运行演示（模板式生成）
mvn -pl playground exec:java -Dexec.mainClass=com.chenbitao.word.playground.demo.template.TemplateDocumentDemo

# 运行批量生成演示，输出到 playground/target/out
mvn -pl playground exec:java \
  -Dexec.mainClass=com.chenbitao.word.playground.demo.batch.TemplateBatchDocumentDemo \
  -Dexec.args="1000 8"

# 运行 Excel 97-2003 生成演示，输出到 playground/target/excel-sales-demo.xls
mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.excel.XlsSalesReportDemo"

# 运行 Excel Open XML 流式生成演示，输出到 playground/target/excel-large-sales-demo.xlsx
mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.excel.XlsxLargeSalesReportDemo" "-Dexec.args=1000"

# 运行 PowerPoint 97-2003 生成演示，输出到 playground/target/project-report-demo.ppt
mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.presentation.PptProjectReportDemo"

# 运行 PowerPoint Open XML 生成演示，输出到 playground/target/project-report-demo.pptx
mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.presentation.PptxProjectReportDemo"

# 启动 Spring Boot 演示服务，actuator 仅在 dev/test 生效
mvn -pl playground spring-boot:run -Dspring-boot.run.profiles=dev
```

## API 文档

### 核心接口 - WordGenerator

**WordGenerator** 是定义 Word 文档生成器的标准接口，所有具体的生成器实现都需要实现此接口。

```java
public interface WordGenerator {
    void createDocument();
    void addParagraph(String text);
    void addTitle(String text, int level);
    void addTable(int rows, int cols);
    void addTable(List<List<String>> rows);
    void addTable(WordTable table);
    void addImage(InputStream inputStream, int width, int height);
    void save(String path);
}
```

#### 方法说明

| 方法名 | 参数说明 | 返回值 | 说明 |
|--------|--------|-------|------|
| `createDocument()` | 无 | void | 创建一个新的 Word 文档实例 |
| `addParagraph(String text)` | text: 段落文本内容 | void | 添加普通文本段落 |
| `addTitle(String text, int level)` | text: 标题文本<br/>level: 标题级别(1-6) | void | 添加标题段落（支持 H1-H6）|
| `addTable(int rows, int cols)` | rows: 行数<br/>cols: 列数 | void | 添加空表格 |
| `addTable(List<List<String>> rows)` | rows: 表格数据 | void | 添加带内容的表格 |
| `addTable(WordTable table)` | table: 表格模型对象 | void | 添加结构化表格 |
| `addImage(InputStream input, int w, int h)` | inputStream: 图片流<br/>width: 宽度(EMU)<br/>height: 高度(EMU) | void | 添加图片 |
| `save(String path)` | path: 保存路径 | void | 保存文档到指定路径 |

---

### 建造者类 - WordBuilder

**WordBuilder** 使用建造者模式提供流式 API，允许链式调用来构建 Word 文档。

```java
public class WordBuilder {
    public WordBuilder(WordGenerator generator);
    public WordBuilder title(String text);
    public WordBuilder paragraph(String text);
    public WordBuilder paragraphList(Iterable<String> list);
    public WordBuilder table(int r, int c);
    public WordBuilder table(List<List<String>> rows);
    public WordBuilder table(WordTable table);
    public WordBuilder image(InputStream inputStream, int width, int height);
    public void build(String path);
}
```

#### 方法说明

| 方法名 | 参数说明 | 返回值 | 说明 |
|--------|--------|-------|------|
| `WordBuilder(WordGenerator generator)` | generator: 生成器实例 | WordBuilder | 构造建造者（自动调用 createDocument）|
| `title(String text)` | text: 标题文本 | WordBuilder | 添加一级标题（H1），支持链式调用 |
| `paragraph(String text)` | text: 段落文本 | WordBuilder | 添加段落，支持链式调用 |
| `paragraphList(Iterable<String> list)` | list: 字符串可迭代对象 | WordBuilder | 添加项目符号列表，支持链式调用 |
| `table(int r, int c)` | r: 行数<br/>c: 列数 | WordBuilder | 添加空表格，支持链式调用 |
| `table(List<List<String>> rows)` | rows: 表格数据 | WordBuilder | 添加带内容的表格，支持链式调用 |
| `table(WordTable table)` | table: 表格模型 | WordBuilder | 添加结构化表格，支持链式调用 |
| `image(InputStream input, int w, int h)` | inputStream: 图片流<br/>width: 宽度(EMU)<br/>height: 高度(EMU) | WordBuilder | 添加图片，支持链式调用 |
| `build(String path)` | path: 保存路径 | void | 构建并保存文档 |

#### 使用示例

```java
import com.chenbitao.word.builder.WordBuilder;
import com.chenbitao.word.factory.WordGeneratorFactory;
import java.util.Arrays;

// 创建编程式 Word 文档
WordBuilder builder = new WordBuilder(WordGeneratorFactory.get("docx"));
builder
    .title("项目报告")
    .paragraph("这是项目的总体概述。")
    .title("工作成果", 2)
    .paragraphList(Arrays.asList(
        "完成需求分析",
        "实现核心功能",
        "通过单元测试"
    ))
    .table(3, 2)
    .paragraph("合作伙伴：ABC 公司")
    .build("output.docx");
```

---

### 工厂类 - WordGeneratorFactory

**WordGeneratorFactory** 使用工厂模式和缓存机制提供不同类型的 Word 文档生成器实例。

```java
public class WordGeneratorFactory {
    public static WordGenerator get(String type);
}
```

#### 方法说明

| 方法名 | 参数说明 | 返回值 | 说明 |
|--------|--------|-------|------|
| `get(String type)` | type: 文档类型 "doc" 或 "docx" | WordGenerator | 返回对应类型的生成器（缓存的单例）|

#### 支持的文档类型

| 类型 | 说明 | 生成器类 |
|------|------|--------|
| `"doc"` | Microsoft Word 97-2003 格式 | DocWordGenerator |
| `"docx"` | Office Open XML 格式 | DocxWordGenerator |

#### 使用示例

```java
// 获取 DOCX 生成器
WordGenerator docxGenerator = WordGeneratorFactory.get("docx");

// 获取 DOC 生成器
WordGenerator docGenerator = WordGeneratorFactory.get("doc");

// 两次调用返回同一个缓存实例
assert WordGeneratorFactory.get("docx") == WordGeneratorFactory.get("docx");
```

---

### 模板生成器 - TemplateWordGenerator

**TemplateWordGenerator** 支持基于 Word 模板进行数据填充和循环渲染。使用 `${key}` 格式的占位符进行文本替换，使用 `${list.field}` 格式进行列表循环填充。

```java
public class TemplateWordGenerator {
    public TemplateWordGenerator(InputStream templateStream);
    public void render(Map<String, Object> data);
    public void save(String path);
    public static Picture picture(Object source);
    public static Picture picture(Object source, int widthEmu, int heightEmu);
}
```

#### 方法说明

| 方法名 | 参数说明 | 返回值 | 说明 |
|--------|--------|-------|------|
| `TemplateWordGenerator(InputStream template)` | template: 模板文件流 | TemplateWordGenerator | 创建模板生成器实例 |
| `render(Map<String, Object> data)` | data: 填充数据 | void | 使用数据渲染模板 |
| `save(String path)` | path: 保存路径 | void | 保存渲染后的文档 |
| `picture(Object source)` | source: 图片源 | Picture | 创建图片对象（默认尺寸）|
| `picture(Object source, int w, int h)` | source: 图片源<br/>widthEmu: 宽度<br/>heightEmu: 高度 | Picture | 创建指定尺寸的图片对象 |

#### 图片源支持

图片对象的 `source` 参数支持以下类型：
- `InputStream`：输入流
- `byte[]`：字节数组
- `File`：本地文件
- `Path`：文件路径
- `URL`：网络 URL
- `URI`：URI 地址
- `String`：Base64 编码字符串

#### 占位符格式

| 格式 | 说明 | 示例 |
|------|------|------|
| `${key}` | 简单文本替换 | `${name}` → "张三" |
| `${list.field}` | 列表循环填充 | `${education.degree}` |
| `${picture}` | 图片替换 | `${photo}` |

如果图片值为空或未传入，图片占位符会被清空。图片尺寸属于业务数据，推荐在构造 `Picture` 时显式传入：

```java
data.put("photo", TemplateWordGenerator.picture(
    source,
    Units.toEMU(80),
    Units.toEMU(100)
));
```

#### 使用示例

```java
import com.chenbitao.word.docx.TemplateWordGenerator;
import java.io.InputStream;
import java.util.*;

// 加载模板
InputStream templateStream = TemplateDocumentDemo.class
    .getResourceAsStream("/template.docx");

// 创建模板生成器
TemplateWordGenerator generator = new TemplateWordGenerator(templateStream);

// 准备填充数据
Map<String, Object> data = new HashMap<>();
data.put("name", "张三");
data.put("sex", "男");
data.put("age", "28");
data.put("position", "技术总监");

// 列表数据类型
List<Map<String, String>> education = new ArrayList<>();
Map<String, String> edu = new HashMap<>();
edu.put("type", "全日制教育");
edu.put("degree", "大学本科");
edu.put("university", "清华大学");
edu.put("major", "计算机科学");
education.add(edu);
data.put("education", education);

// 图片数据
data.put("photo", TemplateWordGenerator.picture(
    new File("photo.jpg"),
    Units.toEMU(80),   // 80pt 宽度
    Units.toEMU(100)   // 100pt 高度
));

// 渲染并保存
generator.render(data);
generator.save("output.docx");
```

---

### DOCX 公共工具 - docx.util

`word-generator` 提供了一组不绑定业务场景的 DOCX 工具，适合编程式生成复杂固定版式时复用。

| 类名 | 说明 | 典型用途 |
|------|------|----------|
| `DocxPageUtils` | 页面和字体工具 | 设置纸张、页边距、居中标题、中文字体 |
| `DocxFixedTable` | 固定宽度表格渲染器 | 指定列宽、行高、跨列、纵向合并、表格内图片 |
| `ImageSourceUtils` | 图片来源读取和格式转换 | 读取文件、路径、URL、URI、Base64、字节数组、输入流，并转换为 PNG |

#### 固定表格示例

```java
import com.chenbitao.word.docx.util.DocxFixedTable;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.Arrays;

import static com.chenbitao.word.docx.util.DocxFixedTable.empty;
import static com.chenbitao.word.docx.util.DocxFixedTable.row;
import static com.chenbitao.word.docx.util.DocxFixedTable.text;

XWPFDocument document = new XWPFDocument();
DocxFixedTable.Options options = new DocxFixedTable.Options(new int[]{1200, 2400, 2400})
    .tableWidthDxa(6000)
    .cellMargins(0, 108, 0, 108)
    .font("宋体", 12);

DocxFixedTable.render(document, options, Arrays.asList(
    row(text("姓名"), text("张三"), text("照片").vRestart()),
    row(text("说明"), text("跨两列内容").span(2)),
    row(text("出生年月"), text("2008-01"), empty().vContinue())
));
```

---

### Excel 97-2003 生成器 - XlsWorkbookGenerator

`XlsWorkbookGenerator` 基于 POI HSSF 生成 `.xls` 文件，适合轻量表格导出场景。

```java
import com.chenbitao.word.excel.ExcelGenerator;
import com.chenbitao.word.factory.SpreadsheetGeneratorFactory;
import java.util.Arrays;

ExcelGenerator generator = SpreadsheetGeneratorFactory.get("xls");
generator.createWorkbook();
generator.addHeaderRow("销售数据", Arrays.asList("产品", "数量", "单价", "小计"));
generator.addRows("销售数据", Arrays.asList(
    Arrays.asList("A", 2, 5.5),
    Arrays.asList("B", 3, 4.0)
));
generator.setFormula("销售数据", 1, 3, "B2*C2");
generator.setFormula("销售数据", 2, 3, "B3*C3");
generator.autoSizeColumns("销售数据");
generator.save("sales.xls");
```

playground 中也提供了可运行的 XLS 演示：

```shell
mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.excel.XlsSalesReportDemo"
```

输出文件：

```text
playground/target/excel-sales-demo.xls
```

---

### Excel Open XML 生成器 - XlsxWorkbookGenerator

`XlsxWorkbookGenerator` 默认基于 POI XSSF 生成 `.xlsx` 文件；大数据量导出可使用 `streaming(...)` 创建 SXSSF 流式生成器。

```java
import com.chenbitao.word.excel.ExcelGenerator;
import com.chenbitao.word.excel.XlsxWorkbookGenerator;
import com.chenbitao.word.factory.SpreadsheetGeneratorFactory;
import java.util.Arrays;

ExcelGenerator generator = SpreadsheetGeneratorFactory.get("xlsx");
generator.createWorkbook();
generator.addHeaderRow("销售数据", Arrays.asList("产品", "数量", "单价", "小计"));
generator.addRows("销售数据", Arrays.asList(
    Arrays.asList("A", 2, 5.5),
    Arrays.asList("B", 3, 4.0)
));
generator.setFormula("销售数据", 1, 3, "B2*C2");
generator.autoSizeColumns("销售数据");
generator.save("sales.xlsx");

ExcelGenerator streamingGenerator = XlsxWorkbookGenerator.streaming(100);
```

playground 中也提供了可运行的 XLSX 流式导出演示：

```shell
mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.excel.XlsxLargeSalesReportDemo" "-Dexec.args=1000"
```

输出文件：

```text
playground/target/excel-large-sales-demo.xlsx
```

---

### PowerPoint 97-2003 生成器 - HslfPresentationGenerator

`HslfPresentationGenerator` 基于 POI HSLF 生成 `.ppt` 文件，支持标题页、文本页、表格页和图片页。

```java
import com.chenbitao.word.factory.PresentationGeneratorFactory;
import com.chenbitao.word.presentation.PresentationGenerator;
import java.util.Arrays;

PresentationGenerator generator = PresentationGeneratorFactory.get("ppt");
generator.createPresentation();
generator.addTitleSlide("项目汇报", "PowerPoint 97-2003 生成演示");
generator.addTextSlide("新增能力", Arrays.asList(
    "支持标题页",
    "支持文本页",
    "支持表格页和图片页"
));
generator.addTableSlide("格式支持", Arrays.asList(
    Arrays.asList("类型", "格式", "状态"),
    Arrays.asList("PowerPoint", ".ppt", "已支持")
));
generator.save("project-report.ppt");
```

playground 中也提供了可运行的 PPT 演示：

```shell
mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.presentation.PptProjectReportDemo"
```

输出文件：

```text
playground/target/project-report-demo.ppt
```

---

### PowerPoint Open XML 生成器 - XslfPresentationGenerator

`XslfPresentationGenerator` 基于 POI XSLF 生成 `.pptx` 文件，支持标题页、文本页、表格页和图片页。

```java
import com.chenbitao.word.factory.PresentationGeneratorFactory;
import com.chenbitao.word.presentation.PresentationGenerator;
import java.util.Arrays;

PresentationGenerator generator = PresentationGeneratorFactory.get("pptx");
generator.createPresentation();
generator.addTitleSlide("项目汇报", "PowerPoint Open XML 生成演示");
generator.addTextSlide("新增能力", Arrays.asList(
    "支持标题页",
    "支持文本页",
    "支持表格页和图片页"
));
generator.addTableSlide("格式支持", Arrays.asList(
    Arrays.asList("类型", "格式", "状态"),
    Arrays.asList("PowerPoint", ".pptx", "已支持")
));
generator.save("project-report.pptx");
```

playground 中也提供了可运行的 PPTX 演示：

```shell
mvn -pl playground exec:java "-Dexec.mainClass=com.chenbitao.word.playground.demo.presentation.PptxProjectReportDemo"
```

输出文件：

```text
playground/target/project-report-demo.pptx
```

---

### 常量类 - BlankConstants

**BlankConstants** 定义了常用的空白/占位符常量。

```java
public class BlankConstants {
    public static final String EMPTY = "";      // 空字符串
    public static final String DASH = "－";      // 中文破折号
    public static final String NONE = "无";      // 无
}
```

---

### 异常类 - WordException

**WordException** 是 Word 生成过程中的自定义异常。

```java
public class WordException extends RuntimeException {
    public WordException(String message);
    public WordException(String message, Throwable cause);
}
```

---

## 使用示例

### 示例 1：编程式生成简单文档

```java
import com.chenbitao.word.builder.WordBuilder;
import com.chenbitao.word.factory.WordGeneratorFactory;
import java.util.Arrays;

public class SimpleExample {
    public static void main(String[] args) {
        WordBuilder builder = new WordBuilder(WordGeneratorFactory.get("docx"));
        builder
            .title("年度总结报告")
            .paragraph("本报告总结了过去一年的工作成果和经验教训。")
            .title("主要成就", 2)
            .paragraphList(Arrays.asList(
                "提升了系统性能 30%",
                "完成了技术架构升级",
                "培养了 5 名技术人才"
            ))
            .title("下年目标", 2)
            .paragraph("继续深化技术创新，提升产品竞争力。")
            .build("annual_report.docx");
        
        System.out.println("文档生成成功：annual_report.docx");
    }
}
```

### 示例 2：使用模板生成复杂文档

```java
import com.chenbitao.word.docx.TemplateWordGenerator;
import org.apache.poi.util.Units;
import java.io.InputStream;
import java.util.*;

public class TemplateExample {
    public static void main(String[] args) throws Exception {
        // 加载模板
        InputStream template = TemplateExample.class
            .getResourceAsStream("/cadre_template.docx");
        
        // 创建生成器
        TemplateWordGenerator generator = new TemplateWordGenerator(template);
        
        // 准备基本信息
        Map<String, Object> data = new HashMap<>();
        data.put("name", "李四");
        data.put("sex", "女");
        data.put("birthday", "1990-05");
        data.put("age", "33");
        data.put("position", "部门经理");
        
        // 准备教育背景（列表）
        List<Map<String, String>> educationList = new ArrayList<>();
        
        Map<String, String> edu1 = new HashMap<>();
        edu1.put("type", "全日制教育");
        edu1.put("degree", "硕士研究生");
        edu1.put("university", "复旦大学");
        edu1.put("major", "工商管理");
        educationList.add(edu1);
        
        Map<String, String> edu2 = new HashMap<>();
        edu2.put("type", "在职教育");
        edu2.put("degree", "博士");
        edu2.put("university", "北京大学");
        edu2.put("major", "经济学");
        educationList.add(edu2);
        
        data.put("education", educationList);
        
        // 准备工作经历
        List<String> workExperience = Arrays.asList(
            "2015.06 - 2017.12  ABC 公司  技术主管",
            "2018.01 - 2020.06  XYZ 公司  部门经理",
            "2020.07 - 现在      DEF 公司  总经理"
        );
        data.put("workExperience", workExperience);
        
        // 准备照片
        data.put("photo", TemplateWordGenerator.picture(
            new File("profile.jpg"),
            Units.toEMU(80),
            Units.toEMU(100)
        ));
        
        // 年度考核结果
        List<String> assessmentResults = Arrays.asList(
            "2023年：优秀",
            "2022年：优秀",
            "2021年：合格"
        );
        data.put("annualAssessment", assessmentResults);
        
        // 渲染并保存
        generator.render(data);
        generator.save("cadre_dossier.docx");
        
        System.out.println("文档生成成功：cadre_dossier.docx");
    }
}
```

### 示例 3：编程方式创建表格

```java
import com.chenbitao.word.builder.WordBuilder;
import com.chenbitao.word.factory.WordGeneratorFactory;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class TableExample {
    public static void main(String[] args) {
        WordBuilder builder = new WordBuilder(WordGeneratorFactory.get("docx"));
        builder
            .title("员工信息表")
            // 创建带数据的表格
            .table(Arrays.asList(
                Arrays.asList("姓名", "部门", "职位", "薪资"),
                Arrays.asList("张三", "技术部", "开发工程师", "15K-20K"),
                Arrays.asList("李四", "产品部", "产品经理", "18K-25K"),
                Arrays.asList("王五", "销售部", "销售总监", "20K-30K")
            ))
            .build("employee_table.docx");
        
        System.out.println("表格文档生成成功");
    }
}
```

### 示例 4：生成与模板同数据的编程式文档

`playground` 中的 `ProgrammaticCadreDocumentDemo` 使用 `CadreTemplateDemoData.create()` 准备和模板演示相同的数据，再由 `ProgrammaticCadreDocumentWriter` 直接编程式生成 Word。业务字段组织、干部表版式仍留在 playground demo 中，页面、固定表格、图片来源等通用能力由 `word-generator` 的 `docx.util` 提供。

```shell
mvn -pl playground exec:java -Dexec.mainClass=com.chenbitao.word.playground.demo.programmatic.ProgrammaticCadreDocumentDemo
```

输出文件：

```text
playground/target/programmatic-demo.docx
```

---

## 架构设计

### 设计模式应用

1. **工厂模式（Factory Pattern）**
   - `WordGeneratorFactory`：统一管理不同类型的生成器实例
   - 利用缓存机制提高性能

2. **建造者模式（Builder Pattern）**
   - `WordBuilder`：提供流式 API 进行文档构建
   - 支持链式调用，提升代码可读性

3. **模板方法模式（Template Method Pattern）**
   - `AbstractWordGenerator`：定义文档生成的骨架
   - 由子类实现具体的格式操作

4. **策略模式（Strategy Pattern）**
   - 不同的生成器实现不同的策略
   - 支持 DOC 和 DOCX 两种格式

---

## 主要实现类

| 类名 | 说明 | 备注 |
|------|------|------|
| `DocxWordGenerator` | DOCX 格式生成器 | 推荐使用，现代格式 |
| `DocWordGenerator` | DOC 格式生成器 | 兼容旧版本 Office |
| `TemplateWordGenerator` | 模板渲染生成器 | 支持占位符和循环 |
| `HslfPresentationGenerator` | PPT 格式生成器 | 标题页、文本页、表格页、图片页 |
| `XslfPresentationGenerator` | PPTX 格式生成器 | 标题页、文本页、表格页、图片页 |
| `WordBuilder` | 建造者类 | 流式 API |
| `WordGeneratorFactory` | 工厂类 | 获取生成器实例 |
| `SpreadsheetGeneratorFactory` | 电子表格工厂类 | 获取 XLS / XLSX 生成器实例 |
| `PresentationGeneratorFactory` | 演示文稿工厂类 | 获取 PPT 生成器实例 |
| `DocxPageUtils` | DOCX 页面工具 | 页面尺寸、边距、标题、字体 |
| `DocxFixedTable` | DOCX 固定表格工具 | 固定列宽、跨列、纵向合并、表格图片 |
| `ImageSourceUtils` | 图片来源工具 | 文件、URL、Base64、字节流读取和 PNG 转换 |

---

## 性能

poi-action 在性能方面表现出色，采用了多项优化策略：

### 性能指标

- **测试环境**：Apple M2 芯片，24GB 内存，macOS 系统
- **批量生成**：支持高效的批量文档生成
  - 生成 **320,000 个 Word 文档**仅需 **20-32 秒**
  - 平均每个文档生成时间：**0.099 毫秒**
  - 吞吐量：**10,278 文档/秒**
- **多线程支持**：支持 8 线程并发生成，充分利用多核 CPU
- **内存效率**：低内存占用，适合大规模批量处理

### 性能优化策略

1. **工厂缓存机制**
   - 生成器实例复用，避免重复创建
   - 单例缓存设计，内存占用少

2. **流式 API 设计**
   - 链式调用高效，减少中间对象创建
   - 建造者模式优化代码执行路径

3. **POI 高效使用**
   - 合理使用 Apache POI 的底层 API
   - 避免不必要的格式转换和复制

4. **内存优化**
   - 及时释放文档资源
   - 支持流式处理大批量数据

5. **多线程并发**
   - 利用现代多核 CPU 架构
   - 线程池管理，控制并发数量

### 性能测试示例

```shell
# 生成 1000 个文档，8 线程写入
mvn -pl playground exec:java \
  -Dexec.mainClass=com.chenbitao.word.playground.demo.batch.TemplateBatchDocumentDemo \
  -Dexec.args="1000 8"

# 通过参数指定数量和线程数，例如生成 320000 个，8 线程写入
mvn -pl playground exec:java \
  -Dexec.mainClass=com.chenbitao.word.playground.demo.batch.TemplateBatchDocumentDemo \
  -Dexec.args="320000 8"
```

`TemplateBatchDocumentDemo` 的优化点：

- 模板文件只读取一次，后续渲染复用模板字节。
- 示例数据只创建一次，避免批量任务把时间消耗在无关对象构造上。
- 渲染后的文档字节一次生成，多线程负责写出不同文件。
- 进度日志包含完成数、失败数、百分比、速度和预计剩余时间。
- 输出目录固定为 `playground/target/out`。

---

## 测试

项目包含完整的单元测试覆盖：

```shell
# 运行所有测试
mvn test

# 运行指定测试类
mvn -pl playground test -Dtest=WordBuilderTest
mvn -pl playground test -Dtest=TemplateWordGeneratorTest
mvn -pl playground test -Dtest=WordGeneratorFactoryTest
mvn -pl playground test -Dtest=ProgrammaticCadreDocumentWriterTest
mvn -pl playground -am test -Dtest=XlsSalesReportDemoTest
mvn -pl playground -am test -Dtest=XlsxLargeSalesReportDemoTest
mvn -pl playground -am test -Dtest=PptProjectReportDemoTest
mvn -pl playground -am test -Dtest=PptxProjectReportDemoTest

# 运行 word-generator 的 DOCX 公共工具测试
mvn -pl word-generator test -Dtest=DocxPageUtilsTest,DocxFixedTableTest,ImageSourceUtilsTest
```

### Actuator

```http
GET /actuator
GET /actuator/health
GET /actuator/info
GET /actuator/metrics
GET /actuator/env
GET /actuator/beans
```

Actuator 只在 `dev` 和 `test` profile 中生效。未启用这两个 profile 时，默认配置会关闭 actuator endpoint 暴露：

```yaml
management:
  endpoints:
    enabled-by-default: false
    web:
      exposure:
        exclude: "*"
```

`dev` / `test` 环境会暴露：

```yaml
management:
  endpoints:
    enabled-by-default: true
    web:
      exposure:
        include: health,info,metrics,env,beans
```

如果 `/actuator` 访问不到，优先检查：

- `playground` 是否包含 `spring-boot-starter-actuator` 依赖。
- 服务是否真的从 `playground` 模块启动。
- 端口是否是 `8080`。
- 是否启用了 `dev` 或 `test` profile。
- 控制台是否出现 actuator endpoint 映射日志。

---

## 📝 许可证

MIT License

---

## 贡献

欢迎提交 Issue 和 Pull Request！

---

## 相关文档与参考

### Apache POI 官方资源

- **Apache POI 官方网站**：https://poi.apache.org/
- **Apache POI 文档中心**：https://poi.apache.org/document/index.html
- **XWPF 用户指南**（处理 Word 文档）：https://poi.apache.org/components/document/quick-guide-xwpf.html
- **POI API 文档**：https://poi.apache.org/apidocs/index.html
- **GitHub 仓库**：https://github.com/apache/poi

### 相关技术文档

- **OOXML 规范**：https://ecma-international.org/publications-and-standards/standards/ecma-376/
- **Java 流式 API**：https://docs.oracle.com/javase/8/docs/api/
- **设计模式 - GOF**：https://en.wikipedia.org/wiki/Design_Patterns

### 社区资源

- **Stack Overflow (tag: apache-poi)**：https://stackoverflow.com/questions/tagged/apache-poi
- **Apache POI 邮件列表**：https://poi.apache.org/help/index.html

---

## 联系方式

如有任何问题或建议，请提交 Issue 或联系项目维护者。

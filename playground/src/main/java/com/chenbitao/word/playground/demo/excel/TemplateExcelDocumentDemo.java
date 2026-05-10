package com.chenbitao.word.playground.demo.excel;

import com.chenbitao.word.excel.TemplateXlsWorkbookGenerator;
import com.chenbitao.word.excel.TemplateXlsxWorkbookGenerator;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * Excel 模板渲染演示。
 */
@Slf4j
public class TemplateExcelDocumentDemo {

    private static final String XLS_OUTPUT_FILE_NAME = "template-sales-demo.xls";
    private static final String XLSX_OUTPUT_FILE_NAME = "template-sales-demo.xlsx";

    /**
     * Excel 模板渲染演示入口。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 如果模板创建、目录创建或文件写出失败
     */
    public static void main(String[] args) throws Exception {
        generate(defaultXlsOutputPath(), defaultXlsxOutputPath());
        log.info("Excel 模板演示文件生成完成");
    }

    /**
     * 生成 XLS 和 XLSX 模板渲染文件。
     *
     * @param xlsOutputPath XLS 输出路径
     * @param xlsxOutputPath XLSX 输出路径
     * @throws Exception 如果模板创建、目录创建或文件写出失败
     */
    public static void generate(Path xlsOutputPath, Path xlsxOutputPath) throws Exception {
        Files.createDirectories(xlsOutputPath.getParent());
        Files.createDirectories(xlsxOutputPath.getParent());

        TemplateXlsWorkbookGenerator xlsGenerator =
                new TemplateXlsWorkbookGenerator(new ByteArrayInputStream(templateBytes(new HSSFWorkbook())));
        xlsGenerator.render(data());
        xlsGenerator.save(xlsOutputPath.toString());

        TemplateXlsxWorkbookGenerator xlsxGenerator =
                new TemplateXlsxWorkbookGenerator(new ByteArrayInputStream(templateBytes(new XSSFWorkbook())));
        xlsxGenerator.render(data());
        xlsxGenerator.save(xlsxOutputPath.toString());

        log.info("XLS 模板文件已保存到：{}", xlsOutputPath.toAbsolutePath());
        log.info("XLSX 模板文件已保存到：{}", xlsxOutputPath.toAbsolutePath());
    }

    /**
     * 创建模板工作簿字节。
     *
     * @param workbook 工作簿
     * @return 模板字节
     * @throws Exception 如果写出失败
     */
    private static byte[] templateBytes(Workbook workbook) throws Exception {
        try (Workbook template = workbook;
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            createTemplate(template);
            template.write(output);
            return output.toByteArray();
        }
    }

    /**
     * 创建模板内容。
     *
     * @param workbook 工作簿
     */
    private static void createTemplate(Workbook workbook) {
        Sheet sheet = workbook.createSheet("模板销售报表");
        Row title = sheet.createRow(0);
        title.createCell(0).setCellValue("销售报告：${report.name}");

        Row header = sheet.createRow(1);
        header.createCell(0).setCellValue("产品");
        header.createCell(1).setCellValue("数量");
        header.createCell(2).setCellValue("单价");
        header.createCell(3).setCellValue("小计");

        Row detail = sheet.createRow(2);
        detail.createCell(0).setCellValue("${product}");
        detail.createCell(1).setCellValue("${quantity}");
        detail.createCell(2).setCellValue("${price}");
        detail.createCell(3).setCellFormula("B3*C3");

        Row note = sheet.createRow(3);
        note.createCell(0).setCellValue("标签：${tags}");
    }

    /**
     * 构造模板数据。
     *
     * @return 模板数据
     */
    private static Map<String, Object> data() {
        Map<String, Object> data = new HashMap<>();
        Map<String, Object> report = new HashMap<>();
        report.put("name", "华东区域");
        data.put("report", report);
        data.put("product", "模板服务");
        data.put("quantity", 12);
        data.put("price", 199.5D);
        data.put("tags", Arrays.asList("重点客户", "续约"));
        return data;
    }

    /**
     * 获取默认 XLS 输出路径。
     *
     * @return 默认 XLS 输出路径
     */
    private static Path defaultXlsOutputPath() {
        return defaultOutputPath(XLS_OUTPUT_FILE_NAME);
    }

    /**
     * 获取默认 XLSX 输出路径。
     *
     * @return 默认 XLSX 输出路径
     */
    private static Path defaultXlsxOutputPath() {
        return defaultOutputPath(XLSX_OUTPUT_FILE_NAME);
    }

    /**
     * 获取默认输出路径。
     *
     * @param fileName 文件名
     * @return 默认输出路径
     */
    private static Path defaultOutputPath(String fileName) {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", fileName);
        }
        return Paths.get("playground", "target", fileName);
    }
}

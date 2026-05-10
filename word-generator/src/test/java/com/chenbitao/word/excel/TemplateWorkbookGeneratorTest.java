package com.chenbitao.word.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * Excel 模板工作簿生成器测试。
 */
public class TemplateWorkbookGeneratorTest {

    /**
     * 测试 XLS 模板可以渲染占位符并保留公式。
     *
     * @throws Exception 如果模板创建或读取失败
     */
    @Test
    public void renderXlsTemplateReplacesPlaceholdersAndKeepsFormula() throws Exception {
        ByteArrayOutputStream template = new ByteArrayOutputStream();
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            createTemplate(workbook);
            workbook.write(template);
        }

        TemplateXlsWorkbookGenerator generator =
                new TemplateXlsWorkbookGenerator(new ByteArrayInputStream(template.toByteArray()));
        generator.render(data());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (HSSFWorkbook workbook = new HSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            assertRenderedWorkbook(workbook);
        }
    }

    /**
     * 测试 XLSX 模板可以渲染占位符并保留公式。
     *
     * @throws Exception 如果模板创建或读取失败
     */
    @Test
    public void renderXlsxTemplateReplacesPlaceholdersAndKeepsFormula() throws Exception {
        ByteArrayOutputStream template = new ByteArrayOutputStream();
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            createTemplate(workbook);
            workbook.write(template);
        }

        TemplateXlsxWorkbookGenerator generator =
                new TemplateXlsxWorkbookGenerator(new ByteArrayInputStream(template.toByteArray()));
        generator.render(data());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            assertRenderedWorkbook(workbook);
        }
    }

    /**
     * 创建测试模板。
     *
     * @param workbook 工作簿
     */
    private void createTemplate(Workbook workbook) {
        Sheet sheet = workbook.createSheet("模板报表");
        Row title = sheet.createRow(0);
        title.createCell(0).setCellValue("销售报告：${report.name}");

        Row row = sheet.createRow(1);
        row.createCell(0).setCellValue("${product}");
        row.createCell(1).setCellValue("${quantity}");
        row.createCell(2).setCellValue("${price}");
        row.createCell(3).setCellFormula("B2*C2");

        Row tags = sheet.createRow(2);
        tags.createCell(0).setCellValue("${tags}");

        Row missing = sheet.createRow(3);
        missing.createCell(0).setCellValue("${missing}");
    }

    /**
     * 构造模板测试数据。
     *
     * @return 模板数据
     */
    private Map<String, Object> data() {
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
     * 断言工作簿已正确渲染。
     *
     * @param workbook 工作簿
     */
    private void assertRenderedWorkbook(Workbook workbook) {
        Sheet sheet = workbook.getSheet("模板报表");
        assertEquals("销售报告：华东区域", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("模板服务", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals(CellType.NUMERIC, sheet.getRow(1).getCell(1).getCellType());
        assertEquals(12D, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.001D);
        assertEquals(199.5D, sheet.getRow(1).getCell(2).getNumericCellValue(), 0.001D);
        assertEquals(CellType.FORMULA, sheet.getRow(1).getCell(3).getCellType());
        assertEquals("B2*C2", sheet.getRow(1).getCell(3).getCellFormula());
        assertEquals("重点客户\n续约", sheet.getRow(2).getCell(0).getStringCellValue());
        assertEquals("${missing}", sheet.getRow(3).getCell(0).getStringCellValue());
    }
}

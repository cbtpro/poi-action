package com.chenbitao.word.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class XlsxWorkbookGeneratorTest {

    @Test
    public void saveWritesXlsxWorkbookWithRowsHeaderStyleAndFormula() throws Exception {
        XlsxWorkbookGenerator generator = new XlsxWorkbookGenerator();
        generator.createWorkbook();
        generator.addHeaderRow("销售数据", Arrays.asList("产品", "数量", "单价", "小计"));
        generator.addRows("销售数据", Arrays.<List<?>>asList(
                Arrays.asList("A", 2, 5.5D),
                Arrays.asList("B", 3, 4D)
        ));
        generator.setFormula("销售数据", 1, 3, "B2*C2");
        generator.setFormula("销售数据", 2, 3, "B3*C3");
        generator.autoSizeColumns("销售数据");

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            Sheet sheet = workbook.getSheet("销售数据");
            assertEquals("产品", sheet.getRow(0).getCell(0).getStringCellValue());
            assertTrue(workbook.getFontAt(sheet.getRow(0).getCell(0).getCellStyle().getFontIndexAsInt()).getBold());
            assertEquals("A", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(2D, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.001D);
            assertEquals(5.5D, sheet.getRow(1).getCell(2).getNumericCellValue(), 0.001D);
            assertEquals(CellType.FORMULA, sheet.getRow(1).getCell(3).getCellType());
            assertEquals("B2*C2", sheet.getRow(1).getCell(3).getCellFormula());
            assertTrue(sheet.getColumnWidth(0) > 0);
        }
    }

    @Test
    public void streamingGeneratorWritesLargeRowsReadableAsXlsx() throws Exception {
        XlsxWorkbookGenerator generator = XlsxWorkbookGenerator.streaming(10);
        generator.addHeaderRow("大数据导出", Arrays.asList("序号", "名称"));
        generator.addRows("大数据导出", rows(250));
        generator.autoSizeColumns("大数据导出");

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            Sheet sheet = workbook.getSheet("大数据导出");
            assertEquals("序号", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(1D, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.001D);
            assertEquals("商品-250", sheet.getRow(250).getCell(1).getStringCellValue());
        }
    }

    @Test
    public void addRowsUsesDefaultSheetWhenNameIsBlank() throws Exception {
        XlsxWorkbookGenerator generator = new XlsxWorkbookGenerator();
        generator.addRows("", Arrays.<List<?>>asList(Arrays.asList("默认工作表")));

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("默认工作表", workbook.getSheet(XlsxWorkbookGenerator.DEFAULT_SHEET_NAME)
                    .getRow(0)
                    .getCell(0)
                    .getStringCellValue());
        }
    }

    private List<List<?>> rows(int count) {
        List<List<?>> rows = new ArrayList<>();
        for (int i = 1; i <= count; i++) {
            rows.add(Arrays.asList(i, "商品-" + i));
        }
        return rows;
    }
}

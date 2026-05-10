package com.chenbitao.word.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class XlsWorkbookGeneratorTest {

    @Test
    public void saveWritesXlsWorkbookWithRowsHeaderStyleAndFormula() throws Exception {
        XlsWorkbookGenerator generator = new XlsWorkbookGenerator();
        generator.createWorkbook();
        generator.addHeaderRow("销售数据", Arrays.asList("产品", "数量", "单价", "小计"));
        generator.addRows("销售数据", Arrays.<java.util.List<?>>asList(
                Arrays.asList("A", 2, 5.5D),
                Arrays.asList("B", 3, 4D)
        ));
        generator.setFormula("销售数据", 1, 3, "B2*C2");
        generator.setFormula("销售数据", 2, 3, "B3*C3");
        generator.autoSizeColumns("销售数据");

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (HSSFWorkbook workbook = new HSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            Sheet sheet = workbook.getSheet("销售数据");
            assertEquals("产品", sheet.getRow(0).getCell(0).getStringCellValue());
            assertTrue(sheet.getRow(0).getCell(0).getCellStyle().getFontIndexAsInt() >= 0);
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
    public void addRowsUsesDefaultSheetWhenNameIsBlank() throws Exception {
        XlsWorkbookGenerator generator = new XlsWorkbookGenerator();
        generator.addRows("", Arrays.<java.util.List<?>>asList(Arrays.asList("默认工作表")));

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (HSSFWorkbook workbook = new HSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("默认工作表", workbook.getSheet(XlsWorkbookGenerator.DEFAULT_SHEET_NAME)
                    .getRow(0)
                    .getCell(0)
                    .getStringCellValue());
        }
    }
}

package com.chenbitao.word.playground.demo.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class XlsSalesReportDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesSalesReportWorkbook() throws Exception {
        File output = temporaryFolder.newFile("excel-sales-demo.xls");

        XlsSalesReportDemo.generate(output.toPath());

        assertTrue(output.length() > 0);
        try (HSSFWorkbook workbook = new HSSFWorkbook(Files.newInputStream(output.toPath()))) {
            Sheet sheet = workbook.getSheet("销售数据");
            assertEquals("产品", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("模板服务", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(12D, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.001D);
            assertEquals(CellType.FORMULA, sheet.getRow(1).getCell(3).getCellType());
            assertEquals("B2*C2", sheet.getRow(1).getCell(3).getCellFormula());
            assertEquals("SUM(D2:D4)", sheet.getRow(4).getCell(3).getCellFormula());
        }
    }
}

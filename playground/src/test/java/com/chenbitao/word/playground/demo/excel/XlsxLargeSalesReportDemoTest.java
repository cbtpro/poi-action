package com.chenbitao.word.playground.demo.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class XlsxLargeSalesReportDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesLargeSalesReportWorkbook() throws Exception {
        File output = temporaryFolder.newFile("excel-large-sales-demo.xlsx");

        XlsxLargeSalesReportDemo.generate(output.toPath(), 25);

        assertTrue(output.length() > 0);
        try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(output.toPath()))) {
            Sheet sheet = workbook.getSheet("销售明细");
            assertEquals("序号", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(1D, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.001D);
            assertEquals("商品-25", sheet.getRow(25).getCell(1).getStringCellValue());
            assertEquals(CellType.FORMULA, sheet.getRow(26).getCell(4).getCellType());
            assertEquals("SUMPRODUCT(C2:C26,D2:D26)", sheet.getRow(26).getCell(4).getCellFormula());
        }
    }
}

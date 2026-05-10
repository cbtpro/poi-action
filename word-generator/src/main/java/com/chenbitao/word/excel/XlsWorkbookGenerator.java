package com.chenbitao.word.excel;

import com.chenbitao.word.exception.SpreadsheetException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;

/**
 * Excel 97-2003 工作簿生成器。
 *
 * <p>基于 Apache POI HSSF 生成 {@code .xls} 文件，支持基础表格数据写入、
 * 表头样式、公式和自动列宽。</p>
 */
public class XlsWorkbookGenerator implements ExcelGenerator {

    /** 默认工作表名称 */
    public static final String DEFAULT_SHEET_NAME = "Sheet1";

    /** HSSF 工作簿对象 */
    private Workbook workbook;

    /** 默认表头样式 */
    private CellStyle headerStyle;

    /**
     * 创建新的 XLS 工作簿。
     */
    @Override
    public void createWorkbook() {
        workbook = new HSSFWorkbook();
        headerStyle = createHeaderStyle();
    }

    /**
     * 创建工作表。
     *
     * @param sheetName 工作表名称
     */
    @Override
    public void createSheet(String sheetName) {
        ensureWorkbook();
        String name = normalizeSheetName(sheetName);
        if (workbook.getSheet(name) == null) {
            workbook.createSheet(name);
        }
    }

    /**
     * 添加带表头样式的第一行数据。
     *
     * @param sheetName 工作表名称
     * @param headers 表头文本
     */
    @Override
    public void addHeaderRow(String sheetName, List<String> headers) {
        Sheet sheet = getOrCreateSheet(sheetName);
        Row row = getOrCreateRow(sheet, nextRowIndex(sheet));
        for (int i = 0; headers != null && i < headers.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers.get(i) == null ? "" : headers.get(i));
            cell.setCellStyle(headerStyle);
        }
    }

    /**
     * 追加多行表格数据。
     *
     * @param sheetName 工作表名称
     * @param rows 表格数据
     */
    @Override
    public void addRows(String sheetName, List<List<?>> rows) {
        Sheet sheet = getOrCreateSheet(sheetName);
        if (rows == null || rows.isEmpty()) {
            return;
        }

        int rowIndex = nextRowIndex(sheet);
        for (List<?> rowData : rows) {
            Row row = getOrCreateRow(sheet, rowIndex++);
            writeRow(row, rowData);
        }
    }

    /**
     * 设置单元格公式。
     *
     * @param sheetName 工作表名称
     * @param rowIndex 行下标，从 0 开始
     * @param columnIndex 列下标，从 0 开始
     * @param formula Excel 公式，不包含等号
     */
    @Override
    public void setFormula(String sheetName, int rowIndex, int columnIndex, String formula) {
        Sheet sheet = getOrCreateSheet(sheetName);
        Cell cell = getOrCreateRow(sheet, rowIndex).createCell(columnIndex);
        cell.setCellFormula(formula == null ? "" : formula);
    }

    /**
     * 自动调整已使用列宽。
     *
     * @param sheetName 工作表名称
     */
    @Override
    public void autoSizeColumns(String sheetName) {
        Sheet sheet = getOrCreateSheet(sheetName);
        int maxColumn = maxColumnIndex(sheet);
        for (int i = 0; i <= maxColumn; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * 保存工作簿到文件。
     *
     * @param path 输出文件路径
     */
    @Override
    public void save(String path) {
        ensureWorkbook();
        try (FileOutputStream output = new FileOutputStream(path)) {
            save(output);
        } catch (Exception e) {
            throw new SpreadsheetException("保存 XLS 失败", e);
        }
    }

    /**
     * 保存工作簿到输出流。
     *
     * @param outputStream 输出流
     */
    @Override
    public void save(OutputStream outputStream) {
        ensureWorkbook();
        try {
            workbook.write(outputStream);
        } catch (Exception e) {
            throw new SpreadsheetException("写出 XLS 失败", e);
        }
    }

    private CellStyle createHeaderStyle() {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private Sheet getOrCreateSheet(String sheetName) {
        ensureWorkbook();
        String name = normalizeSheetName(sheetName);
        Sheet sheet = workbook.getSheet(name);
        return sheet == null ? workbook.createSheet(name) : sheet;
    }

    private String normalizeSheetName(String sheetName) {
        return sheetName == null || sheetName.trim().isEmpty() ? DEFAULT_SHEET_NAME : sheetName.trim();
    }

    private int nextRowIndex(Sheet sheet) {
        return sheet.getPhysicalNumberOfRows() == 0 ? 0 : sheet.getLastRowNum() + 1;
    }

    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row == null ? sheet.createRow(rowIndex) : row;
    }

    private void writeRow(Row row, List<?> rowData) {
        if (rowData == null || rowData.isEmpty()) {
            return;
        }
        for (int i = 0; i < rowData.size(); i++) {
            writeCell(row.createCell(i), rowData.get(i));
        }
    }

    private void writeCell(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private int maxColumnIndex(Sheet sheet) {
        int maxColumn = -1;
        for (Row row : sheet) {
            if (row.getLastCellNum() > maxColumn) {
                maxColumn = row.getLastCellNum() - 1;
            }
        }
        return maxColumn;
    }

    private void ensureWorkbook() {
        if (workbook == null) {
            createWorkbook();
        }
    }
}

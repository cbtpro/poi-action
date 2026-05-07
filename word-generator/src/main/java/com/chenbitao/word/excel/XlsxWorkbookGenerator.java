package com.chenbitao.word.excel;

import com.chenbitao.word.exception.SpreadsheetException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;

/**
 * Excel Open XML 工作簿生成器。
 *
 * <p>默认基于 XSSF 生成 {@code .xlsx} 文件；通过 {@link #streaming(int)} 可创建
 * SXSSF 流式工作簿，适合大数据量写出。</p>
 */
public class XlsxWorkbookGenerator implements ExcelGenerator {

    /** 默认工作表名称 */
    public static final String DEFAULT_SHEET_NAME = "Sheet1";

    /** 默认 SXSSF 内存窗口行数 */
    public static final int DEFAULT_STREAMING_WINDOW_SIZE = 100;

    /** 工作簿对象，可为 XSSFWorkbook 或 SXSSFWorkbook */
    private Workbook workbook;

    /** 是否使用 SXSSF 流式写出 */
    private final boolean streaming;

    /** SXSSF 内存窗口行数 */
    private final int rowAccessWindowSize;

    /** 默认表头样式 */
    private CellStyle headerStyle;

    /**
     * 创建基于 XSSF 的 XLSX 生成器。
     */
    public XlsxWorkbookGenerator() {
        this(false, DEFAULT_STREAMING_WINDOW_SIZE);
    }

    private XlsxWorkbookGenerator(boolean streaming, int rowAccessWindowSize) {
        this.streaming = streaming;
        this.rowAccessWindowSize = Math.max(1, rowAccessWindowSize);
    }

    /**
     * 创建基于 SXSSF 的流式 XLSX 生成器。
     *
     * @param rowAccessWindowSize SXSSF 在内存中保留的行数窗口
     * @return 流式 XLSX 生成器
     */
    public static XlsxWorkbookGenerator streaming(int rowAccessWindowSize) {
        return new XlsxWorkbookGenerator(true, rowAccessWindowSize);
    }

    /**
     * 创建新的 XLSX 工作簿。
     */
    @Override
    public void createWorkbook() {
        workbook = streaming ? new SXSSFWorkbook(rowAccessWindowSize) : new XSSFWorkbook();
        headerStyle = createHeaderStyle();
    }

    /**
     * 创建工作表。
     *
     * @param sheetName 工作表名称
     */
    @Override
    public void createSheet(String sheetName) {
        Sheet sheet = getOrCreateSheet(sheetName);
        trackColumnsForStreaming(sheet);
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
            throw new SpreadsheetException("保存 XLSX 失败", e);
        } finally {
            disposeStreamingWorkbook();
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
            throw new SpreadsheetException("写出 XLSX 失败", e);
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
        if (sheet == null) {
            sheet = workbook.createSheet(name);
            trackColumnsForStreaming(sheet);
        }
        return sheet;
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

    private void trackColumnsForStreaming(Sheet sheet) {
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        }
    }

    private void disposeStreamingWorkbook() {
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }
    }

    private void ensureWorkbook() {
        if (workbook == null) {
            createWorkbook();
        }
    }
}

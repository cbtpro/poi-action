package com.chenbitao.word.excel;

import java.io.OutputStream;
import java.util.List;

/**
 * Excel 工作簿生成器接口。
 *
 * <p>该接口描述 Excel 生成的通用能力，具体格式由实现类决定，例如 HSSF 对应
 * Excel 97-2003 的 {@code .xls} 文件。</p>
 */
public interface ExcelGenerator {

    /**
     * 创建新的工作簿实例。
     */
    void createWorkbook();

    /**
     * 创建工作表。
     *
     * @param sheetName 工作表名称
     */
    void createSheet(String sheetName);

    /**
     * 添加表头行，并应用默认表头样式。
     *
     * @param sheetName 工作表名称
     * @param headers 表头文本
     */
    void addHeaderRow(String sheetName, List<String> headers);

    /**
     * 追加多行表格数据。
     *
     * @param sheetName 工作表名称
     * @param rows 表格数据
     */
    void addRows(String sheetName, List<List<?>> rows);

    /**
     * 设置单元格公式。
     *
     * @param sheetName 工作表名称
     * @param rowIndex 行下标，从 0 开始
     * @param columnIndex 列下标，从 0 开始
     * @param formula Excel 公式，不包含等号
     */
    void setFormula(String sheetName, int rowIndex, int columnIndex, String formula);

    /**
     * 自动调整工作表已使用列的宽度。
     *
     * @param sheetName 工作表名称
     */
    void autoSizeColumns(String sheetName);

    /**
     * 保存工作簿到指定路径。
     *
     * @param path 输出文件路径
     */
    void save(String path);

    /**
     * 保存工作簿到输出流。
     *
     * @param outputStream 输出流
     */
    void save(OutputStream outputStream);
}

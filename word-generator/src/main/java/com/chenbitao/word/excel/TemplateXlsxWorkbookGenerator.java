package com.chenbitao.word.excel;

import com.chenbitao.word.exception.SpreadsheetException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;

/**
 * XLSX 模板工作簿生成器。
 *
 * <p>基于 Apache POI XSSF 渲染 Excel Open XML 的 {@code .xlsx} 模板。</p>
 */
public class TemplateXlsxWorkbookGenerator extends AbstractTemplateWorkbookGenerator {

    /**
     * 构造 XLSX 模板工作簿生成器。
     *
     * @param template 模板输入流
     * @throws SpreadsheetException 如果模板加载失败
     */
    public TemplateXlsxWorkbookGenerator(InputStream template) {
        super(loadWorkbook(template));
    }

    /**
     * 加载 XSSF 工作簿。
     *
     * @param template 模板输入流
     * @return XSSF 工作簿
     */
    private static XSSFWorkbook loadWorkbook(InputStream template) {
        try {
            return new XSSFWorkbook(template);
        } catch (Exception e) {
            throw new SpreadsheetException("加载 XLSX 模板失败", e);
        }
    }
}

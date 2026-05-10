package com.chenbitao.word.excel;

import com.chenbitao.word.exception.SpreadsheetException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.InputStream;

/**
 * XLS 模板工作簿生成器。
 *
 * <p>基于 Apache POI HSSF 渲染 Excel 97-2003 的 {@code .xls} 模板。</p>
 */
public class TemplateXlsWorkbookGenerator extends AbstractTemplateWorkbookGenerator {

    /**
     * 构造 XLS 模板工作簿生成器。
     *
     * @param template 模板输入流
     * @throws SpreadsheetException 如果模板加载失败
     */
    public TemplateXlsWorkbookGenerator(InputStream template) {
        super(loadWorkbook(template));
    }

    /**
     * 加载 HSSF 工作簿。
     *
     * @param template 模板输入流
     * @return HSSF 工作簿
     */
    private static HSSFWorkbook loadWorkbook(InputStream template) {
        try {
            return new HSSFWorkbook(template);
        } catch (Exception e) {
            throw new SpreadsheetException("加载 XLS 模板失败", e);
        }
    }
}

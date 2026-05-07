package com.chenbitao.word.factory;

import com.chenbitao.word.excel.ExcelGenerator;
import com.chenbitao.word.excel.XlsWorkbookGenerator;
import com.chenbitao.word.excel.XlsxWorkbookGenerator;

/**
 * 电子表格生成器工厂。
 *
 * <p>当前支持 Excel 97-2003 的 {@code xls} 和 Excel Open XML 的 {@code xlsx} 类型。</p>
 */
public class SpreadsheetGeneratorFactory {

    private SpreadsheetGeneratorFactory() {
    }

    /**
     * 获取指定类型的 Excel 生成器。
     *
     * @param type 文件类型，目前支持 {@code xls} 和 {@code xlsx}
     * @return 新的 Excel 生成器实例
     */
    public static ExcelGenerator get(String type) {
        if ("xls".equalsIgnoreCase(type)) {
            return new XlsWorkbookGenerator();
        }
        if ("xlsx".equalsIgnoreCase(type)) {
            return new XlsxWorkbookGenerator();
        }
        throw new IllegalArgumentException("不支持电子表格类型: " + type);
    }
}

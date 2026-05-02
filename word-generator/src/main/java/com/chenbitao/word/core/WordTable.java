package com.chenbitao.word.core;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * 编程式Word表格模型。
 * 用于描述包含文本、图片和单元格跨行跨列的表格结构。
 */
public class WordTable {

    /** 表格行数据 */
    private final List<List<WordTableCell>> rows;

    /**
     * 创建表格模型。
     *
     * @param rows 表格行数据
     */
    public WordTable(List<List<WordTableCell>> rows) {
        this.rows = copyRows(rows);
    }

    /**
     * 创建表格模型。
     *
     * @param rows 表格行数据
     * @return 表格模型
     */
    public static WordTable of(List<List<WordTableCell>> rows) {
        return new WordTable(rows);
    }

    /**
     * 获取表格行数据。
     *
     * @return 不可变表格行数据
     */
    public List<List<WordTableCell>> getRows() {
        return rows;
    }

    private List<List<WordTableCell>> copyRows(List<List<WordTableCell>> rows) {
        if (rows == null || rows.isEmpty()) {
            return Collections.emptyList();
        }

        List<List<WordTableCell>> result = new ArrayList<>();
        for (List<WordTableCell> row : rows) {
            if (row == null) {
                result.add(Collections.<WordTableCell>emptyList());
            } else {
                result.add(Collections.unmodifiableList(new ArrayList<>(row)));
            }
        }
        return Collections.unmodifiableList(result);
    }
}

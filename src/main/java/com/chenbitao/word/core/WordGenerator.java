package com.chenbitao.word.core;

import java.io.InputStream;

public interface WordGenerator {

    /**
     * 创建一个新的 Word 文档实例。
     */
    void createDocument();

    /**
     * 添加普通文本段落。
     *
     * @param text 段落文本
     */
    void addParagraph(String text);

    /**
     * 添加标题段落。
     *
     * @param text 标题文本
     * @param level 标题级别
     */
    void addTitle(String text, int level);

    /**
     * 添加表格。
     *
     * @param rows 表格行数
     * @param cols 表格列数
     */
    void addTable(int rows, int cols);

    /**
     * 添加图片。
     *
     * @param inputStream 图片输入流
     * @param width 图片宽度
     * @param height 图片高度
     */
    void addImage(InputStream inputStream, int width, int height);

    /**
     * 保存 Word 文档到指定路径。
     *
     * @param path 输出文件路径
     */
    void save(String path);
}

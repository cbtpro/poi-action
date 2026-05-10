package com.chenbitao.word.pdf;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * PDF 文档生成器接口。
 */
public interface PdfDocumentGenerator {

    /**
     * 创建新的 PDF 文档。
     */
    void createDocument();

    /**
     * 添加标题。
     *
     * @param text 标题文本
     */
    void addTitle(String text);

    /**
     * 添加段落。
     *
     * @param text 段落文本
     */
    void addParagraph(String text);

    /**
     * 添加表格。
     *
     * @param rows 表格数据，每个内部列表代表一行
     */
    void addTable(List<List<String>> rows);

    /**
     * 添加图片。
     *
     * @param image 图片输入流
     * @param width 图片宽度
     * @param height 图片高度
     */
    void addImage(InputStream image, float width, float height);

    /**
     * 保存 PDF 到指定路径。
     *
     * @param path 输出文件路径
     */
    void save(String path);

    /**
     * 保存 PDF 到输出流。
     *
     * @param outputStream 输出流
     */
    void save(OutputStream outputStream);
}

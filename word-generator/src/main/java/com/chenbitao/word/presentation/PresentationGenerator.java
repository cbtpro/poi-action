package com.chenbitao.word.presentation;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * PowerPoint 演示文稿生成器接口。
 *
 * <p>该接口描述演示文稿生成的通用能力，具体格式由实现类决定，例如 HSLF 对应
 * PowerPoint 97-2003 的 {@code .ppt} 文件。</p>
 */
public interface PresentationGenerator {

    /**
     * 创建新的演示文稿实例。
     */
    void createPresentation();

    /**
     * 添加封面页。
     *
     * @param title 标题
     * @param subtitle 副标题
     */
    void addTitleSlide(String title, String subtitle);

    /**
     * 添加文本页。
     *
     * @param title 标题
     * @param bulletItems 项目符号文本
     */
    void addTextSlide(String title, List<String> bulletItems);

    /**
     * 添加表格页。
     *
     * @param title 标题
     * @param rows 表格数据，每个内部列表代表一行
     */
    void addTableSlide(String title, List<List<String>> rows);

    /**
     * 添加图片页。
     *
     * @param title 标题
     * @param image 图片输入流
     */
    void addImageSlide(String title, InputStream image);

    /**
     * 保存演示文稿到指定路径。
     *
     * @param path 输出文件路径
     */
    void save(String path);

    /**
     * 保存演示文稿到输出流。
     *
     * @param outputStream 输出流
     */
    void save(OutputStream outputStream);
}

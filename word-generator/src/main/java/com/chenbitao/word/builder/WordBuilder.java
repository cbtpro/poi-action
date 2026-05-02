package com.chenbitao.word.builder;

import com.chenbitao.word.core.WordGenerator;
import com.chenbitao.word.core.WordTable;

import java.io.InputStream;
import java.util.List;

/**
 * Word文档建造者类
 * 使用建造者模式提供流式API来构建Word文档，支持添加标题、段落、列表和表格等元素
 */
public class WordBuilder {

    private final WordGenerator generator;

    /**
     * 构造Word文档建造者
     *
     * @param generator Word文档生成器实例
     */
    public WordBuilder(WordGenerator generator) {
        this.generator = generator;
        generator.createDocument();
    }

    /**
     * 添加一级标题
     *
     * @param text 标题文本
     * @return 当前建造者实例，支持链式调用
     */
    public WordBuilder title(String text) {
        generator.addTitle(text, 1);
        return this;
    }

    /**
     * 添加段落
     *
     * @param text 段落文本
     * @return 当前建造者实例，支持链式调用
     */
    public WordBuilder paragraph(String text) {
        generator.addParagraph(text);
        return this;
    }

    /**
     * 添加段落列表
     * 将可迭代对象中的每个元素转换为带项目符号的段落
     *
     * @param list 包含段落文本的可迭代对象
     * @return 当前建造者实例，支持链式调用
     */
    public WordBuilder paragraphList(Iterable<String> list) {
        for (String item : list) {
            generator.addParagraph("• " + item);
        }
        return this;
    }

    /**
     * 添加表格
     *
     * @param r 表格行数
     * @param c 表格列数
     * @return 当前建造者实例，支持链式调用
     */
    public WordBuilder table(int r, int c) {
        generator.addTable(r, c);
        return this;
    }

    /**
     * 添加带内容的表格
     *
     * @param rows 表格数据，每个内部列表代表一行
     * @return 当前建造者实例，支持链式调用
     */
    public WordBuilder table(List<List<String>> rows) {
        generator.addTable(rows);
        return this;
    }

    /**
     * 添加结构化表格
     *
     * @param table 表格模型
     * @return 当前建造者实例，支持链式调用
     */
    public WordBuilder table(WordTable table) {
        generator.addTable(table);
        return this;
    }

    /**
     * 添加图片
     *
     * @param inputStream 图片输入流
     * @param width 图片宽度（EMU单位）
     * @param height 图片高度（EMU单位）
     * @return 当前建造者实例，支持链式调用
     */
    public WordBuilder image(InputStream inputStream, int width, int height) {
        generator.addImage(inputStream, width, height);
        return this;
    }

    /**
     * 构建并保存文档
     *
     * @param path 文档保存路径
     */
    public void build(String path) {
        generator.save(path);
    }
}

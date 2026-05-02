package com.chenbitao.playground.builder;

import com.chenbitao.word.builder.WordBuilder;
import com.chenbitao.word.core.WordGenerator;
import org.junit.Test;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertSame;

/**
 * WordBuilder建造者类测试
 * 测试WordBuilder的流式API功能，包括文档创建、标题添加、段落添加、列表添加和表格添加等操作
 */
public class WordBuilderTest {

    /**
     * 测试构造方法创建文档并按顺序委托调用
     * 验证WordBuilder的构造函数会创建文档，并且所有方法调用都会按正确的顺序委托给底层生成器
     */
    @Test
    public void constructorCreatesDocumentAndDelegatesCallsInOrder() {
        RecordingWordGenerator generator = new RecordingWordGenerator();

        WordBuilder builder = new WordBuilder(generator);
        WordBuilder returned = builder
                .title("标题")
                .paragraph("正文段落")
                .paragraphList(Arrays.asList("第一项", "第二项"))
                .table(2, 3);
        builder.build("target/测试输出.docx");

        assertSame(builder, returned);
        assertEquals("createDocument", generator.calls.get(0));
        assertEquals("addTitle:标题:1", generator.calls.get(1));
        assertEquals("addParagraph:正文段落", generator.calls.get(2));
        assertEquals("addTable:2:3", generator.calls.get(5));
        assertEquals("save:target/测试输出.docx", generator.calls.get(6));
        assertEquals(7, generator.calls.size());
        assertEquals(Arrays.asList(
                "createDocument",
                "addTitle:标题:1",
                "addParagraph:正文段落",
                generator.calls.get(3),
                generator.calls.get(4),
                "addTable:2:3",
                "save:target/测试输出.docx"
        ), generator.calls);
        assertEquals("第一项", generator.calls.get(3).substring(generator.calls.get(3).length() - "第一项".length()));
        assertEquals("第二项", generator.calls.get(4).substring(generator.calls.get(4).length() - "第二项".length()));
    }

    /**
     * 记录调用的Word生成器实现
     * 用于测试的WordGenerator实现，记录所有方法调用以便验证调用顺序和参数
     */
    private static class RecordingWordGenerator implements WordGenerator {
        /** 记录的方法调用列表 */
        private final List<String> calls = new ArrayList<>();

        /**
         * 创建文档
         */
        @Override
        public void createDocument() {
            calls.add("createDocument");
        }

        /**
         * 添加段落
         *
         * @param text 段落文本
         */
        @Override
        public void addParagraph(String text) {
            calls.add("addParagraph:" + text);
        }

        /**
         * 添加标题
         *
         * @param text 标题文本
         * @param level 标题级别
         */
        @Override
        public void addTitle(String text, int level) {
            calls.add("addTitle:" + text + ":" + level);
        }

        /**
         * 添加表格
         *
         * @param rows 行数
         * @param cols 列数
         */
        @Override
        public void addTable(int rows, int cols) {
            calls.add("addTable:" + rows + ":" + cols);
        }

        /**
         * 添加图片
         *
         * @param inputStream 图片输入流
         * @param width 图片宽度
         * @param height 图片高度
         */
        @Override
        public void addImage(InputStream inputStream, int width, int height) {
            calls.add("addImage:" + width + ":" + height);
        }

        /**
         * 保存文档
         *
         * @param path 保存路径
         */
        @Override
        public void save(String path) {
            calls.add("save:" + path);
        }
    }
}

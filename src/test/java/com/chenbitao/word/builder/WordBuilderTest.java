package com.chenbitao.word.builder;

import com.chenbitao.word.core.WordGenerator;
import org.junit.Test;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertSame;

public class WordBuilderTest {

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

    private static class RecordingWordGenerator implements WordGenerator {
        private final List<String> calls = new ArrayList<>();

        @Override
        public void createDocument() {
            calls.add("createDocument");
        }

        @Override
        public void addParagraph(String text) {
            calls.add("addParagraph:" + text);
        }

        @Override
        public void addTitle(String text, int level) {
            calls.add("addTitle:" + text + ":" + level);
        }

        @Override
        public void addTable(int rows, int cols) {
            calls.add("addTable:" + rows + ":" + cols);
        }

        @Override
        public void addImage(InputStream inputStream, int width, int height) {
            calls.add("addImage:" + width + ":" + height);
        }

        @Override
        public void save(String path) {
            calls.add("save:" + path);
        }
    }
}

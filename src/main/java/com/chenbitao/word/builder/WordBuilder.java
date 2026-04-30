package com.chenbitao.word.builder;

import com.chenbitao.word.core.WordGenerator;

public class WordBuilder {

    private final WordGenerator generator;

    public WordBuilder(WordGenerator generator) {
        this.generator = generator;
        generator.createDocument();
    }

    public WordBuilder title(String text) {
        generator.addTitle(text, 1);
        return this;
    }

    public WordBuilder paragraph(String text) {
        generator.addParagraph(text);
        return this;
    }

    public WordBuilder paragraphList(Iterable<String> list) {
        for (String item : list) {
            generator.addParagraph("• " + item);
        }
        return this;
    }

    public WordBuilder table(int r, int c) {
        generator.addTable(r, c);
        return this;
    }

    public void build(String path) {
        generator.save(path);
    }
}
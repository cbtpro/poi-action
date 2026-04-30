package com.chenbitao.word.factory;

import com.chenbitao.word.core.WordGenerator;
import com.chenbitao.word.doc.DocWordGenerator;
import com.chenbitao.word.docx.DocxWordGenerator;

import java.util.HashMap;
import java.util.Map;

public class WordGeneratorFactory {

    private static final Map<String, WordGenerator> CACHE = new HashMap<>();

    static {
        CACHE.put("doc", new DocWordGenerator());
        CACHE.put("docx", new DocxWordGenerator());
    }

    public static WordGenerator get(String type) {
        WordGenerator generator = CACHE.get(type.toLowerCase());
        if (generator == null) {
            throw new IllegalArgumentException("不支持类型: " + type);
        }
        return generator;
    }
}
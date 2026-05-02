package com.chenbitao.word.factory;

import com.chenbitao.word.core.WordGenerator;
import com.chenbitao.word.doc.DocWordGenerator;
import com.chenbitao.word.docx.DocxWordGenerator;

import java.util.HashMap;
import java.util.Map;

/**
 * Word文档生成器工厂类
 * 使用工厂模式和缓存机制提供不同类型的Word文档生成器实例
 * 支持DOC和DOCX格式的文档生成
 */
public class WordGeneratorFactory {

    private WordGeneratorFactory() {}

    /** 文档生成器缓存 */
    private static final Map<String, WordGenerator> CACHE = new HashMap<>();

    static {
        CACHE.put("doc", new DocWordGenerator());
        CACHE.put("docx", new DocxWordGenerator());
    }

    /**
     * 获取指定类型的Word文档生成器
     *
     * @param type 文档类型，支持"doc"和"docx"（不区分大小写）
     * @return 对应的Word文档生成器实例
     * @throws IllegalArgumentException 如果指定的类型不支持
     */
    public static WordGenerator get(String type) {
        WordGenerator generator = CACHE.get(type.toLowerCase());
        if (generator == null) {
            throw new IllegalArgumentException("不支持类型: " + type);
        }
        return generator;
    }
}
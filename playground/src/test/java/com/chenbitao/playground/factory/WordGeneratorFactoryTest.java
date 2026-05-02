package com.chenbitao.playground.factory;

import com.chenbitao.word.doc.DocWordGenerator;
import com.chenbitao.word.docx.DocxWordGenerator;
import com.chenbitao.word.factory.WordGeneratorFactory;
import org.junit.Test;

import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertTrue;

/**
 * WordGeneratorFactory工厂类测试
 * 测试WordGeneratorFactory的工厂方法功能，包括类型匹配、缓存机制和异常处理
 */
public class WordGeneratorFactoryTest {

    /**
     * 测试get方法根据类型返回缓存的生成器实例（忽略大小写）
     * 验证工厂方法能够正确识别"doc"和"docx"类型（大小写不敏感），
     * 并返回缓存的生成器实例
     */
    @Test
    public void getReturnsCachedGeneratorByTypeIgnoringCase() {
        assertTrue(WordGeneratorFactory.get("doc") instanceof DocWordGenerator);
        assertTrue(WordGeneratorFactory.get("DOCX") instanceof DocxWordGenerator);
        assertSame(WordGeneratorFactory.get("docx"), WordGeneratorFactory.get("DOCX"));
    }

    /**
     * 测试get方法对不支持的类型抛出异常
     * 验证工厂方法在遇到不支持的文档类型时会抛出IllegalArgumentException
     */
    @Test(expected = IllegalArgumentException.class)
    public void getThrowsExceptionForUnsupportedType() {
        WordGeneratorFactory.get("pdf");
    }
}

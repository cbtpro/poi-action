package com.chenbitao.word.factory;

import com.chenbitao.word.doc.DocWordGenerator;
import com.chenbitao.word.docx.DocxWordGenerator;
import org.junit.Test;

import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertTrue;

public class WordGeneratorFactoryTest {

    @Test
    public void getReturnsCachedGeneratorByTypeIgnoringCase() {
        assertTrue(WordGeneratorFactory.get("doc") instanceof DocWordGenerator);
        assertTrue(WordGeneratorFactory.get("DOCX") instanceof DocxWordGenerator);
        assertSame(WordGeneratorFactory.get("docx"), WordGeneratorFactory.get("DOCX"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void getThrowsExceptionForUnsupportedType() {
        WordGeneratorFactory.get("pdf");
    }
}

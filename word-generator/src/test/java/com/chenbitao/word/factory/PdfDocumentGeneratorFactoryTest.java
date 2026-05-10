package com.chenbitao.word.factory;

import com.chenbitao.word.pdf.PdfBoxDocumentGenerator;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class PdfDocumentGeneratorFactoryTest {

    @Test
    public void getReturnsPdfGeneratorByTypeIgnoringCase() {
        assertTrue(PdfDocumentGeneratorFactory.get("PDF") instanceof PdfBoxDocumentGenerator);
    }

    @Test(expected = IllegalArgumentException.class)
    public void getThrowsExceptionForUnsupportedType() {
        PdfDocumentGeneratorFactory.get("txt");
    }
}

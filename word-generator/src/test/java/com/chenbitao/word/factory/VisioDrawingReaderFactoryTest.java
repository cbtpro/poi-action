package com.chenbitao.word.factory;

import com.chenbitao.word.visio.VisioDrawingReader;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class VisioDrawingReaderFactoryTest {

    @Test
    public void getReturnsReaderForSupportedTypes() {
        assertTrue(VisioDrawingReaderFactory.get("vsd") instanceof VisioDrawingReader);
        assertTrue(VisioDrawingReaderFactory.get("vsdx") instanceof VisioDrawingReader);
    }

    @Test(expected = IllegalArgumentException.class)
    public void getRejectsUnsupportedType() {
        VisioDrawingReaderFactory.get("docx");
    }
}

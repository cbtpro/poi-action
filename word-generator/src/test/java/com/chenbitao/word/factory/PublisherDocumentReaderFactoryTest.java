package com.chenbitao.word.factory;

import com.chenbitao.word.publisher.PublisherDocumentReader;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class PublisherDocumentReaderFactoryTest {

    @Test
    public void getReturnsReaderForPubType() {
        assertTrue(PublisherDocumentReaderFactory.get("pub") instanceof PublisherDocumentReader);
    }

    @Test(expected = IllegalArgumentException.class)
    public void getRejectsUnsupportedType() {
        PublisherDocumentReaderFactory.get("doc");
    }
}

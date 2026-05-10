package com.chenbitao.word.factory;

import com.chenbitao.word.publisher.PublisherDocumentReader;

/**
 * Publisher document reader factory.
 */
public class PublisherDocumentReaderFactory {

    private PublisherDocumentReaderFactory() {
    }

    /**
     * Get a Publisher document reader for the specified file type.
     *
     * @param type file type, currently supports {@code pub}
     * @return Publisher document reader
     */
    public static PublisherDocumentReader get(String type) {
        if ("pub".equalsIgnoreCase(type)) {
            return new PublisherDocumentReader();
        }
        throw new IllegalArgumentException("Unsupported Publisher document type: " + type);
    }
}

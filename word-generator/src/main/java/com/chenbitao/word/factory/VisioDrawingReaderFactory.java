package com.chenbitao.word.factory;

import com.chenbitao.word.visio.VisioDrawingReader;

/**
 * Visio drawing reader factory.
 */
public class VisioDrawingReaderFactory {

    private VisioDrawingReaderFactory() {
    }

    /**
     * Get a Visio drawing reader for the specified file type.
     *
     * @param type file type, currently supports {@code vsd} and {@code vsdx}
     * @return Visio drawing reader
     */
    public static VisioDrawingReader get(String type) {
        if ("vsd".equalsIgnoreCase(type) || "vsdx".equalsIgnoreCase(type)) {
            return new VisioDrawingReader();
        }
        throw new IllegalArgumentException("Unsupported Visio drawing type: " + type);
    }
}

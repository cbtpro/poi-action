package com.chenbitao.word.exception;

/**
 * PDF 文档生成异常。
 *
 * <p>用于封装 PDF 创建、绘制和写出过程中的底层异常。</p>
 */
public class PdfDocumentException extends RuntimeException {

    public PdfDocumentException(String message) {
        super(message);
    }

    public PdfDocumentException(String message, Throwable cause) {
        super(message, cause);
    }
}

package com.chenbitao.word.factory;

import com.chenbitao.word.pdf.PdfBoxDocumentGenerator;
import com.chenbitao.word.pdf.PdfDocumentGenerator;

/**
 * PDF 文档生成器工厂。
 *
 * <p>当前支持 PDFBox 的 {@code pdf} 类型。</p>
 */
public class PdfDocumentGeneratorFactory {

    private PdfDocumentGeneratorFactory() {
    }

    /**
     * 获取指定类型的 PDF 生成器。
     *
     * @param type 文件类型，目前支持 {@code pdf}
     * @return 新的 PDF 生成器实例
     */
    public static PdfDocumentGenerator get(String type) {
        if ("pdf".equalsIgnoreCase(type)) {
            return new PdfBoxDocumentGenerator();
        }
        throw new IllegalArgumentException("不支持 PDF 文档类型: " + type);
    }
}

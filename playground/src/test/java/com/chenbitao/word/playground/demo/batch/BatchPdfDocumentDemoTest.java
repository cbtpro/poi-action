package com.chenbitao.word.playground.demo.batch;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class BatchPdfDocumentDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesReadablePdfBatch() throws Exception {
        Path outputDirectory = temporaryFolder.newFolder("pdf-batch").toPath();

        BatchPdfDocumentDemo.generate(1, 1, outputDirectory);

        assertEquals("应只生成 1 个 PDF 文件",
                1,
                Files.list(outputDirectory).filter(path -> path.toString().endsWith(".pdf")).count());
        File first = outputDirectory.resolve("batch-pdf-00001.pdf").toFile();
        assertTrue("批量生成的 PDF 文件应为非空", first.length() > 0);
        try (PDDocument document = PDDocument.load(first)) {
            assertEquals("批量 PDF 应复用 100 页报告数据", 100, document.getNumberOfPages());
            String text = new PDFTextStripper().getText(document);
            assertTrue("应包含 PDF 报告标题", text.contains("poi-action PDF Report"));
            assertTrue("应包含第一页组合报告内容", text.contains("Portfolio Review 1"));
            assertTrue("应包含第一百页组合报告内容", text.contains("Portfolio Review 100"));
            assertTrue("应包含第一百页页脚", text.contains("Page 100 of 100"));
        }
    }
}

package com.chenbitao.word.playground.demo.pdf;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class PdfProjectReportDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesProjectReportPdf() throws Exception {
        File output = temporaryFolder.newFile("pdf-report-demo.pdf");

        PdfProjectReportDemo.generate(output.toPath());

        assertTrue("应生成非空 PDF 文件", output.length() > 0);
        try (PDDocument document = PDDocument.load(output)) {
            assertEquals("应生成 100 页 PDF", 100, document.getNumberOfPages());
            String text = new PDFTextStripper().getText(document);
            assertTrue("应包含 PDF 报告标题", text.contains("poi-action PDF Report"));
            assertTrue("应包含第一页组合报告内容", text.contains("Portfolio Review 1"));
            assertTrue("应包含第一百页组合报告内容", text.contains("Portfolio Review 100"));
            assertTrue("应包含第一百页页脚", text.contains("Page 100 of 100"));
            assertTrue("应包含收入指标卡片", text.contains("Revenue"));
            assertTrue("应包含运营备注面板", text.contains("Operational notes"));
        }
    }
}

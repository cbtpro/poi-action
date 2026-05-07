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

        assertTrue(output.length() > 0);
        try (PDDocument document = PDDocument.load(output)) {
            assertEquals(1, document.getNumberOfPages());
            String text = new PDFTextStripper().getText(document);
            assertTrue(text.contains("poi-action PDF Report"));
            assertTrue(text.contains("Apache PDFBox"));
            assertTrue(text.contains("Supported"));
        }
    }
}

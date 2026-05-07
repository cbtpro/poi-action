package com.chenbitao.word.pdf;

import com.chenbitao.word.util.ImageBytesUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.BeforeClass;
import org.junit.Test;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class PdfBoxDocumentGeneratorTest {

    @BeforeClass
    public static void configurePdfBoxLogging() {
        PdfBoxTestLoggingConfigurer.configure();
    }

    @Test
    public void saveWritesPdfWithTextTableAndImage() throws Exception {
        PdfBoxDocumentGenerator generator = new PdfBoxDocumentGenerator();
        generator.createDocument();
        generator.addTitle("Project Report");
        generator.addParagraph("PDF generation is supported by PDFBox.");
        generator.addTable(Arrays.asList(
                Arrays.asList("Type", "Status"),
                Arrays.asList("PDF", "Supported")
        ));
        generator.addImage(new ByteArrayInputStream(demoImage()), 120F, 60F);

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (PDDocument document = PDDocument.load(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals(1, document.getNumberOfPages());
            String text = new PDFTextStripper().getText(document);
            assertTrue(text.contains("Project Report"));
            assertTrue(text.contains("PDF generation is supported"));
            assertTrue(text.contains("Supported"));
        }
    }

    private byte[] demoImage() throws Exception {
        return ImageBytesUtils.solidPng(120, 60, Color.BLUE.getRGB());
    }
}

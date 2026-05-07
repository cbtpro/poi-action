package com.chenbitao.word.pdf;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.BeforeClass;
import org.junit.Test;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.image.BufferedImage;
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
        BufferedImage image = new BufferedImage(120, 60, BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < image.getWidth(); x++) {
            for (int y = 0; y < image.getHeight(); y++) {
                image.setRGB(x, y, Color.BLUE.getRGB());
            }
        }
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
    }
}

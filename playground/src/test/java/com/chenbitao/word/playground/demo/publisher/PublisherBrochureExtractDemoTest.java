package com.chenbitao.word.playground.demo.publisher;

import com.chenbitao.word.publisher.PublisherDocumentInfo;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class PublisherBrochureExtractDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesSamplePublisherAndSummary() throws Exception {
        Path outputDirectory = temporaryFolder.newFolder("publisher").toPath();

        PublisherDocumentInfo document = PublisherBrochureExtractDemo.generate(outputDirectory);

        Path sample = outputDirectory.resolve("brochure-demo.pub");
        Path summary = outputDirectory.resolve("brochure-summary.txt");
        assertTrue(Files.size(sample) > 0);
        assertTrue(Files.size(summary) > 0);
        assertEquals("Product Brochure", document.getTitle());
        assertEquals("Marketing", document.getSubject());
        assertEquals("poi-action", document.getAuthor());
        assertTrue(document.getText().contains("Publisher brochure extraction"));

        String summaryText = new String(Files.readAllBytes(summary), StandardCharsets.UTF_8);
        assertTrue(summaryText.contains("Title: Product Brochure"));
        assertTrue(summaryText.contains("Text: Publisher brochure extraction"));
    }
}

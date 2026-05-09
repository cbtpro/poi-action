package com.chenbitao.word.playground.demo.visio;

import com.chenbitao.word.visio.VisioDrawingInfo;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class VisioWorkflowExtractDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesSampleDrawingAndSummary() throws Exception {
        Path outputDirectory = temporaryFolder.newFolder("visio").toPath();

        VisioDrawingInfo drawing = VisioWorkflowExtractDemo.generate(outputDirectory);

        Path sample = outputDirectory.resolve("workflow-demo.vsdx");
        Path summary = outputDirectory.resolve("workflow-summary.txt");
        assertTrue(Files.size(sample) > 0);
        assertTrue(Files.size(summary) > 0);
        assertEquals(1, drawing.getPageCount());
        assertEquals(2, drawing.getShapeCount());

        String summaryText = new String(Files.readAllBytes(summary), StandardCharsets.UTF_8);
        assertTrue(summaryText.contains("Page: Workflow"));
        assertTrue(summaryText.contains("Shape: Start"));
        assertTrue(summaryText.contains("text: Approve"));
    }
}

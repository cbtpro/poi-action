package com.chenbitao.word.playground.demo.outlook;

import com.chenbitao.word.outlook.OutlookMessageInfo;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class OutlookMessageExtractDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesSampleMessageAndSummary() throws Exception {
        Path outputDirectory = temporaryFolder.newFolder("outlook").toPath();

        OutlookMessageInfo message = OutlookMessageExtractDemo.generate(outputDirectory);

        Path sample = outputDirectory.resolve("weekly-report-demo.msg");
        Path summary = outputDirectory.resolve("weekly-report-summary.txt");
        assertTrue(Files.size(sample) > 0);
        assertTrue(Files.size(summary) > 0);
        assertEquals("项目周报", message.getSubject());
        assertEquals(1, message.getAttachmentCount());
        String summaryText = new String(Files.readAllBytes(summary), StandardCharsets.UTF_8);
        assertTrue(summaryText.contains("主题: 项目周报"));
        assertTrue(summaryText.contains("附件: weekly-report.txt"));
    }
}

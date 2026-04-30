package com.chenbitao.word.docx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.io.FileInputStream;

import static org.junit.Assert.assertEquals;

public class DocxWordGeneratorTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void saveWritesReadableTitleParagraphAndTable() throws Exception {
        File output = temporaryFolder.newFile("生成结果.docx");
        DocxWordGenerator generator = new DocxWordGenerator();

        generator.createDocument();
        generator.addTitle("季度报告", 1);
        generator.addParagraph("收入增长。");
        generator.addTable(2, 2);
        generator.save(output.getAbsolutePath());

        try (XWPFDocument document = new XWPFDocument(new FileInputStream(output))) {
            assertEquals("季度报告", document.getParagraphArray(0).getText());
            assertEquals("收入增长。", document.getParagraphArray(1).getText());
            assertEquals(1, document.getTables().size());
            assertEquals(2, document.getTables().get(0).getRows().size());
            assertEquals(2, document.getTables().get(0).getRow(0).getTableCells().size());
            assertEquals("cell", document.getTables().get(0).getRow(1).getCell(1).getText());
        }
    }
}

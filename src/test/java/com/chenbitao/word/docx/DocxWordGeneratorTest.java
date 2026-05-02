package com.chenbitao.word.docx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;

import static org.junit.Assert.assertEquals;

/**
 * DocxWordGenerator DOCX生成器测试
 * 测试DocxWordGenerator的文档创建、标题添加、段落添加和表格添加功能
 */
public class DocxWordGeneratorTest {

    /** 临时文件夹规则，用于创建临时测试文件 */
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    /**
     * 测试保存操作写入可读的标题、段落和表格
     * 验证生成的DOCX文档包含正确的标题、段落和表格结构
     *
     * @throws Exception 如果测试过程中发生IO或其他错误
     */
    @Test
    public void saveWritesReadableTitleParagraphAndTable() throws Exception {
        File output = temporaryFolder.newFile("生成结果.docx");
        DocxWordGenerator generator = new DocxWordGenerator();

        generator.createDocument();
        generator.addTitle("季度报告", 1);
        generator.addParagraph("收入增长。");
        generator.addTable(2, 2);
        generator.save(output.getAbsolutePath());

        try (XWPFDocument document = new XWPFDocument(Files.newInputStream(output.toPath()))) {
            assertEquals("季度报告", document.getParagraphArray(0).getText());
            assertEquals("收入增长。", document.getParagraphArray(1).getText());
            assertEquals(1, document.getTables().size());
            assertEquals(2, document.getTables().get(0).getRows().size());
            assertEquals(2, document.getTables().get(0).getRow(0).getTableCells().size());
            assertEquals("cell", document.getTables().get(0).getRow(1).getCell(1).getText());
        }
    }
}

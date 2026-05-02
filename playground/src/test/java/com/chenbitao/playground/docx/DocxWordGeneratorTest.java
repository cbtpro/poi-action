package com.chenbitao.playground.docx;

import com.chenbitao.word.core.WordTable;
import com.chenbitao.word.core.WordTableCell;
import com.chenbitao.word.docx.DocxWordGenerator;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.math.BigInteger;
import java.nio.file.Files;
import java.util.Arrays;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

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
            assertTableFullWidth(document);
            assertEquals(2, document.getTables().get(0).getRows().size());
            assertEquals(2, document.getTables().get(0).getRow(0).getTableCells().size());
            assertEquals("cell", document.getTables().get(0).getRow(1).getCell(1).getText());
        }
    }

    /**
     * 测试结构化表格支持单元格图片和跨行。
     *
     * @throws Exception 如果测试过程中发生IO或其他错误
     */
    @Test
    public void structuredTableSupportsImageCellWithRowSpan() throws Exception {
        File output = temporaryFolder.newFile("结构化表格.docx");
        DocxWordGenerator generator = new DocxWordGenerator();

        generator.createDocument();
        generator.addTable(WordTable.of(Arrays.asList(
                Arrays.asList(
                        WordTableCell.text("姓名"),
                        WordTableCell.text("张三"),
                        WordTableCell.text("性别"),
                        WordTableCell.text("男"),
                        WordTableCell.image(new ByteArrayInputStream(pngBytes()),
                                Units.toEMU(20),
                                Units.toEMU(30)).rowSpan(3)
                ),
                Arrays.asList(
                        WordTableCell.text("出生年月"),
                        WordTableCell.text("2008-01"),
                        WordTableCell.text("民族"),
                        WordTableCell.text("-")
                ),
                Arrays.asList(
                        WordTableCell.text("籍贯"),
                        WordTableCell.text("-"),
                        WordTableCell.text("出生地"),
                        WordTableCell.text("-")
                )
        )));
        generator.save(output.getAbsolutePath());

        try (XWPFDocument document = new XWPFDocument(Files.newInputStream(output.toPath()))) {
            assertEquals(1, document.getTables().size());
            assertTableFullWidth(document);
            assertEquals(5, document.getTables().get(0).getRow(0).getTableCells().size());
            assertEquals("男", document.getTables().get(0).getRow(0).getCell(3).getText());
            assertTrue(document.getTables().get(0).getRow(0).getCell(4).getCTTc().getTcPr().isSetVMerge());
            assertTrue(document.getTables().get(0).getRow(1).getCell(4).getCTTc().getTcPr().isSetVMerge());
            assertTrue(document.getTables().get(0).getRow(2).getCell(4).getCTTc().getTcPr().isSetVMerge());
            assertEquals(1, document.getAllPictures().size());
        }
    }

    private byte[] pngBytes() throws Exception {
        BufferedImage image = new BufferedImage(10, 10, BufferedImage.TYPE_INT_RGB);
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
    }

    private void assertTableFullWidth(XWPFDocument document) {
        assertEquals(BigInteger.valueOf(5000), document.getTables().get(0).getCTTbl().getTblPr().getTblW().getW());
        assertEquals(STTblWidth.PCT, document.getTables().get(0).getCTTbl().getTblPr().getTblW().getType());
    }
}

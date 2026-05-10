package com.chenbitao.word.playground.demo.programmatic;

import com.chenbitao.word.playground.demo.commons.TestConstants;
import com.chenbitao.word.playground.demo.template.CadreTemplateDemoData;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.math.BigInteger;
import java.nio.file.Files;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class ProgrammaticCadreDocumentWriterTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void writeCreatesProgrammaticCadreDocument() throws Exception {
        File output = temporaryFolder.newFile("programmatic-demo.docx");

        new ProgrammaticCadreDocumentWriter().write(CadreTemplateDemoData.create(), output.toPath());

        try (XWPFDocument document = new XWPFDocument(Files.newInputStream(output.toPath()))) {
            assertEquals("干部任免审批表", document.getParagraphArray(0).getText());
            assertEquals(1, document.getTables().size());
            assertEquals(1, document.getAllPictures().size());

            XWPFTable table = document.getTables().get(0);
            assertEquals(BigInteger.valueOf(9752), table.getCTTbl().getTblPr().getTblW().getW());
            assertEquals(7, table.getCTTbl().getTblGrid().sizeOfGridColArray());
            assertEquals("姓名", table.getRow(0).getCell(0).getText());
            assertEquals(TestConstants.MOCK_NAME, table.getRow(0).getCell(1).getText());
            assertTrue(table.getRow(0).getCell(6).getCTTc().getTcPr().isSetVMerge());
            assertTrue(table.getRow(4).getCell(0).getCTTc().getTcPr().isSetVMerge());
        }
    }

    @Test
    public void writeCanGenerateMultipleDocuments() throws Exception {
        ProgrammaticCadreDocumentWriter writer = new ProgrammaticCadreDocumentWriter();

        for (int i = 0; i < 3; i++) {
            File output = temporaryFolder.newFile("programmatic-demo-" + i + ".docx");

            writer.write(CadreTemplateDemoData.create(), output.toPath());

            assertTrue(output.length() > 0);
            try (XWPFDocument document = new XWPFDocument(Files.newInputStream(output.toPath()))) {
                assertEquals("干部任免审批表", document.getParagraphArray(0).getText());
                assertEquals(1, document.getTables().size());
            }
        }
    }
}

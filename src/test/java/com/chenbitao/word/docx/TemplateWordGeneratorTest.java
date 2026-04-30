package com.chenbitao.word.docx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;

public class TemplateWordGeneratorTest {

    @Test
    public void renderSupportsCellLoopAndRowLoop() throws Exception {
        byte[] template = createTemplate();

        TemplateWordGenerator generator =
                new TemplateWordGenerator(new ByteArrayInputStream(template));
        Map<String, Object> data = new HashMap<>();
        data.put("name", "张三");
        data.put("exp", Arrays.asList(
                exp("2024", "后端开发"),
                exp("2025", "平台建设")
        ));

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.render(data);
        generator.save(output);

        try (XWPFDocument document =
                     new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("你好张三", document.getParagraphArray(0).getText());

            XWPFTable cellLoopTable = document.getTables().get(0);
            assertEquals(1, cellLoopTable.getRows().size());
            assertEquals("工作经历", cellLoopTable.getRow(0).getCell(0).getText());
            assertEquals("2024 后端开发\n2025 平台建设", cellLoopTable.getRow(0).getCell(1).getText());

            XWPFTable rowLoopTable = document.getTables().get(1);
            assertEquals(3, rowLoopTable.getRows().size());
            assertEquals("时间", rowLoopTable.getRow(0).getCell(0).getText());
            assertEquals("描述", rowLoopTable.getRow(0).getCell(1).getText());
            assertEquals("2024", rowLoopTable.getRow(1).getCell(0).getText());
            assertEquals("后端开发", rowLoopTable.getRow(1).getCell(1).getText());
            assertEquals("2025", rowLoopTable.getRow(2).getCell(0).getText());
            assertEquals("平台建设", rowLoopTable.getRow(2).getCell(1).getText());
        }
    }

    private static byte[] createTemplate() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText("你好${name}");

            XWPFTable cellLoopTable = document.createTable(1, 2);
            XWPFTableRow cellLoopRow = cellLoopTable.getRow(0);
            cellLoopRow.getCell(0).setText("工作经历");
            cellLoopRow.getCell(1).setText("${exp.time} ${exp.desc}");

            XWPFTable rowLoopTable = document.createTable(2, 2);
            rowLoopTable.getRow(0).getCell(0).setText("时间");
            rowLoopTable.getRow(0).getCell(1).setText("描述");
            XWPFTableRow rowLoopRow = rowLoopTable.getRow(1);
            rowLoopRow.getCell(0).setText("${exp.time}");
            rowLoopRow.getCell(1).setText("${exp.desc}");

            document.write(output);
            return output.toByteArray();
        }
    }

    private static Map<String, String> exp(String time, String desc) {
        Map<String, String> exp = new HashMap<>();
        exp.put("time", time);
        exp.put("desc", desc);
        return exp;
    }
}

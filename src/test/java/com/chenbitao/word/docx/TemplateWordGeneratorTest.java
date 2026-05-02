package com.chenbitao.word.docx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;

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

    @Test
    public void renderSupportsEducationArrayAsTableRows() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFTable table = document.createTable(1, 4);
            XWPFTableRow row = table.getRow(0);
            row.getCell(0).setText("学历学位");
            row.getCell(1).setText("${education.type}");
            row.getCell(2).setText("${education.degree}");
            row.getCell(3).setText("${education.department}${education.major}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("education", Arrays.asList(
                    education("全日制", "学士", "计算机科学", "软件工程"),
                    education("在职", "硕士", "软件学院", "软件工程")
            ));

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFTable renderedTable = rendered.getTables().get(0);
                assertEquals(2, renderedTable.getRows().size());
                assertEquals("学历学位", renderedTable.getRow(0).getCell(0).getText());
                assertEquals("全日制", renderedTable.getRow(0).getCell(1).getText());
                assertEquals("学士", renderedTable.getRow(0).getCell(2).getText());
                assertEquals("计算机科学软件工程", renderedTable.getRow(0).getCell(3).getText());
                assertEquals("", renderedTable.getRow(1).getCell(0).getText());
                assertEquals("在职", renderedTable.getRow(1).getCell(1).getText());
                assertEquals("硕士", renderedTable.getRow(1).getCell(2).getText());
                assertEquals("软件学院软件工程", renderedTable.getRow(1).getCell(3).getText());
            }
        }
    }

    @Test
    public void renderSupportsSplitRunAndNestedMapPlaceholders() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFParagraph splitParagraph = document.createParagraph();
            splitParagraph.createRun().setText("姓名：${na");
            splitParagraph.createRun().setText("meSc}");

            XWPFParagraph nestedParagraph = document.createParagraph();
            nestedParagraph.createRun().setText("学历：${education.t");
            nestedParagraph.createRun().setText("ype}/${education.degree}");

            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("nameSc", "李四");

            Map<String, Object> education = education("全日制", "硕士", "计算机科学", "软件工程");
            data.put("education", education);

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                assertEquals("姓名：李四", rendered.getParagraphArray(0).getText());
                assertEquals("学历：全日制/硕士", rendered.getParagraphArray(1).getText());
            }
        }
    }

    @Test
    public void renderKeepsLineBreakBetweenPlaceholders() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun().setText("${birthday}");
            paragraph.createRun().addBreak();
            paragraph.createRun().setText("${ag");
            paragraph.createRun().setText("e}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("birthday", "2008-01");
            data.put("age", "18");

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                assertEquals("2008-01\n18", rendered.getParagraphArray(0).getText());
            }
        }
    }

    @Test
    public void renderSupportsListAndArrayValuesAsMultipleLines() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText("${annualAssessmentResult}");
            document.createParagraph().createRun().setText("${tags}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("annualAssessmentResult", Arrays.asList(
                    "2023年度：2等（卓越)",
                    "2022年度：2等（卓越）",
                    "2021年度：3等（合格）"
            ));
            data.put("tags", new String[]{"A", "B"});

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                assertEquals("2023年度：2等（卓越)\n2022年度：2等（卓越）\n2021年度：3等（合格）",
                        rendered.getParagraphArray(0).getText());
                assertEquals("A\nB", rendered.getParagraphArray(1).getText());
            }
        }
    }

    @Test
    public void renderRealTemplateWithoutLeavingPlaceholders() throws Exception {
        InputStream template = TemplateWordGeneratorTest.class.getResourceAsStream("/template.docx");
        assertNotNull(template);

        TemplateWordGenerator generator = new TemplateWordGenerator(template);
        Map<String, Object> data = new HashMap<>();
        data.put("nameSc", "张三");
        data.put("sex", "男");
        data.put("age", "18");
        data.put("birthday", "2008-01");
        data.put("firstJobDate", "2021-07");
        data.put("qualificationName", "高级工程师");
        data.put("annualAssessmentResult", Arrays.asList(
                "2023年度：2等（卓越)",
                "2022年度：2等（卓越）",
                "2021年度：3等（合格）"
        ));
        data.put("workExperience", "2020.01 - 2021.06  ABC公司  高级开发工程师\n2021.07 - 2023.12  XYZ公司  技术总监");

        data.put("education", Arrays.asList(
                education("全日制", "学士", "计算机科学", "软件工程"),
                education("在职", "硕士", "软件学院", "软件工程")
        ));

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.render(data);
        generator.save(output);

        try (XWPFDocument rendered =
                     new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
            assertFalse(allText(rendered).contains("${"));
            assertFalse(allText(rendered).contains("2008-0118"));
            assertFalse(allText(rendered).contains("2008-01\n\n18"));
            assertFalse(allText(rendered).contains("2008-01\r\n\r\n18"));
            assertFalse(allText(rendered).contains("[2023年度"));
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

    private static Map<String, Object> exp(String time, String desc) {
        Map<String, Object> exp = new HashMap<>();
        exp.put("time", time);
        exp.put("desc", desc);
        return exp;
    }

    private static Map<String, Object> education(String type, String degree,
                                                 String department, String major) {
        Map<String, Object> education = new HashMap<>();
        education.put("type", type);
        education.put("degree", degree);
        education.put("department", department);
        education.put("major", major);
        return education;
    }

    private static String allText(XWPFDocument document) {
        StringBuilder text = new StringBuilder();
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            text.append(paragraph.getText()).append('\n');
        }
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (int i = 0; i < row.getTableCells().size(); i++) {
                    text.append(row.getCell(i).getText()).append('\n');
                }
            }
        }
        return text.toString();
    }
}

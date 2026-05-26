package com.chenbitao.word.docx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class TemplateWordGeneratorStyleTest {

    @Test
    public void renderKeepsRunStyleWhenReplacingTextPlaceholders() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun prefix = paragraph.createRun();
            prefix.setColor("FF0000");
            prefix.setText("姓名：");
            XWPFRun placeholderStart = paragraph.createRun();
            placeholderStart.setBold(true);
            placeholderStart.setText("${na");
            XWPFRun placeholderEnd = paragraph.createRun();
            placeholderEnd.setItalic(true);
            placeholderEnd.setText("me}");
            XWPFRun suffix = paragraph.createRun();
            suffix.setColor("00AA00");
            suffix.setText("（正式）");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("name", "张三");

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFParagraph renderedParagraph = rendered.getParagraphArray(0);
                assertEquals("姓名：张三（正式）", renderedParagraph.getText());
                assertEquals("姓名：", renderedParagraph.getRuns().get(0).getText(0));
                assertEquals("FF0000", renderedParagraph.getRuns().get(0).getColor());
                assertEquals("张三", renderedParagraph.getRuns().get(1).getText(0));
                assertTrue(renderedParagraph.getRuns().get(1).isBold());
                assertEquals("", renderedParagraph.getRuns().get(2).getText(0));
                assertTrue(renderedParagraph.getRuns().get(2).isItalic());
                assertEquals("（正式）", renderedParagraph.getRuns().get(3).getText(0));
                assertEquals("00AA00", renderedParagraph.getRuns().get(3).getColor());
            }
        }
    }

    @Test
    public void renderKeepsLineBreakRunBetweenPlaceholders() throws Exception {
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
                XWPFParagraph renderedParagraph = rendered.getParagraphArray(0);
                assertEquals("2008-01\n18", renderedParagraph.getText());
                assertEquals(1, renderedParagraph.getRuns().get(1).getCTR().sizeOfBrArray());
            }
        }
    }

    @Test
    public void renderKeepsCarriageReturnBetweenPlaceholders() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun().setText("${birthday}");
            paragraph.createRun().addCarriageReturn();
            paragraph.createRun().setText("${age}");
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
                XWPFParagraph renderedParagraph = rendered.getParagraphArray(0);
                assertEquals("2008-01\n18", renderedParagraph.getText());
                assertEquals(1, renderedParagraph.getRuns().get(1).getCTR().sizeOfBrArray());
            }
        }
    }

    @Test
    public void renderUsesStyledRunInsideSplitPlaceholder() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun().setText("${na");
            XWPFRun styledRun = paragraph.createRun();
            styledRun.setFontFamily("宋体");
            styledRun.setFontSize(12);
            styledRun.setText("me}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("name", "张三");

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFRun renderedRun = rendered.getParagraphArray(0).getRuns().get(1);
                assertEquals("张三", renderedRun.getText(0));
                assertEquals("宋体", renderedRun.getFontFamily());
                assertEquals(12, renderedRun.getFontSize());
            }
        }
    }

    @Test
    public void renderKeepsEastAsiaFontAndXiaoSiSize() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setFontFamily("宋体");
            run.setFontSize(12);
            CTFonts fonts = run.getCTR().getRPr().getRFonts();
            fonts.setEastAsia("宋体");
            fonts.setHAnsi("宋体");
            fonts.setAscii("宋体");
            run.setText("${name}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("name", "张三");

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFRun renderedRun = rendered.getParagraphArray(0).getRuns().get(0);
                assertEquals("张三", renderedRun.getText(0));
                assertEquals("宋体", renderedRun.getCTR().getRPr().getRFonts().getEastAsia());
                assertEquals("宋体", renderedRun.getCTR().getRPr().getRFonts().getHAnsi());
                assertEquals("宋体", renderedRun.getCTR().getRPr().getRFonts().getAscii());
                assertEquals(12, renderedRun.getFontSize());
            }
        }
    }

    @Test
    public void renderUsesRunWithExplicitNonFontStyleInsideSplitPlaceholder() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun().setText("${na");
            XWPFRun styledRun = paragraph.createRun();
            styledRun.getCTR().addNewRPr().addNewU().setVal(STUnderline.SINGLE);
            styledRun.getCTR().getRPr().addNewHighlight().setVal(STHighlightColor.YELLOW);
            styledRun.setText("me}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("name", "张三");

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFRun renderedRun = rendered.getParagraphArray(0).getRuns().get(1);
                assertEquals("张三", renderedRun.getText(0));
                assertEquals(STUnderline.SINGLE, renderedRun.getCTR().getRPr().getU().getVal());
                assertEquals(STHighlightColor.YELLOW, renderedRun.getCTR().getRPr().getHighlight().getVal());
            }
        }
    }

    @Test
    public void renderLoopCellsKeepTemplateFontAndLineBreaks() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFTable table = document.createTable(1, 2);
            table.getRow(0).getCell(0).setText("经历");
            setStyledCellText(table.getRow(0).getCell(1), "${exp.time}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> first = new HashMap<>();
            first.put("time", "2024");
            Map<String, Object> second = new HashMap<>();
            second.put("time", "2025");
            Map<String, Object> data = new HashMap<>();
            data.put("exp", java.util.Arrays.asList(first, second));

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFRun renderedRun = rendered.getTables().get(0).getRow(0).getCell(1).getParagraphs().get(0).getRuns().get(0);
                assertEquals("2024\n2025", rendered.getTables().get(0).getRow(0).getCell(1).getText());
                assertEquals("宋体", renderedRun.getCTR().getRPr().getRFonts().getEastAsia());
                assertEquals(12, renderedRun.getFontSize());
            }
        }
    }

    @Test
    public void renderLoopRowsKeepTemplateFont() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFTable table = document.createTable(1, 2);
            setStyledCellText(table.getRow(0).getCell(0), "${exp.time}");
            setStyledCellText(table.getRow(0).getCell(1), "${exp.desc}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> first = new HashMap<>();
            first.put("time", "2024");
            first.put("desc", "后端开发");
            Map<String, Object> second = new HashMap<>();
            second.put("time", "2025");
            second.put("desc", "平台建设");
            Map<String, Object> data = new HashMap<>();
            data.put("exp", java.util.Arrays.asList(first, second));

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFRun firstRun = rendered.getTables().get(0).getRow(0).getCell(0).getParagraphs().get(0).getRuns().get(0);
                XWPFRun secondRun = rendered.getTables().get(0).getRow(1).getCell(0).getParagraphs().get(0).getRuns().get(0);
                assertEquals("2024", firstRun.getText(0));
                assertEquals("2025", secondRun.getText(0));
                assertEquals("宋体", firstRun.getCTR().getRPr().getRFonts().getEastAsia());
                assertEquals("宋体", secondRun.getCTR().getRPr().getRFonts().getEastAsia());
                assertEquals(12, firstRun.getFontSize());
                assertEquals(12, secondRun.getFontSize());
            }
        }
    }

    @Test
    public void renderLoopRowsKeepLineBreakBetweenPlaceholdersInSameCell() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFTable table = document.createTable(1, 1);
            XWPFParagraph paragraph = table.getRow(0).getCell(0).getParagraphs().get(0);
            for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                paragraph.removeRun(i);
            }
            XWPFRun firstRun = styledRun(paragraph);
            firstRun.setText("${education.department}");
            XWPFRun breakRun = styledRun(paragraph);
            breakRun.addBreak();
            XWPFRun secondRun = styledRun(paragraph);
            secondRun.setText("${education.major}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> education = new HashMap<>();
            education.put("department", "计算机科学");
            education.put("major", "软件工程");
            Map<String, Object> data = new HashMap<>();
            data.put("education", java.util.Collections.singletonList(education));

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFRun renderedRun = rendered.getTables().get(0).getRow(0).getCell(0).getParagraphs().get(0).getRuns().get(0);
                assertEquals("计算机科学\n软件工程", rendered.getTables().get(0).getRow(0).getCell(0).getText());
                assertEquals("宋体", renderedRun.getCTR().getRPr().getRFonts().getEastAsia());
                assertEquals(12, renderedRun.getFontSize());
            }
        }
    }

    private void setStyledCellText(org.apache.poi.xwpf.usermodel.XWPFTableCell cell, String text) {
        XWPFParagraph paragraph = cell.getParagraphs().get(0);
        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
        XWPFRun run = paragraph.createRun();
        applySongXiaoSi(run);
        run.setText(text);
    }

    private XWPFRun styledRun(XWPFParagraph paragraph) {
        XWPFRun run = paragraph.createRun();
        applySongXiaoSi(run);
        return run;
    }

    private void applySongXiaoSi(XWPFRun run) {
        run.setFontFamily("宋体");
        run.setFontSize(12);
        CTFonts fonts = run.getCTR().getRPr().getRFonts();
        fonts.setEastAsia("宋体");
        fonts.setHAnsi("宋体");
        fonts.setAscii("宋体");
    }
}

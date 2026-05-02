package com.chenbitao.playground.docx;

import com.chenbitao.word.constant.BlankConstants;
import com.chenbitao.word.docx.TemplateWordGenerator;
import com.chenbitao.word.exception.WordException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.util.Base64;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

/**
 * TemplateWordGenerator模板生成器测试
 * 全面测试TemplateWordGenerator的功能，包括：
 * - 文本占位符替换
 * - 表格循环填充（单元格内循环和行循环）
 * - 图片占位符替换
 * - 嵌套对象属性访问
 * - 列表和数组值处理
 * - 垂直合并单元格保持
 */
public class TemplateWordGeneratorTest {

    /**
     * 测试渲染支持单元格循环和行循环
     * 验证模板生成器能够正确处理${list.field}格式的循环占位符，
     * 包括在单元格内的循环填充和通过新增行进行循环填充
     *
     * @throws Exception 如果测试过程中发生错误
     */
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
    public void renderSupportsMultipleLoopRowsInSameTable() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFTable table = document.createTable(2, 5);
            XWPFTableRow educationRow = table.getRow(0);
            educationRow.getCell(0).setText("学历学位");
            educationRow.getCell(1).setText("${education.type}");
            educationRow.getCell(2).setText("${education.degree}");
            educationRow.getCell(3).setText("${education.department}${education.major}");
            educationRow.getCell(4).setText("备注");

            XWPFTableRow relationRow = table.getRow(1);
            relationRow.getCell(0).setText("${familyAndSocialRelations.appellation}");
            relationRow.getCell(1).setText("${familyAndSocialRelations.name}");
            relationRow.getCell(2).setText("${familyAndSocialRelations.age}");
            relationRow.getCell(3).setText("${familyAndSocialRelations.political}");
            relationRow.getCell(4).setText("${familyAndSocialRelations.workUnit}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("education", Arrays.asList(
                    education("全日制", "学士", "计算机科学", "软件工程"),
                    education("在职", "硕士", "软件学院", "软件工程")
            ));
            data.put("familyAndSocialRelations", Arrays.asList(
                    relation("父亲", "张父", "56", "党员", BlankConstants.EMPTY),
                    relation("母亲", "张母", "55", "群众", BlankConstants.DASH)
            ));

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFTable renderedTable = rendered.getTables().get(0);
                assertEquals(4, renderedTable.getRows().size());
                assertEquals("学历学位", renderedTable.getRow(0).getCell(0).getText());
                assertEquals("全日制", renderedTable.getRow(0).getCell(1).getText());
                assertEquals("", renderedTable.getRow(1).getCell(0).getText());
                assertEquals("在职", renderedTable.getRow(1).getCell(1).getText());
                assertEquals("父亲", renderedTable.getRow(2).getCell(0).getText());
                assertEquals("张父", renderedTable.getRow(2).getCell(1).getText());
                assertEquals("56", renderedTable.getRow(2).getCell(2).getText());
                assertEquals("党员", renderedTable.getRow(2).getCell(3).getText());
                assertEquals(BlankConstants.EMPTY, renderedTable.getRow(2).getCell(4).getText());
                assertEquals("母亲", renderedTable.getRow(3).getCell(0).getText());
                assertEquals("张母", renderedTable.getRow(3).getCell(1).getText());
                assertEquals("55", renderedTable.getRow(3).getCell(2).getText());
                assertEquals("群众", renderedTable.getRow(3).getCell(3).getText());
                assertEquals(BlankConstants.DASH, renderedTable.getRow(3).getCell(4).getText());
            }
        }
    }

    @Test
    public void renderKeepsTemplateVerticalMergeForLoopRows() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            XWPFTable table = document.createTable(2, 6);
            XWPFTableRow headerRow = table.getRow(0);
            setVerticalMerge(headerRow.getCell(0), STMerge.RESTART);
            headerRow.getCell(0).setText("家庭主要成员及重要社会关系");
            headerRow.getCell(1).setText("称谓");
            headerRow.getCell(2).setText("姓名");
            headerRow.getCell(3).setText("年龄");
            headerRow.getCell(4).setText("政治面貌");
            headerRow.getCell(5).setText("工作单位及职务");

            XWPFTableRow relationRow = table.getRow(1);
            setVerticalMerge(relationRow.getCell(0), STMerge.CONTINUE);
            relationRow.getCell(1).setText("${familyAndSocialRelations.appellation}");
            relationRow.getCell(2).setText("${familyAndSocialRelations.name}");
            relationRow.getCell(3).setText("${familyAndSocialRelations.age}");
            relationRow.getCell(4).setText("${familyAndSocialRelations.political}");
            relationRow.getCell(5).setText("${familyAndSocialRelations.workUnit}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("familyAndSocialRelations", Arrays.asList(
                    relation("父亲", "张父", "56", "党员", BlankConstants.EMPTY),
                    relation("母亲", "张母", "55", "群众", BlankConstants.DASH)
            ));

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                XWPFTable renderedTable = rendered.getTables().get(0);
                assertEquals(3, renderedTable.getRows().size());
                assertEquals(STMerge.RESTART, verticalMerge(renderedTable.getRow(0).getCell(0)));
                assertEquals(STMerge.CONTINUE, verticalMerge(renderedTable.getRow(1).getCell(0)));
                assertEquals(STMerge.CONTINUE, verticalMerge(renderedTable.getRow(2).getCell(0)));
                assertEquals("父亲", renderedTable.getRow(1).getCell(1).getText());
                assertEquals("母亲", renderedTable.getRow(2).getCell(1).getText());
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
    public void renderSupportsPicturePlaceholderFromMultipleSources() throws Exception {
        byte[] bytesImage = pngBytes(0x3366CC);
        byte[] base64Image = pngBytes(0xCC6633);
        byte[] fileImage = pngBytes(0x33CC66);
        byte[] urlImage = pngBytes(0x6633CC);
        File imageFile = File.createTempFile("template-photo", ".png");
        java.nio.file.Files.write(imageFile.toPath(), fileImage);
        imageFile.deleteOnExit();
        File urlFile = File.createTempFile("template-photo-url", ".png");
        java.nio.file.Files.write(urlFile.toPath(), urlImage);
        urlFile.deleteOnExit();

        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText("${photoBytes}");
            document.createParagraph().createRun().setText("${photoBase64}");
            document.createParagraph().createRun().setText("${photoFile}");
            document.createParagraph().createRun().setText("${photoUrl}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("photoBytes", bytesImage);
            data.put("photoBase64", "data:image/png;base64," + Base64.getEncoder().encodeToString(base64Image));
            data.put("photoFile", imageFile.toPath());
            data.put("photoUrl", urlFile.toURI().toURL());

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                assertEquals(4, rendered.getAllPictures().size());
                assertFalse(allText(rendered).contains("${photo"));
            }
        }
    }

    @Test
    public void renderConvertsImageIoReadablePictureToWordPicture() throws Exception {
        byte[] imageBytes = bmpBytes();
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText("${photo}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("photo", TemplateWordGenerator.picture(
                    Base64.getEncoder().encodeToString(imageBytes),
                    org.apache.poi.util.Units.toEMU(20),
                    org.apache.poi.util.Units.toEMU(30)));

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                assertEquals(1, rendered.getAllPictures().size());
                assertTrue(rendered.getAllPictures().get(0).getData().length > 0);
                String xml = rendered.getDocument().xmlText();
                assertTrue(xml.contains("cx=\"" + org.apache.poi.util.Units.toEMU(20) + "\""));
                assertTrue(xml.contains("cy=\"" + org.apache.poi.util.Units.toEMU(30) + "\""));
            }
        }
    }

    @Test
    public void renderClearsPicturePlaceholderWhenValueIsBlank() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream template = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText("${photoNull}");
            document.createParagraph().createRun().setText("${photoEmpty}");
            document.write(template);

            TemplateWordGenerator generator =
                    new TemplateWordGenerator(new ByteArrayInputStream(template.toByteArray()));
            Map<String, Object> data = new HashMap<>();
            data.put("photoNull", null);
            data.put("photoEmpty", "");

            ByteArrayOutputStream output = new ByteArrayOutputStream();
            generator.render(data);
            generator.save(output);

            try (XWPFDocument rendered =
                         new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
                assertEquals("", rendered.getParagraphArray(0).getText());
                assertEquals("", rendered.getParagraphArray(1).getText());
                assertEquals(0, rendered.getAllPictures().size());
                assertFalse(allText(rendered).contains("${photo"));
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
        data.put("nationality", "汉族");
        data.put("nativePlace", "浙江杭州");
        data.put("birthPlace", "浙江杭州");
        data.put("healthCondition", "健康");
        data.put("joinPartyDate", BlankConstants.DASH);
        data.put("majorExpertise", BlankConstants.NONE);
        data.put("qualificationName", "高级工程师");
        data.put("currentPosition", BlankConstants.DASH);
        data.put("proposedPosition", BlankConstants.DASH);
        data.put("proposedRemovedPosition", BlankConstants.DASH);
        data.put("annualAssessmentResult", Arrays.asList(
                "2023年度：2等（卓越)",
                "2022年度：2等（卓越）",
                "2021年度：3等（合格）"
        ));
        data.put("appointmentRemovalReason", BlankConstants.NONE);
        data.put("rewardPunishmentRecord", BlankConstants.NONE);
        data.put("reportingUnit", BlankConstants.DASH);
        data.put("workExperience", "2020.01 - 2021.06  ABC公司  高级开发工程师\n2021.07 - 2023.12  XYZ公司  技术总监");
        data.put("familyAndSocialRelations", Arrays.asList(
                relation("父亲", "张父", "56", "群众", BlankConstants.DASH),
                relation("母亲", "张母", "55", "群众", BlankConstants.DASH)
        ));

        data.put("education", Arrays.asList(
                education("全日制", "学士", "计算机科学", "软件工程"),
                education("在职", "硕士", "软件学院", "软件工程")
        ));
        data.put("photo", TemplateWordGenerator.picture(
                pngBytes(),
                org.apache.poi.util.Units.toEMU(25),
                org.apache.poi.util.Units.toEMU(35)));

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
            assertEquals(1, rendered.getAllPictures().size());
        }
    }

    @Test
    public void renderRealTemplateClearsPhotoWhenValueIsNull() throws Exception {
        InputStream template = TemplateWordGeneratorTest.class.getResourceAsStream("/template.docx");
        assertNotNull(template);

        TemplateWordGenerator generator = new TemplateWordGenerator(template);
        Map<String, Object> data = new HashMap<>();
        data.put("photo", null);

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.render(data);
        generator.save(output);

        try (XWPFDocument rendered =
                     new XWPFDocument(new ByteArrayInputStream(output.toByteArray()))) {
            assertFalse(allText(rendered).contains("${photo}"));
            assertEquals(0, rendered.getAllPictures().size());
        }
    }

    @Test(expected = WordException.class)
    public void renderRejectsLoopListItemThatIsNotMap() throws Exception {
        TemplateWordGenerator generator = new TemplateWordGenerator(new ByteArrayInputStream(createTemplate()));
        Map<String, Object> data = new HashMap<>();
        data.put("name", "张三");
        data.put("exp", Arrays.asList("not-a-map"));

        generator.render(data);
    }

    @Test(expected = WordException.class)
    public void renderRejectsLoopMapWithNonStringKey() throws Exception {
        TemplateWordGenerator generator = new TemplateWordGenerator(new ByteArrayInputStream(createTemplate()));
        Map<Object, Object> item = new HashMap<>();
        item.put(1, "2024");

        Map<String, Object> data = new HashMap<>();
        data.put("name", "张三");
        data.put("exp", Arrays.asList(item));

        generator.render(data);
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

    private static Map<String, Object> relation(String appellation, String name,
                                                String age, String political,
                                                String workUnit) {
        Map<String, Object> relation = new HashMap<>();
        relation.put("appellation", appellation);
        relation.put("name", name);
        relation.put("age", age);
        relation.put("political", political);
        relation.put("workUnit", workUnit);
        return relation;
    }

    private static void setVerticalMerge(XWPFTableCell cell, STMerge.Enum value) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr()
                ? cell.getCTTc().getTcPr()
                : cell.getCTTc().addNewTcPr();
        CTVMerge vMerge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
        vMerge.setVal(value);
    }

    private static STMerge.Enum verticalMerge(XWPFTableCell cell) {
        if (!cell.getCTTc().isSetTcPr() || !cell.getCTTc().getTcPr().isSetVMerge()) {
            return null;
        }
        return cell.getCTTc().getTcPr().getVMerge().getVal();
    }

    private static byte[] pngBytes() throws Exception {
        return pngBytes(0x3366CC);
    }

    private static byte[] pngBytes(int rgb) throws Exception {
        return imageBytes("png", rgb);
    }

    private static byte[] bmpBytes() throws Exception {
        return imageBytes("bmp", 0x3366CC);
    }

    private static byte[] imageBytes(String formatName, int rgb) throws Exception {
        BufferedImage image = new BufferedImage(10, 12, BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < image.getWidth(); x++) {
            for (int y = 0; y < image.getHeight(); y++) {
                image.setRGB(x, y, rgb);
            }
        }
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, formatName, output);
        return output.toByteArray();
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

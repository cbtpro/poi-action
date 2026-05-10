package com.chenbitao.word.presentation;

import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.hslf.usermodel.HSLFTextShape;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * PowerPoint 模板生成器测试。
 */
public class TemplatePresentationGeneratorTest {

    /**
     * 验证 PPT 模板可以渲染文本框、表格单元格、嵌套字段和列表字段。
     *
     * @throws Exception 如果模板读取或结果解析失败
     */
    @Test
    public void renderHslfTemplateReplacesTextAndTablePlaceholders() throws Exception {
        TemplateHslfPresentationGenerator generator =
                new TemplateHslfPresentationGenerator(templateStream("/template-project.ppt"));
        generator.render(templateData());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (HSLFSlideShow slideShow = new HSLFSlideShow(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals(3, slideShow.getSlides().size());
            assertTrue(slideContainsText(slideShow.getSlides().get(0), "季度经营复盘"));
            assertTrue(slideContainsText(slideShow.getSlides().get(0), "区域：华东"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "新签客户"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "续约提升"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "${missing}"));
            assertTrue(tableContainsText(slideShow.getSlides().get(2), "1280"));
            assertTrue(tableContainsText(slideShow.getSlides().get(2), "94.6%"));
        }
    }

    /**
     * 验证 PPTX 模板可以渲染文本框、表格单元格、嵌套字段和列表字段。
     *
     * @throws Exception 如果模板读取或结果解析失败
     */
    @Test
    public void renderXslfTemplateReplacesTextAndTablePlaceholders() throws Exception {
        TemplateXslfPresentationGenerator generator =
                new TemplateXslfPresentationGenerator(templateStream("/template-project.pptx"));
        generator.render(templateData());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (XMLSlideShow slideShow = new XMLSlideShow(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals(3, slideShow.getSlides().size());
            assertTrue(slideContainsText(slideShow.getSlides().get(0), "季度经营复盘"));
            assertTrue(slideContainsText(slideShow.getSlides().get(0), "区域：华东"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "新签客户"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "续约提升"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "${missing}"));
            assertTrue(tableContainsText(slideShow.getSlides().get(2), "1280"));
            assertTrue(tableContainsText(slideShow.getSlides().get(2), "94.6%"));
        }
    }

    /**
     * 加载测试模板文件。
     *
     * @param resourceName 模板资源名称
     * @return 模板输入流
     */
    private InputStream templateStream(String resourceName) {
        InputStream template = TemplatePresentationGeneratorTest.class.getResourceAsStream(resourceName);
        if (template == null) {
            throw new IllegalStateException("未找到测试模板文件：" + resourceName);
        }
        return template;
    }

    /**
     * 构造模板数据。
     *
     * @return 模板数据
     */
    private Map<String, Object> templateData() {
        Map<String, Object> data = new HashMap<>();

        Map<String, Object> report = new HashMap<>();
        report.put("title", "季度经营复盘");
        report.put("region", "华东");
        data.put("report", report);

        List<Map<String, Object>> highlights = new ArrayList<>();
        highlights.add(highlight("新签客户"));
        highlights.add(highlight("续约提升"));
        data.put("highlights", highlights);

        Map<String, Object> metrics = new HashMap<>();
        metrics.put("revenue", "1280");
        metrics.put("renewalRate", "94.6%");
        data.put("metrics", metrics);
        return data;
    }

    /**
     * 构造单条亮点数据。
     *
     * @param name 亮点名称
     * @return 亮点数据
     */
    private Map<String, Object> highlight(String name) {
        Map<String, Object> highlight = new HashMap<>();
        highlight.put("name", name);
        return highlight;
    }

    /**
     * 判断 PPT 幻灯片是否包含指定文本。
     *
     * @param slide 幻灯片
     * @param expected 期望文本
     * @return 如果包含期望文本则返回 true
     */
    private boolean slideContainsText(HSLFSlide slide, String expected) {
        for (HSLFShape shape : slide.getShapes()) {
            if (shape instanceof HSLFTextShape) {
                String text = ((HSLFTextShape) shape).getText();
                if (text != null && text.contains(expected)) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断 PPT 表格是否包含指定文本。
     *
     * @param slide 幻灯片
     * @param expected 期望文本
     * @return 如果表格包含期望文本则返回 true
     */
    private boolean tableContainsText(HSLFSlide slide, String expected) {
        for (HSLFShape shape : slide.getShapes()) {
            if (shape instanceof HSLFTable) {
                HSLFTable table = (HSLFTable) shape;
                for (int row = 0; row < table.getNumberOfRows(); row++) {
                    for (int column = 0; column < table.getNumberOfColumns(); column++) {
                        String text = table.getCell(row, column).getText();
                        if (text != null && text.contains(expected)) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    /**
     * 判断 PPTX 幻灯片是否包含指定文本。
     *
     * @param slide 幻灯片
     * @param expected 期望文本
     * @return 如果包含期望文本则返回 true
     */
    private boolean slideContainsText(XSLFSlide slide, String expected) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
                String text = ((XSLFTextShape) shape).getText();
                if (text != null && text.contains(expected)) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断 PPTX 表格是否包含指定文本。
     *
     * @param slide 幻灯片
     * @param expected 期望文本
     * @return 如果表格包含期望文本则返回 true
     */
    private boolean tableContainsText(XSLFSlide slide, String expected) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTable) {
                XSLFTable table = (XSLFTable) shape;
                for (int row = 0; row < table.getNumberOfRows(); row++) {
                    for (int column = 0; column < table.getNumberOfColumns(); column++) {
                        String text = table.getCell(row, column).getText();
                        if (text != null && text.contains(expected)) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }
}

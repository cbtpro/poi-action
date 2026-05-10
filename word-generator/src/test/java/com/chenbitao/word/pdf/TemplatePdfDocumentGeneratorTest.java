package com.chenbitao.word.pdf;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * PDF 模板文档生成器测试。
 */
public class TemplatePdfDocumentGeneratorTest {

    /**
     * 配置 PDFBox 测试日志。
     */
    @BeforeClass
    public static void configurePdfBoxLogging() {
        PdfBoxTestLoggingConfigurer.configure();
    }

    /**
     * 验证 PDF 模板文件可以按表单域名称填充并扁平化输出。
     *
     * @throws Exception 如果模板读取或结果解析失败
     */
    @Test
    public void renderFillsAcroFormTemplateFields() throws Exception {
        TemplatePdfDocumentGenerator generator =
                new TemplatePdfDocumentGenerator(templateStream("/template-report.pdf"));
        generator.render(templateData());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (PDDocument document = PDDocument.load(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals(1, document.getNumberOfPages());
            String text = new PDFTextStripper().getText(document);
            assertTrue(text.contains("Q2 Portfolio Review"));
            assertTrue(text.contains("North Region"));
            assertTrue(text.contains("Galaxy Manufacturing"));
            assertTrue(text.contains("Renewal rate improved"));
            assertTrue(text.contains("$12.8M"));
            assertTrue(text.contains("94.6%"));
            assertTrue(text.contains("91"));
        }
    }

    /**
     * 加载测试模板文件。
     *
     * @param resourceName 模板资源名称
     * @return 模板输入流
     */
    private InputStream templateStream(String resourceName) {
        InputStream template = TemplatePdfDocumentGeneratorTest.class.getResourceAsStream(resourceName);
        if (template == null) {
            throw new IllegalStateException("未找到测试模板文件：" + resourceName);
        }
        return template;
    }

    /**
     * 构造 PDF 模板数据。
     *
     * @return 模板数据
     */
    private Map<String, Object> templateData() {
        Map<String, Object> data = new HashMap<>();
        data.put("reportTitle", "Q2 Portfolio Review");
        data.put("reportRegion", "North Region");
        data.put("customerName", "Galaxy Manufacturing");
        data.put("highlights", Arrays.asList(
                "Renewal rate improved",
                "Delivery risks reduced",
                "Template rollout accelerated"
        ));
        data.put("revenue", "$12.8M");
        data.put("renewalRate", "94.6%");
        data.put("deliveryHealth", "91");
        data.put("reportNote", "Generated from a real PDF template file with AcroForm fields.");
        return data;
    }
}

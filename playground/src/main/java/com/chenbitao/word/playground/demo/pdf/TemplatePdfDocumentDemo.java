package com.chenbitao.word.playground.demo.pdf;

import com.chenbitao.word.pdf.TemplatePdfDocumentGenerator;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * PDF 模板渲染演示。
 */
@Slf4j
public class TemplatePdfDocumentDemo {

    private static final String TEMPLATE = "/template-report.pdf";
    private static final String OUTPUT_FILE_NAME = "template-report-demo.pdf";

    /**
     * PDF 模板渲染演示入口。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 如果模板读取、目录创建或文件写出失败
     */
    public static void main(String[] args) throws Exception {
        PdfBoxLoggingConfigurer.configure();
        Path outputPath = defaultOutputPath();
        generate(outputPath);
        log.info("PDF 模板演示文件生成完成：{}", outputPath.toAbsolutePath());
    }

    /**
     * 基于 PDF 模板文件生成 PDF。
     *
     * @param outputPath 输出路径
     * @throws Exception 如果目录创建或文件写出失败
     */
    public static void generate(Path outputPath) throws Exception {
        Files.createDirectories(outputPath.getParent());

        TemplatePdfDocumentGenerator generator =
                new TemplatePdfDocumentGenerator(templateStream(TEMPLATE));
        generator.render(data());
        generator.save(outputPath.toString());
    }

    /**
     * 加载 classpath 模板文件。
     *
     * @param templateName 模板资源名称
     * @return 模板输入流
     */
    private static InputStream templateStream(String templateName) {
        InputStream template = TemplatePdfDocumentDemo.class.getResourceAsStream(templateName);
        if (template == null) {
            throw new IllegalStateException("未找到 PDF 模板文件：" + templateName);
        }
        return template;
    }

    /**
     * 构造 PDF 模板数据。
     *
     * @return 模板数据
     */
    private static Map<String, Object> data() {
        Map<String, Object> data = new HashMap<>();
        data.put("reportTitle", "Q2 Portfolio Review");
        data.put("reportRegion", "East Region");
        data.put("customerName", "Galaxy Manufacturing");
        data.put("highlights", Arrays.asList(
                "18 new high-intent opportunities qualified",
                "Renewal rate improved to 94.6%",
                "Template rollout reduced average delivery time by 32%"
        ));
        data.put("revenue", "$12.8M");
        data.put("renewalRate", "94.6%");
        data.put("deliveryHealth", "91");
        data.put("reportNote", "This file was generated from playground/src/main/resources/template-report.pdf.");
        return data;
    }

    /**
     * 获取默认输出路径。
     *
     * @return 默认输出路径
     */
    private static Path defaultOutputPath() {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", OUTPUT_FILE_NAME);
        }
        return Paths.get("playground", "target", OUTPUT_FILE_NAME);
    }
}

package com.chenbitao.word.playground.demo.presentation;

import com.chenbitao.word.presentation.TemplateHslfPresentationGenerator;
import com.chenbitao.word.presentation.TemplateXslfPresentationGenerator;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * PowerPoint 模板渲染演示。
 */
@Slf4j
public class TemplatePresentationDocumentDemo {

    private static final String PPT_TEMPLATE = "/template-project.ppt";
    private static final String PPTX_TEMPLATE = "/template-project.pptx";
    private static final String PPT_OUTPUT_FILE_NAME = "template-project-demo.ppt";
    private static final String PPTX_OUTPUT_FILE_NAME = "template-project-demo.pptx";

    /**
     * PowerPoint 模板渲染演示入口。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 如果模板读取、目录创建或文件写出失败
     */
    public static void main(String[] args) throws Exception {
        generate(defaultPptOutputPath(), defaultPptxOutputPath());
        log.info("PowerPoint 模板演示文件生成完成");
    }

    /**
     * 生成 PPT 和 PPTX 模板渲染文件。
     *
     * @param pptOutputPath PPT 输出路径
     * @param pptxOutputPath PPTX 输出路径
     * @throws Exception 如果模板读取、目录创建或文件写出失败
     */
    public static void generate(Path pptOutputPath, Path pptxOutputPath) throws Exception {
        Files.createDirectories(pptOutputPath.getParent());
        Files.createDirectories(pptxOutputPath.getParent());

        Map<String, Object> templateData = data();
        renderPpt(pptOutputPath, templateData);
        renderPptx(pptxOutputPath, templateData);

        log.info("PPT 模板文件已保存到：{}", pptOutputPath.toAbsolutePath());
        log.info("PPTX 模板文件已保存到：{}", pptxOutputPath.toAbsolutePath());
    }

    /**
     * 渲染 PPT 模板文件。
     *
     * @param outputPath 输出路径
     * @param templateData 模板数据
     */
    private static void renderPpt(Path outputPath, Map<String, Object> templateData) {
        TemplateHslfPresentationGenerator generator =
                new TemplateHslfPresentationGenerator(templateStream(PPT_TEMPLATE));
        generator.render(templateData);
        generator.save(outputPath.toString());
    }

    /**
     * 渲染 PPTX 模板文件。
     *
     * @param outputPath 输出路径
     * @param templateData 模板数据
     */
    private static void renderPptx(Path outputPath, Map<String, Object> templateData) {
        TemplateXslfPresentationGenerator generator =
                new TemplateXslfPresentationGenerator(templateStream(PPTX_TEMPLATE));
        generator.render(templateData);
        generator.save(outputPath.toString());
    }

    /**
     * 加载 classpath 模板文件。
     *
     * @param templateName 模板资源名称
     * @return 模板输入流
     */
    private static InputStream templateStream(String templateName) {
        InputStream template = TemplatePresentationDocumentDemo.class.getResourceAsStream(templateName);
        if (template == null) {
            throw new IllegalStateException("未找到 PowerPoint 模板文件：" + templateName);
        }
        return template;
    }

    /**
     * 构造模板数据。
     *
     * @return 模板数据
     */
    private static Map<String, Object> data() {
        Map<String, Object> data = new HashMap<>();

        Map<String, Object> report = new HashMap<>();
        report.put("title", "2026 Q2 项目经营复盘");
        report.put("region", "华东大区");
        report.put("note", "重点客户续约稳定，新增线索正在进入方案确认阶段。");
        data.put("report", report);

        Map<String, Object> customer = new HashMap<>();
        customer.put("name", "星河制造集团");
        data.put("customer", customer);

        data.put("highlights", highlights());
        data.put("metrics", metrics());
        return data;
    }

    /**
     * 构造亮点列表。
     *
     * @return 亮点列表
     */
    private static List<Map<String, Object>> highlights() {
        List<Map<String, Object>> highlights = new ArrayList<>();
        highlights.add(highlight("新增 18 家高意向客户，制造业客户占比 61%"));
        highlights.add(highlight("核心客户续约率提升至 94.6%，高于目标 2.1 个百分点"));
        highlights.add(highlight("模板化交付流程缩短平均制作周期 32%"));
        return highlights;
    }

    /**
     * 构造单条亮点。
     *
     * @param name 亮点名称
     * @return 亮点数据
     */
    private static Map<String, Object> highlight(String name) {
        Map<String, Object> highlight = new HashMap<>();
        highlight.put("name", name);
        return highlight;
    }

    /**
     * 构造指标数据。
     *
     * @return 指标数据
     */
    private static Map<String, Object> metrics() {
        Map<String, Object> metrics = new HashMap<>();
        metrics.put("revenue", "1280 万");
        metrics.put("revenueStatus", "达成 106%");
        metrics.put("renewalRate", "94.6%");
        metrics.put("renewalStatus", "健康");
        metrics.put("deliveryHealth", "91 分");
        metrics.put("deliveryStatus", "可控");
        return metrics;
    }

    /**
     * 获取默认 PPT 输出路径。
     *
     * @return 默认 PPT 输出路径
     */
    private static Path defaultPptOutputPath() {
        return defaultOutputPath(PPT_OUTPUT_FILE_NAME);
    }

    /**
     * 获取默认 PPTX 输出路径。
     *
     * @return 默认 PPTX 输出路径
     */
    private static Path defaultPptxOutputPath() {
        return defaultOutputPath(PPTX_OUTPUT_FILE_NAME);
    }

    /**
     * 获取默认输出路径。
     *
     * @param fileName 文件名
     * @return 默认输出路径
     */
    private static Path defaultOutputPath(String fileName) {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", fileName);
        }
        return Paths.get("playground", "target", fileName);
    }
}

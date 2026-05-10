package com.chenbitao.word.playground.demo.template;

import com.chenbitao.word.docx.TemplateWordGenerator;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@Slf4j
public class TemplateDocumentDemo {

    /**
     * 模板渲染演示入口，默认读取 classpath 下的 {@code /template.docx}。
     */
    public static void main(String[] args) throws Exception {
        TemplateWordGenerator generator = new TemplateWordGenerator(loadTemplate());
        generator.render(CadreTemplateDemoData.create());
        saveDocument(generator);

        log.info("文档生成完成");
    }

    /**
     * 加载模板文件
     * @param templateName 模板路径
     * @return 模板输入流
     */
    public static InputStream loadTemplate(String templateName) {
        // 如果传入null或空，使用默认模板
        String template = (templateName == null || templateName.isEmpty())
                ? "/template.docx"
                : templateName;

        // 从classpath下加载资源
        InputStream inputStream = TemplateDocumentDemo.class.getResourceAsStream(template);

        // 为空则抛出明确异常
        if (inputStream == null) {
            throw new IllegalStateException("未找到模板文件：" + template);
        }

        return inputStream;
    }

    /**
     * 重载方法：无参时使用默认模板
     */
    public static InputStream loadTemplate() {
        return loadTemplate("/template.docx");
    }

    private static void saveDocument(TemplateWordGenerator generator) throws Exception {
        Path output = Paths.get("target", "template-demo.docx");
        Files.createDirectories(output.getParent());
        generator.save(output.toString());
        log.info("文档已保存到：{}", output.toAbsolutePath());
    }
}

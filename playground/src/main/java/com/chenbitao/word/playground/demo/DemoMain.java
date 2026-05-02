package com.chenbitao.word.playground.demo;

import lombok.extern.slf4j.Slf4j;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;

/**
 * 编程式Word文档生成演示主类。
 * 使用和TemplateDemo相同的业务数据，通过代码直接构建与template.docx一致的版式。
 */
@Slf4j
public class DemoMain {

    private static final Path OUTPUT_PATH = Paths.get("target", "programmatic-demo.docx");

    /**
     * 程序入口方法。
     *
     * @param args 命令行参数（未使用）
     * @throws Exception 如果业务数据构造或文档生成失败
     */
    public static void main(String[] args) throws Exception {
        Map<String, Object> data = TemplateDemoData.create();
        Files.createDirectories(OUTPUT_PATH.getParent());

        ProgrammaticCadreDocumentWriter writer = new ProgrammaticCadreDocumentWriter();
        writer.write(data, OUTPUT_PATH);

        log.info("编程式Word生成成功：{}", OUTPUT_PATH.toAbsolutePath());
    }
}

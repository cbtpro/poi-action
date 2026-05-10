package com.chenbitao.word.playground.demo.programmatic;

import com.chenbitao.word.playground.demo.template.CadreTemplateDemoData;
import lombok.extern.slf4j.Slf4j;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;

/**
 * 编程式干部任免审批表生成演示。
 *
 * <p>该入口复用模板演示的数据，但不依赖 Word 模板文件，便于对比“模板渲染”和“代码构建”
 * 两种生成方式在同一业务场景下的差异。</p>
 */
@Slf4j
public class ProgrammaticCadreDocumentDemo {

    private static final Path OUTPUT_PATH = Paths.get("target", "programmatic-demo.docx");

    /**
     * 程序入口方法。
     *
     * @param args 命令行参数（未使用）
     * @throws Exception 如果业务数据构造或文档生成失败
     */
    public static void main(String[] args) throws Exception {
        Map<String, Object> data = CadreTemplateDemoData.create();
        Files.createDirectories(OUTPUT_PATH.getParent());

        ProgrammaticCadreDocumentWriter writer = new ProgrammaticCadreDocumentWriter();
        writer.write(data, OUTPUT_PATH);

        log.info("编程式Word生成成功：{}", OUTPUT_PATH.toAbsolutePath());
    }
}

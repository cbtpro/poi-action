package com.chenbitao.word.playground.demo;

import com.chenbitao.word.builder.WordBuilder;
import com.chenbitao.word.factory.WordGeneratorFactory;
import lombok.extern.slf4j.Slf4j;

import java.util.Arrays;

/**
 * Word文档生成演示主类
 * 演示如何使用WordBuilder和WordGeneratorFactory来生成简历文档
 */
@Slf4j
public class DemoMain {

    /**
     * 程序入口方法
     * 演示完整的简历文档生成流程：
     * 1. 构造简历数据
     * 2. 创建Word文档建造者
     * 3. 使用流式API构建文档内容
     * 4. 保存生成的文档
     *
     * @param args 命令行参数（未使用）
     */
    public static void main(String[] args) {
        Resume resume = Resume.builder().build();
        resume.name = "张三";
        resume.phone = "13800000000";
        resume.email = "zhangsan@example.com";

        resume.education = Arrays.asList(
                "2018-2022 本科 - 计算机科学与技术",
                "主修课程：数据结构、操作系统"
        );

        resume.experience = Arrays.asList(
                "2022-至今 某互联网公司 Java开发工程师",
                "负责系统架构设计与开发"
        );

        resume.skills = Arrays.asList(
                "Java / Spring Boot",
                "MySQL / Redis",
                "Docker / Linux"
        );

        // 2️⃣ 创建文档
        WordBuilder builder = new WordBuilder(
                WordGeneratorFactory.get("docx")
        );

        // 3️⃣ 开始构建简历
        builder
                .title(resume.name)
                .paragraph("电话：" + resume.phone + "    邮箱：" + resume.email)

                .title("教育经历")
                .paragraphList(resume.education)

                .title("工作经历")
                .paragraphList(resume.experience)

                .title("技能")
                .paragraphList(resume.skills)

                .table(3, 2) // 简单表格（后面可升级）
                .build(System.getProperty("user.dir") + "/target/test.docx");

        log.info("生成成功！");
    }
}
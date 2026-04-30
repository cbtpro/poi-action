package com.chenbitao.word.demo;

import com.chenbitao.word.builder.WordBuilder;
import com.chenbitao.word.factory.WordGeneratorFactory;
import lombok.extern.slf4j.Slf4j;

import java.util.Arrays;

@Slf4j
public class DemoMain {

    public static void main(String[] args) {
// 1️⃣ 构造数据
        Resume resume = new Resume();
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
                .build("D:/test.docx");

        log.info("生成成功！");
    }
}
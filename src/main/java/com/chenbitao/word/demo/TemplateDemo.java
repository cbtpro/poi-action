package com.chenbitao.word.demo;

import com.chenbitao.word.docx.TemplateWordGenerator;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

@Slf4j
public class TemplateDemo {

    public static void main(String[] args) throws Exception {
        // 加载模板
        InputStream template = loadTemplate();

        // 创建生成器
        TemplateWordGenerator generator = new TemplateWordGenerator(template);

        // 准备数据
        Map<String, Object> data = prepareTemplateData();

        // 渲染文档
        generator.render(data);

        // 保存结果
        saveDocument(generator);

        log.info("文档生成完成");
    }

    private static InputStream loadTemplate() {
        InputStream template = TemplateDemo.class.getResourceAsStream("/template.docx");
        if (template == null) {
            throw new IllegalStateException("未找到模板文件：/template.docx");
        }
        return template;
    }

    private static Map<String, Object> prepareTemplateData() {
        Map<String, Object> data = new HashMap<>();

        // 基本信息
        addBasicInfo(data);

        // 专业资格
        addQualification(data);

        // 年度考核结果
        addAnnualAssessment(data);

        // 学历学位信息
        addEducation(data);

        // 工作经历
        addWorkExperience(data);

        return data;
    }

    private static void addBasicInfo(Map<String, Object> data) {
        data.put("nameSc", "张三");
        data.put("sex", "男");
        data.put("birthday", "2008-01");
        data.put("age", "18");
        data.put("nationality", "-");
        data.put("firstJobDate", "2021-07");
    }

    private static void addQualification(Map<String, Object> data) {
        data.put("qualificationName", "高级工程师");
    }

    private static void addAnnualAssessment(Map<String, Object> data) {
        data.put("annualAssessmentResult", Arrays.asList(
                "2023年度：2等（卓越)",
                "2022年度：2等（卓越）",
                "2021年度：3等（合格）"
        ));
    }

    private static void addEducation(Map<String, Object> data) {
        data.put("education", Arrays.asList(
                education("全日制", "学士", "计算机科学", "软件工程"),
                education("在职", "硕士", "软件学院", "软件工程")
        ));
    }

    private static void addWorkExperience(Map<String, Object> data) {
        String workExperience = "2020.01 - 2021.06  ABC公司  高级开发工程师\n" +
                "2021.07 - 2023.12  XYZ公司  技术总监\n" +
                "2024.01 - 2026.05  XX亚洲  技术架构师";
        data.put("workExperience", workExperience);
    }

    private static void saveDocument(TemplateWordGenerator generator) throws Exception {
        Path output = Paths.get("target", "template-demo.docx");
        Files.createDirectories(output.getParent());
        generator.save(output.toString());
        log.info("文档已保存到：{}", output.toAbsolutePath());
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
}

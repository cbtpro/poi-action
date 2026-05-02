package com.chenbitao.word.demo;

import com.chenbitao.word.constant.BlankConstants;
import com.chenbitao.word.docx.TemplateWordGenerator;
import lombok.extern.slf4j.Slf4j;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

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

    private static Map<String, Object> prepareTemplateData() throws Exception {
        Map<String, Object> data = new HashMap<>();

        // 基本信息
        addBasicInfo(data);

        // 专业资格
        addQualification(data);

        // 任免信息
        addAppointmentInfo(data);

        // 年度考核结果
        addAnnualAssessment(data);

        // 学历学位信息
        addEducation(data);

        // 工作经历
        addWorkExperience(data);

        // 家庭及社会关系
        addFamilyAndSocialRelations(data);

        // 照片
        addPhoto(data);

        return data;
    }

    private static void addBasicInfo(Map<String, Object> data) {
        data.put("nameSc", "张三");
        data.put("sex", "男");
        data.put("birthday", "2008-01");
        data.put("age", "18");
        data.put("nationality", BlankConstants.DASH);
        data.put("nativePlace", BlankConstants.DASH);
        data.put("birthPlace", BlankConstants.DASH);
        data.put("healthCondition", BlankConstants.DASH);
        data.put("joinPartyDate", BlankConstants.DASH);
        data.put("majorExpertise", BlankConstants.DASH);
        data.put("firstJobDate", "2021-07");
    }

    private static void addQualification(Map<String, Object> data) {
        data.put("qualificationName", "高级工程师\n技术总监");
    }

    private static void addAppointmentInfo(Map<String, Object> data) {
        data.put("currentPosition", BlankConstants.DASH);
        data.put("proposedPosition", BlankConstants.DASH);
        data.put("proposedRemovedPosition", BlankConstants.DASH);
        data.put("appointmentRemovalReason", BlankConstants.NONE);
        data.put("rewardPunishmentRecord", BlankConstants.NONE);
        data.put("reportingUnit", BlankConstants.EMPTY);
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
                buildEducation("全日制教育", "大学本科工商管理学士学位", "香港暨南大学", "工商管理学院"),
                buildEducation("在职教育", "硕士研究生理学博士", "香港中文大学", "市场营销")
        ));
    }

    private static void addWorkExperience(Map<String, Object> data) {
//        String workExperience = "2020.01 - 2021.06  ABC公司  高级开发工程师\n" +
//                "2021.07 - 2023.12  XYZ公司  技术总监\n" +
//                "2024.01 - 2026.05  XX亚洲  技术架构师";
//        data.put("workExperience", workExperience);
        data.put("workExperience", Arrays.asList(
                "2020.01 - 2021.06  ABC公司  高级开发工程师",
                "2021.07 - 2023.12  XYZ公司  技术总监",
                "2024.01 - 2026.05  XX亚洲  技术架构师"
        ));

    }

    private static void addFamilyAndSocialRelations(Map<String, Object> data) {
//        data.put("familyAndSocialRelations", Arrays.asList(
//                buildRelation("父亲", "张父", "56", "党员", BlankConstants.EMPTY),
//                buildRelation("母亲", "张母", "55", "群众", BlankConstants.DASH)
//        ));
        List<Map<String, String>> relationList = new ArrayList<>();
        // 生成7个空关系
        for (int i = 0; i < 7; i++) {
            relationList.add(buildRelation("", "", "", "", BlankConstants.EMPTY));
        }
        data.put("familyAndSocialRelations", relationList);
    }

    private static void addPhoto(Map<String, Object> data) throws Exception {
        data.put("photo", TemplateWordGenerator.picture(
                demoPhoto(),
                org.apache.poi.util.Units.toEMU(80),
                org.apache.poi.util.Units.toEMU(100)));
    }

    private static byte[] demoPhoto() throws Exception {
        BufferedImage image = new BufferedImage(600, 800, BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < image.getWidth(); x++) {
            for (int y = 0; y < image.getHeight(); y++) {
                image.setRGB(x, y, 0x4F7DD1);
            }
        }
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
    }

    private static void saveDocument(TemplateWordGenerator generator) throws Exception {
        Path output = Paths.get("target", "template-demo.docx");
        Files.createDirectories(output.getParent());
        generator.save(output.toString());
        log.info("文档已保存到：{}", output.toAbsolutePath());
    }

    private static Map<String, String> buildEducation(String type, String degree,
                                                 String department, String major) {
        Map<String, String> education = new HashMap<>();
        education.put("type", type);
        education.put("degree", degree);
        education.put("department", department);
        education.put("major", major);
        return education;
    }

    private static Map<String, String> buildRelation(String appellation, String name,
                                                     String age, String political,
                                                     String workUnit) {
        Map<String, String> relation = new HashMap<>();
        relation.put("appellation", appellation);
        relation.put("name", name);
        relation.put("age", age);
        relation.put("political", political);
        relation.put("workUnit", workUnit);
        return relation;
    }
}

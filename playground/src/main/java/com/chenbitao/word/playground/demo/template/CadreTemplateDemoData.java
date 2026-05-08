package com.chenbitao.word.playground.demo.template;

import com.chenbitao.word.constant.BlankConstants;
import com.chenbitao.word.docx.TemplateWordGenerator;
import com.chenbitao.word.playground.demo.commons.TestConstants;
import com.chenbitao.word.util.ImageBytesUtils;
import org.apache.poi.util.Units;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 干部任免审批表模板演示数据。
 *
 * <p>字段名需要与 {@code template.docx} 中的占位符保持一致，编程式演示也复用这份数据，
 * 以便两种生成方式输出同一业务内容。</p>
 */
public final class CadreTemplateDemoData {

    private CadreTemplateDemoData() {
    }

    /**
     * 创建一份完整的演示数据。
     *
     * @return 可直接传给 {@code TemplateWordGenerator.render(...)} 的数据映射
     */
    public static Map<String, Object> create() throws IOException {
        Map<String, Object> data = new HashMap<>();

        addBasicInfo(data);
        addQualification(data);
        addAppointmentInfo(data);
        addAnnualAssessment(data);
        addEducation(data);
        addWorkExperience(data);
        addFamilyAndSocialRelations(data);
        addPhoto(data);

        return data;
    }

    private static void addBasicInfo(Map<String, Object> data) {
        data.put("nameSc", TestConstants.MOCK_NAME);
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
        data.put("qualificationName", "高级工程师\\n技术总监");
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
        data.put("workExperience", Arrays.asList(
                "2020.01 - 2021.06  ABC公司  高级开发工程师",
                "2021.07 - 2023.12  XYZ公司  技术总监",
                "2024.01 - 2026.05  XX亚洲  技术架构师"
        ));
    }

    private static void addFamilyAndSocialRelations(Map<String, Object> data) {
        List<Map<String, String>> relationList = new ArrayList<>();
        // 模板表格预留 7 行关系信息，空数据用于验证空白占位符的渲染效果。
        for (int i = 0; i < 7; i++) {
            relationList.add(buildRelation("", "", "", "", BlankConstants.EMPTY));
        }
        data.put("familyAndSocialRelations", relationList);
    }

    private static void addPhoto(Map<String, Object> data) throws IOException {
        data.put("photo", TemplateWordGenerator.picture(
                demoPhoto(),
                Units.toEMU(80),
                Units.toEMU(100)));
    }

    private static byte[] demoPhoto() throws IOException {
        // 使用纯色图片替代真实证件照，避免演示代码携带个人图片资产。
        return ImageBytesUtils.solidPng(600, 800, 0x4F7DD1);
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

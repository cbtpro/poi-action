package com.chenbitao.playground.doc;

import com.chenbitao.word.constant.BlankConstants;
import com.chenbitao.word.doc.TemplateDocWordGenerator;
import com.chenbitao.word.docx.TemplateWordGenerator;
import org.apache.poi.hwpf.HWPFDocument;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

/**
 * DOC 模板文档生成器测试。
 */
public class TemplateDocWordGeneratorTest {

    /**
     * 测试真实 DOC 模板可以完成文本、嵌套字段、列表字段和图片占位符渲染。
     *
     * @throws Exception 如果模板读取或文档解析失败
     */
    @Test
    public void renderDocTemplateWithoutLeavingKnownPlaceholders() throws Exception {
        InputStream template = TemplateDocWordGeneratorTest.class.getResourceAsStream("/template.doc");
        assertNotNull(template);

        TemplateDocWordGenerator generator = new TemplateDocWordGenerator(template);
        generator.render(templateData());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        String text = docText(output.toByteArray());
        assertFalse(text.contains("${nameSc}"));
        assertFalse(text.contains("${education.type}"));
        assertFalse(text.contains("${familyAndSocialRelations.name}"));
        assertFalse(text.contains("${photo}"));
        assertTrue(text.contains("张三"));
        assertTrue(text.contains("2008-01"));
        assertTrue(text.contains("全日制"));
        assertTrue(text.contains("在职"));
        assertTrue(text.contains("张父"));
        assertTrue(text.contains("张母"));
        assertTrue(text.contains("[图片]"));
    }

    /**
     * 测试未提供数据的占位符会保留原样，便于调用方发现缺失字段。
     *
     * @throws Exception 如果模板读取或文档解析失败
     */
    @Test
    public void renderKeepsUnknownPlaceholders() throws Exception {
        InputStream template = TemplateDocWordGeneratorTest.class.getResourceAsStream("/template.doc");
        assertNotNull(template);

        TemplateDocWordGenerator generator = new TemplateDocWordGenerator(template);
        generator.render(new HashMap<String, Object>());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        String text = docText(output.toByteArray());
        assertTrue(text.contains("${nameSc}"));
        assertTrue(text.contains("${education.type}"));
    }

    /**
     * 构造 DOC 模板测试数据。
     *
     * @return 模板数据
     */
    private static Map<String, Object> templateData() {
        Map<String, Object> data = new HashMap<>();
        data.put("nameSc", "张三");
        data.put("sex", "男");
        data.put("age", "18");
        data.put("birthday", "2008-01");
        data.put("firstJobDate", "2021-07");
        data.put("nationality", "汉族");
        data.put("nativePlace", "杭州");
        data.put("birthPlace", "杭州");
        data.put("healthCondition", "健康");
        data.put("joinPartyDate", BlankConstants.DASH);
        data.put("majorExpertise", BlankConstants.NONE);
        data.put("qualificationName", "工程师");
        data.put("currentPosition", BlankConstants.DASH);
        data.put("proposedPosition", BlankConstants.DASH);
        data.put("proposedRemovedPosition", BlankConstants.DASH);
        data.put("annualAssessmentResult", Arrays.asList("优", "良"));
        data.put("appointmentRemovalReason", BlankConstants.NONE);
        data.put("rewardPunishmentRecord", BlankConstants.NONE);
        data.put("reportingUnit", BlankConstants.DASH);
        data.put("workExperience", new String[]{"ABC", "XYZ"});
        data.put("education", Arrays.asList(
                education("全日制", "学士", "计算机科学", "软件工程"),
                education("在职", "硕士", "软件学院", "软件工程")
        ));
        data.put("familyAndSocialRelations", Arrays.asList(
                relation("父亲", "张父", "56", "群众", BlankConstants.DASH),
                relation("母亲", "张母", "55", "群众", BlankConstants.DASH)
        ));
        data.put("photo", TemplateWordGenerator.picture(new byte[]{1, 2, 3}));
        return data;
    }

    /**
     * 构造学历学位数据。
     *
     * @param type 学历类型
     * @param degree 学位
     * @param department 院系
     * @param major 专业
     * @return 学历学位数据
     */
    private static Map<String, Object> education(String type, String degree,
                                                 String department, String major) {
        Map<String, Object> education = new HashMap<>();
        education.put("type", type);
        education.put("degree", degree);
        education.put("department", department);
        education.put("major", major);
        return education;
    }

    /**
     * 构造家庭成员和社会关系数据。
     *
     * @param appellation 称谓
     * @param name 姓名
     * @param age 年龄
     * @param political 政治面貌
     * @param workUnit 工作单位及职务
     * @return 关系数据
     */
    private static Map<String, Object> relation(String appellation, String name,
                                                String age, String political,
                                                String workUnit) {
        Map<String, Object> relation = new HashMap<>();
        relation.put("appellation", appellation);
        relation.put("name", name);
        relation.put("age", age);
        relation.put("political", political);
        relation.put("workUnit", workUnit);
        return relation;
    }

    /**
     * 读取 DOC 文档的全部文本。
     *
     * @param bytes DOC 文档字节
     * @return 文档文本
     * @throws Exception 如果文档解析失败
     */
    private static String docText(byte[] bytes) throws Exception {
        try (HWPFDocument document = new HWPFDocument(new ByteArrayInputStream(bytes))) {
            return document.getRange().text();
        }
    }
}

package com.chenbitao.word.playground.demo.template;

import com.chenbitao.word.doc.TemplateDocWordGenerator;
import com.chenbitao.word.docx.TemplateWordGenerator;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * DOC 模板渲染演示。
 */
@Slf4j
public class TemplateDocDocumentDemo {

    /**
     * DOC 模板渲染演示入口。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 如果模板读取或文件写出失败
     */
    public static void main(String[] args) throws Exception {
        TemplateDocWordGenerator generator = new TemplateDocWordGenerator(loadTemplate());
        generator.render(demoData());
        saveDocument(generator);
        log.info("DOC 模板文档生成完成");
    }

    /**
     * 加载默认 DOC 模板。
     *
     * @return 模板输入流
     */
    public static InputStream loadTemplate() {
        return loadTemplate("/template.doc");
    }

    /**
     * 加载指定 DOC 模板。
     *
     * @param templateName 模板资源路径
     * @return 模板输入流
     */
    public static InputStream loadTemplate(String templateName) {
        String template = (templateName == null || templateName.isEmpty())
                ? "/template.doc"
                : templateName;
        InputStream inputStream = TemplateDocDocumentDemo.class.getResourceAsStream(template);
        if (inputStream == null) {
            throw new IllegalStateException("未找到 DOC 模板文件：" + template);
        }
        return inputStream;
    }

    /**
     * 保存渲染后的 DOC 文档。
     *
     * @param generator DOC 模板生成器
     * @throws Exception 如果目录创建或文件写出失败
     */
    private static void saveDocument(TemplateDocWordGenerator generator) throws Exception {
        Path output = Paths.get("target", "template-demo.doc");
        Files.createDirectories(output.getParent());
        generator.save(output.toString());
        log.info("DOC 文档已保存到：{}", output.toAbsolutePath());
    }

    /**
     * 构造 DOC 模板演示数据。
     *
     * @return 模板数据
     */
    private static Map<String, Object> demoData() {
        Map<String, Object> data = new HashMap<>();
        data.put("nameSc", "张三");
        data.put("sex", "男");
        data.put("birthday", "2008-01");
        data.put("age", "18");
        data.put("firstJobDate", "2021-07");
        data.put("nationality", "汉族");
        data.put("nativePlace", "杭州");
        data.put("birthPlace", "杭州");
        data.put("joinPartyDate", "-");
        data.put("healthCondition", "健康");
        data.put("qualificationName", "工程师");
        data.put("majorExpertise", "Java");
        data.put("currentPosition", "-");
        data.put("proposedPosition", "-");
        data.put("proposedRemovedPosition", "-");
        data.put("workExperience", Arrays.asList("ABC", "XYZ"));
        data.put("rewardPunishmentRecord", "无");
        data.put("annualAssessmentResult", Arrays.asList("优", "良"));
        data.put("appointmentRemovalReason", "无");
        data.put("reportingUnit", "-");
        data.put("education", Arrays.asList(
                mapOf("type", "全日制", "degree", "学士", "department", "计算机", "major", "软件"),
                mapOf("type", "在职", "degree", "硕士", "department", "软件", "major", "工程")
        ));
        data.put("familyAndSocialRelations", Arrays.asList(
                mapOf("appellation", "父亲", "name", "张父", "age", "56", "political", "群众", "workUnit", "-"),
                mapOf("appellation", "母亲", "name", "张母", "age", "55", "political", "群众", "workUnit", "-")
        ));
        data.put("photo", TemplateWordGenerator.picture(new byte[]{1, 2, 3}));
        return data;
    }

    /**
     * 构造简短的键值映射。
     *
     * @param values 交替排列的键和值
     * @return 映射数据
     */
    private static Map<String, Object> mapOf(String... values) {
        Map<String, Object> data = new HashMap<>();
        for (int i = 0; i + 1 < values.length; i += 2) {
            data.put(values[i], values[i + 1]);
        }
        return data;
    }
}

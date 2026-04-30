package com.chenbitao.word.demo;

import com.chenbitao.word.docx.TemplateWordGenerator;
import lombok.extern.slf4j.Slf4j;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class TemplateDemo {

    public static void main(String[] args) throws Exception {

        TemplateWordGenerator generator =
                new TemplateWordGenerator(
                        TemplateDemo.class.getResourceAsStream("/template.docx")
                );

        Map<String, Object> data = new HashMap<>();
        data.put("name", "张三");
        data.put("sex", "男");
        data.put("age", "18");
        data.put("birthday", "1999-01-01");
        data.put("phone", "13800000000");
        data.put("email", "zhangsan@test.com");

        data.put("education", "本科 - 计算机科学");
        data.put("skills", "Java / Spring / MySQL");

        List<Map<String, String>> expList = new ArrayList<>();

        Map<String, String> e1 = new HashMap<>();
        e1.put("time", "2022-2024");
        e1.put("desc", "Java开发");

        Map<String, String> e2 = new HashMap<>();
        e2.put("time", "2025-2026");
        e2.put("desc", "高级Java开发");

        expList.add(e1);
        expList.add(e2);

        data.put("experience", expList);

        generator.render(data);

        generator.save(Files.newOutputStream(Paths.get("D:/resume.docx")));
        log.debug("生成成功");
    }
}
package com.chenbitao.word.playground.demo.web;

import com.chenbitao.word.docx.TemplateWordGenerator;
import com.chenbitao.word.playground.demo.template.CadreTemplateDemoData;
import com.chenbitao.word.playground.demo.template.TemplateDocumentDemo;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

@RestController
@RequestMapping("/resume")
public class ResumeController {

    /**
     * 下载模板渲染后的演示文档。
     */
    @GetMapping("/download")
    public void download(HttpServletResponse response) throws IOException {
        TemplateWordGenerator generator = new TemplateWordGenerator(TemplateDocumentDemo.loadTemplate());
        generator.render(CadreTemplateDemoData.create());

        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.setHeader("Content-Disposition", contentDisposition("template-demo.docx"));
        generator.save(response.getOutputStream());
    }

    private String contentDisposition(String fileName) throws UnsupportedEncodingException {
        String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name()).replace("+", "%20");
        // 同时提供 filename 和 RFC 5987 filename*，兼容常见浏览器的中文文件名下载。
        return "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + encodedFileName;
    }
}

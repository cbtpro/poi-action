package com.chenbitao.word.demo;

import com.chenbitao.word.docx.TemplateWordGenerator;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

@RestController
@RequestMapping("/resume")
public class ResumeController {

    @GetMapping("/download")
    public void download(HttpServletResponse response) throws Exception {
        TemplateWordGenerator generator = new TemplateWordGenerator(TemplateDemo.loadTemplate());
        generator.render(TemplateDemoData.create());

        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.setHeader("Content-Disposition", contentDisposition("template-demo.docx"));
        generator.save(response.getOutputStream());
    }

    private String contentDisposition(String fileName) throws UnsupportedEncodingException {
        String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name()).replace("+", "%20");
        return "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + encodedFileName;
    }
}

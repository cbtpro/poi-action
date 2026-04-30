//package com.chenbitao.word.demo;
//
//import com.chenbitao.word.docx.TemplateWordGenerator;
//
//import java.util.HashMap;
//import java.util.Map;
//
//@RestController
//@RequestMapping("/resume")
//public class ResumeController {
//
//    @GetMapping("/download")
//    public void download(HttpServletResponse response) throws Exception {
//
//        TemplateWordGenerator generator =
//                new TemplateWordGenerator(
//                        getClass().getResourceAsStream("/template.docx")
//                );
//
//        Map<String, Object> data = new HashMap<>();
//        data.put("name", "张三");
//        data.put("phone", "13800000000");
//
//        generator.render(data);
//
//        response.setContentType(
//                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
//        );
//        response.setHeader(
//                "Content-Disposition",
//                "attachment;filename=resume.docx"
//        );
//
//        generator.save(response.getOutputStream());
//    }
//}
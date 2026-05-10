package com.chenbitao.word.playground.demo.presentation;

import com.chenbitao.word.factory.PresentationGeneratorFactory;
import com.chenbitao.word.presentation.PresentationGenerator;
import com.chenbitao.word.util.ImageBytesUtils;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;

/**
 * PowerPoint 97-2003 项目汇报演示。
 *
 * <p>该演示生成一份包含封面、文本要点、表格和图片的 {@code .ppt} 文件。</p>
 */
@Slf4j
public class PptProjectReportDemo {

    private static final String OUTPUT_FILE_NAME = "project-report-demo.ppt";

    /**
     * 生成 PPT 演示文件。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 如果目录创建、图片生成或文件写出失败
     */
    public static void main(String[] args) throws Exception {
        Path outputPath = defaultOutputPath();

        generate(outputPath);

        log.info("PPT 演示文件生成成功：{}", outputPath.toAbsolutePath());
    }

    /**
     * 生成项目汇报 PPT。
     *
     * @param outputPath 输出路径
     * @throws Exception 如果目录创建、图片生成或文件写出失败
     */
    public static void generate(Path outputPath) throws Exception {
        Files.createDirectories(outputPath.getParent());

        PresentationGenerator generator = PresentationGeneratorFactory.get("ppt");
        generator.createPresentation();
        generator.addTitleSlide("poi-action 项目汇报", "PowerPoint 97-2003 生成演示");
        generator.addTextSlide("新增能力", Arrays.asList(
                "支持 HSLF 生成 .ppt 演示文稿",
                "支持标题页、文本页、表格页和图片页",
                "与 Word、Excel 生成能力保持独立工厂入口"
        ));
        generator.addTableSlide("格式支持进度", Arrays.asList(
                Arrays.asList("类型", "格式", "状态"),
                Arrays.asList("Word", ".doc / .docx", "已支持"),
                Arrays.asList("Excel", ".xls / .xlsx", "已支持"),
                Arrays.asList("PowerPoint", ".ppt", "已支持")
        ));
        generator.addImageSlide("演示图片", new ByteArrayInputStream(demoImage()));
        generator.save(outputPath.toString());
    }

    private static byte[] demoImage() throws Exception {
        return ImageBytesUtils.png(480, 300, (x, y, width, height) -> {
            int blue = 160 + (x * 80 / width);
            int green = 110 + (y * 90 / height);
            return (80 << 16) | (green << 8) | blue;
        });
    }

    private static Path defaultOutputPath() {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", OUTPUT_FILE_NAME);
        }
        // 从仓库根目录执行 mvn -pl playground 时，user.dir 是根目录。
        return Paths.get("playground", "target", OUTPUT_FILE_NAME);
    }
}

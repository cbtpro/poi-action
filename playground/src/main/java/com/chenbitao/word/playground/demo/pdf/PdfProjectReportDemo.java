package com.chenbitao.word.playground.demo.pdf;

import com.chenbitao.word.factory.PdfDocumentGeneratorFactory;
import com.chenbitao.word.pdf.PdfDocumentGenerator;
import lombok.extern.slf4j.Slf4j;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;

/**
 * PDF 项目汇报演示。
 *
 * <p>该演示生成一份包含标题、段落、表格和图片的 {@code .pdf} 文件。</p>
 */
@Slf4j
public class PdfProjectReportDemo {

    private static final String OUTPUT_FILE_NAME = "pdf-report-demo.pdf";

    /**
     * 生成 PDF 演示文件。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 如果目录创建、图片生成或文件写出失败
     */
    public static void main(String[] args) throws Exception {
        Path outputPath = defaultOutputPath();
        generate(outputPath);
        log.info("PDF 演示文件生成成功：{}", outputPath.toAbsolutePath());
    }

    /**
     * 生成项目汇报 PDF。
     *
     * @param outputPath 输出路径
     * @throws Exception 如果目录创建、图片生成或文件写出失败
     */
    public static void generate(Path outputPath) throws Exception {
        Files.createDirectories(outputPath.getParent());

        PdfDocumentGenerator generator = PdfDocumentGeneratorFactory.get("pdf");
        generator.createDocument();
        generator.addTitle("poi-action PDF Report");
        generator.addParagraph("PDF generation is powered by Apache PDFBox.");
        generator.addParagraph("This demo writes text, table data and an image into a PDF document.");
        generator.addTable(Arrays.asList(
                Arrays.asList("Document", "Format", "Status"),
                Arrays.asList("Word", ".doc / .docx", "Supported"),
                Arrays.asList("Excel", ".xls / .xlsx", "Supported"),
                Arrays.asList("PowerPoint", ".ppt / .pptx", "Supported"),
                Arrays.asList("PDF", ".pdf", "Supported")
        ));
        generator.addImage(new ByteArrayInputStream(demoImage()), 240F, 120F);
        generator.save(outputPath.toString());
    }

    private static byte[] demoImage() throws Exception {
        BufferedImage image = new BufferedImage(480, 240, BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < image.getWidth(); x++) {
            for (int y = 0; y < image.getHeight(); y++) {
                int red = 60 + (x * 80 / image.getWidth());
                int green = 120 + (y * 70 / image.getHeight());
                int blue = 150 + ((x + y) * 40 / (image.getWidth() + image.getHeight()));
                image.setRGB(x, y, new Color(red, green, blue).getRGB());
            }
        }
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
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

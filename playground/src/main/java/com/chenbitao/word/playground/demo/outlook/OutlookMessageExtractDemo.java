package com.chenbitao.word.playground.demo.outlook;

import com.chenbitao.word.factory.OutlookMessageReaderFactory;
import com.chenbitao.word.outlook.OutlookAttachmentInfo;
import com.chenbitao.word.outlook.OutlookMessageInfo;
import com.chenbitao.word.outlook.OutlookMessageReader;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.StringUtil;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * Outlook MSG 邮件读取演示。
 *
 * <p>该演示先生成一份最小 {@code .msg} 示例文件，再提取主题、正文和附件摘要。</p>
 */
@Slf4j
public class OutlookMessageExtractDemo {

    private static final String OUTPUT_DIR_NAME = "outlook";
    private static final String SAMPLE_FILE_NAME = "weekly-report-demo.msg";
    private static final String SUMMARY_FILE_NAME = "weekly-report-summary.txt";

    /**
     * 运行 Outlook 邮件读取演示。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 如果目录创建、示例邮件写出或摘要写出失败
     */
    public static void main(String[] args) throws Exception {
        Path outputDirectory = defaultOutputDirectory();
        OutlookMessageInfo message = generate(outputDirectory);

        log.info("Outlook 邮件示例生成成功：{}", outputDirectory.resolve(SAMPLE_FILE_NAME).toAbsolutePath());
        log.info("Outlook 邮件摘要生成成功：{}", outputDirectory.resolve(SUMMARY_FILE_NAME).toAbsolutePath());
        log.info("邮件主题：{}，附件数量：{}", message.getSubject(), message.getAttachmentCount());
    }

    /**
     * 生成示例 MSG 并输出解析摘要。
     *
     * @param outputDirectory 输出目录
     * @return 解析后的邮件摘要信息
     * @throws Exception 如果文件写出或读取失败
     */
    public static OutlookMessageInfo generate(Path outputDirectory) throws Exception {
        Files.createDirectories(outputDirectory);

        Path samplePath = outputDirectory.resolve(SAMPLE_FILE_NAME);
        Path summaryPath = outputDirectory.resolve(SUMMARY_FILE_NAME);
        writeSampleMessage(samplePath);

        OutlookMessageReader reader = OutlookMessageReaderFactory.get("msg");
        OutlookMessageInfo message = reader.read(samplePath);
        Files.write(summaryPath, summaryLines(message), StandardCharsets.UTF_8);
        return message;
    }

    private static List<String> summaryLines(OutlookMessageInfo message) {
        List<String> lines = new ArrayList<>();
        lines.add("主题: " + message.getSubject());
        lines.add("发件人: " + message.getFrom());
        lines.add("收件人: " + message.getTo());
        lines.add("正文: " + message.getTextBody());
        lines.add("附件数量: " + message.getAttachmentCount());
        for (OutlookAttachmentInfo attachment : message.getAttachments()) {
            lines.add("附件: " + attachment.getFileName()
                    + ", 类型: " + attachment.getMimeType()
                    + ", 大小: " + attachment.getSize());
        }
        return lines;
    }

    private static void writeSampleMessage(Path path) throws Exception {
        try (POIFSFileSystem fileSystem = new POIFSFileSystem();
             FileOutputStream output = new FileOutputStream(path.toFile())) {
            writeUnicode(fileSystem.getRoot(), "__substg1.0_001A001F", "IPM.Note");
            writeUnicode(fileSystem.getRoot(), "__substg1.0_0037001F", "项目周报");
            writeUnicode(fileSystem.getRoot(), "__substg1.0_0C1A001F", "张三");
            writeUnicode(fileSystem.getRoot(), "__substg1.0_0E04001F", "李四");
            writeUnicode(fileSystem.getRoot(), "__substg1.0_1000001F", "本周已完成 Outlook 邮件读取能力。");

            DirectoryEntry attachment = fileSystem.getRoot().createDirectory("__attach_version1.0_#00000000");
            writeUnicode(attachment, "__substg1.0_3707001F", "weekly-report.txt");
            writeUnicode(attachment, "__substg1.0_370E001F", "text/plain");
            writeBytes(attachment, "__substg1.0_37010102", "hello outlook".getBytes(StandardCharsets.UTF_8));

            fileSystem.writeFilesystem(output);
        }
    }

    private static void writeUnicode(DirectoryEntry directory, String name, String value) throws Exception {
        writeBytes(directory, name, StringUtil.getToUnicodeLE(value));
    }

    private static void writeBytes(DirectoryEntry directory, String name, byte[] bytes) throws Exception {
        directory.createDocument(name, new ByteArrayInputStream(bytes));
    }

    private static Path defaultOutputDirectory() {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", OUTPUT_DIR_NAME);
        }
        // 从仓库根目录执行 mvn -pl playground 时，user.dir 是根目录。
        return Paths.get("playground", "target", OUTPUT_DIR_NAME);
    }
}

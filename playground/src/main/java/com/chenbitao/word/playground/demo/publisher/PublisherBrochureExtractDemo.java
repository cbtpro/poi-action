package com.chenbitao.word.playground.demo.publisher;

import com.chenbitao.word.factory.PublisherDocumentReaderFactory;
import com.chenbitao.word.publisher.PublisherDocumentInfo;
import com.chenbitao.word.publisher.PublisherDocumentReader;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.LittleEndian;
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
 * Publisher PUB document read demo.
 */
@Slf4j
public class PublisherBrochureExtractDemo {

    private static final String OUTPUT_DIR_NAME = "publisher";
    private static final String SAMPLE_FILE_NAME = "brochure-demo.pub";
    private static final String SUMMARY_FILE_NAME = "brochure-summary.txt";

    /**
     * Run the Publisher document read demo.
     *
     * @param args command line arguments, currently unused
     * @throws Exception if sample or summary generation fails
     */
    public static void main(String[] args) throws Exception {
        Path outputDirectory = defaultOutputDirectory();
        PublisherDocumentInfo document = generate(outputDirectory);

        log.info("Publisher sample generated: {}", outputDirectory.resolve(SAMPLE_FILE_NAME).toAbsolutePath());
        log.info("Publisher summary generated: {}", outputDirectory.resolve(SUMMARY_FILE_NAME).toAbsolutePath());
        log.info("Title: {}, has text: {}", document.getTitle(), document.hasText());
    }

    /**
     * Generate a minimal Publisher sample and write its extracted summary.
     *
     * @param outputDirectory output directory
     * @return extracted Publisher document summary
     * @throws Exception if file IO or document parsing fails
     */
    public static PublisherDocumentInfo generate(Path outputDirectory) throws Exception {
        Files.createDirectories(outputDirectory);

        Path samplePath = outputDirectory.resolve(SAMPLE_FILE_NAME);
        Path summaryPath = outputDirectory.resolve(SUMMARY_FILE_NAME);
        writeSamplePublisher(samplePath);

        PublisherDocumentReader reader = PublisherDocumentReaderFactory.get("pub");
        PublisherDocumentInfo document = reader.read(samplePath);
        Files.write(summaryPath, summaryLines(document), StandardCharsets.UTF_8);
        return document;
    }

    private static List<String> summaryLines(PublisherDocumentInfo document) {
        List<String> lines = new ArrayList<>();
        lines.add("Title: " + document.getTitle());
        lines.add("Subject: " + document.getSubject());
        lines.add("Author: " + document.getAuthor());
        lines.add("Keywords: " + document.getKeywords());
        lines.add("Comments: " + document.getComments());
        lines.add("Text: " + document.getText());
        return lines;
    }

    private static void writeSamplePublisher(Path path) throws Exception {
        try (POIFSFileSystem fileSystem = new POIFSFileSystem();
             FileOutputStream output = new FileOutputStream(path.toFile())) {
            DirectoryEntry root = fileSystem.getRoot();
            root.createDocument("Contents", new ByteArrayInputStream(new byte[0]));

            DirectoryEntry quill = root.createDirectory("Quill");
            DirectoryEntry quillSub = quill.createDirectory("QuillSub");
            quillSub.createDocument("CONTENTS", new ByteArrayInputStream(quillContents(
                    "Publisher brochure extraction\rMetadata and text summary"
            )));

            DirectoryEntry escher = root.createDirectory("Escher");
            escher.createDocument("EscherStm", new ByteArrayInputStream(new byte[0]));
            escher.createDocument("EscherDelayStm", new ByteArrayInputStream(new byte[0]));

            SummaryInformation summary = PropertySetFactory.newSummaryInformation();
            summary.setTitle("Product Brochure");
            summary.setSubject("Marketing");
            summary.setAuthor("poi-action");
            summary.setKeywords("publisher,poi");
            summary.setComments("Generated Publisher demo");
            summary.write(root, SummaryInformation.DEFAULT_STREAM_NAME);

            fileSystem.writeFilesystem(output);
        }
    }

    private static byte[] quillContents(String text) {
        byte[] textBytes = StringUtil.getToUnicodeLE(text);
        int dataOffset = 0x20 + 20 * 24;
        byte[] data = new byte[dataOffset + textBytes.length];
        byte[] header = "CHNKINK ".getBytes(StandardCharsets.US_ASCII);
        System.arraycopy(header, 0, data, 0, header.length);

        int offset = 0x20;
        data[offset] = 0x18;
        data[offset + 1] = 0x00;
        ascii(data, offset + 2, "TEXT");
        LittleEndian.putUShort(data, offset + 6, 0);
        LittleEndian.putUShort(data, offset + 8, 0);
        LittleEndian.putUShort(data, offset + 10, 0);
        ascii(data, offset + 12, "TEXT");
        LittleEndian.putInt(data, offset + 16, dataOffset);
        LittleEndian.putInt(data, offset + 20, textBytes.length);
        System.arraycopy(textBytes, 0, data, dataOffset, textBytes.length);
        return data;
    }

    private static void ascii(byte[] data, int offset, String value) {
        byte[] bytes = value.getBytes(StandardCharsets.US_ASCII);
        System.arraycopy(bytes, 0, data, offset, bytes.length);
    }

    private static Path defaultOutputDirectory() {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", OUTPUT_DIR_NAME);
        }
        return Paths.get("playground", "target", OUTPUT_DIR_NAME);
    }
}

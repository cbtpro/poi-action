package com.chenbitao.word.playground.demo.visio;

import com.chenbitao.word.factory.VisioDrawingReaderFactory;
import com.chenbitao.word.visio.VisioDrawingInfo;
import com.chenbitao.word.visio.VisioDrawingReader;
import com.chenbitao.word.visio.VisioPageInfo;
import com.chenbitao.word.visio.VisioShapeInfo;
import lombok.extern.slf4j.Slf4j;

import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Visio VSDX drawing read demo.
 */
@Slf4j
public class VisioWorkflowExtractDemo {

    private static final String OUTPUT_DIR_NAME = "visio";
    private static final String SAMPLE_FILE_NAME = "workflow-demo.vsdx";
    private static final String SUMMARY_FILE_NAME = "workflow-summary.txt";

    /**
     * Run the Visio drawing read demo.
     *
     * @param args command line arguments, currently unused
     * @throws Exception if sample or summary generation fails
     */
    public static void main(String[] args) throws Exception {
        Path outputDirectory = defaultOutputDirectory();
        VisioDrawingInfo drawing = generate(outputDirectory);

        log.info("Visio sample generated: {}", outputDirectory.resolve(SAMPLE_FILE_NAME).toAbsolutePath());
        log.info("Visio summary generated: {}", outputDirectory.resolve(SUMMARY_FILE_NAME).toAbsolutePath());
        log.info("Pages: {}, shapes: {}", drawing.getPageCount(), drawing.getShapeCount());
    }

    /**
     * Generate a minimal VSDX sample and write its extracted summary.
     *
     * @param outputDirectory output directory
     * @return extracted Visio drawing summary
     * @throws Exception if file IO or drawing parsing fails
     */
    public static VisioDrawingInfo generate(Path outputDirectory) throws Exception {
        Files.createDirectories(outputDirectory);

        Path samplePath = outputDirectory.resolve(SAMPLE_FILE_NAME);
        Path summaryPath = outputDirectory.resolve(SUMMARY_FILE_NAME);
        writeSampleVsdx(samplePath);

        VisioDrawingReader reader = VisioDrawingReaderFactory.get("vsdx");
        VisioDrawingInfo drawing = reader.read(samplePath);
        Files.write(summaryPath, summaryLines(drawing), StandardCharsets.UTF_8);
        return drawing;
    }

    private static List<String> summaryLines(VisioDrawingInfo drawing) {
        List<String> lines = new ArrayList<>();
        lines.add("Format: " + drawing.getFormat());
        lines.add("Pages: " + drawing.getPageCount());
        lines.add("Shapes: " + drawing.getShapeCount());
        for (VisioPageInfo page : drawing.getPages()) {
            lines.add("Page: " + page.getName()
                    + ", size: " + page.getWidth() + " x " + page.getHeight()
                    + ", connections: " + page.getConnectionCount());
            for (VisioShapeInfo shape : page.getShapes()) {
                lines.add("Shape: " + shape.getName()
                        + ", type: " + shape.getType()
                        + ", text: " + shape.getText());
            }
        }
        return lines;
    }

    private static void writeSampleVsdx(Path path) throws Exception {
        try (ZipOutputStream zip = new ZipOutputStream(new FileOutputStream(path.toFile()), StandardCharsets.UTF_8)) {
            entry(zip, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                            + "<Override PartName=\"/visio/document.xml\" ContentType=\"application/vnd.ms-visio.drawing.main+xml\"/>"
                            + "<Override PartName=\"/visio/pages/pages.xml\" ContentType=\"application/vnd.ms-visio.pages+xml\"/>"
                            + "<Override PartName=\"/visio/pages/page1.xml\" ContentType=\"application/vnd.ms-visio.page+xml\"/>"
                            + "</Types>");
            entry(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                            + "<Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/visio/2010/relationships/document\" Target=\"visio/document.xml\"/>"
                            + "</Relationships>");
            entry(zip, "visio/_rels/document.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                            + "<Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/visio/2010/relationships/pages\" Target=\"pages/pages.xml\"/>"
                            + "</Relationships>");
            entry(zip, "visio/pages/_rels/pages.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                            + "<Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/visio/2010/relationships/page\" Target=\"page1.xml\"/>"
                            + "</Relationships>");
            entry(zip, "visio/document.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                            + "<VisioDocument xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\">"
                            + "<DocumentSettings/>"
                            + "<Colors/>"
                            + "<FaceNames/>"
                            + "<StyleSheets/>"
                            + "</VisioDocument>");
            entry(zip, "visio/pages/pages.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                            + "<Pages xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\" "
                            + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                            + "<Page ID=\"0\" Name=\"Workflow\" NameU=\"Workflow\" IsCustomName=\"1\" "
                            + "Background=\"0\" ViewScale=\"1\" ViewCenterX=\"4.25\" ViewCenterY=\"5.5\">"
                            + "<PageSheet>"
                            + "<Cell N=\"PageWidth\" V=\"8.5\"/>"
                            + "<Cell N=\"PageHeight\" V=\"11\"/>"
                            + "</PageSheet>"
                            + "<Rel r:id=\"rId1\"/>"
                            + "</Page>"
                            + "</Pages>");
            entry(zip, "visio/pages/page1.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                            + "<PageContents xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\">"
                            + "<Shapes>"
                            + shape("1", "Start", "2", "8", "2", "1", "Start review")
                            + shape("2", "Decision", "5", "8", "2", "1", "Approve")
                            + "</Shapes>"
                            + "<Connects>"
                            + "<Connect FromSheet=\"1\" ToSheet=\"2\" FromCell=\"EndX\" ToCell=\"BeginX\"/>"
                            + "</Connects>"
                            + "</PageContents>");
        }
    }

    private static String shape(String id, String name, String pinX, String pinY, String width, String height, String text) {
        return "<Shape ID=\"" + id + "\" Name=\"" + name + "\" NameU=\"" + name + "\" Type=\"Shape\">"
                + "<Cell N=\"PinX\" V=\"" + pinX + "\"/>"
                + "<Cell N=\"PinY\" V=\"" + pinY + "\"/>"
                + "<Cell N=\"Width\" V=\"" + width + "\"/>"
                + "<Cell N=\"Height\" V=\"" + height + "\"/>"
                + "<Text>" + text + "</Text>"
                + "</Shape>";
    }

    private static void entry(ZipOutputStream zip, String name, String value) throws Exception {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(value.getBytes(StandardCharsets.UTF_8));
        zip.closeEntry();
    }

    private static Path defaultOutputDirectory() {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", OUTPUT_DIR_NAME);
        }
        return Paths.get("playground", "target", OUTPUT_DIR_NAME);
    }
}

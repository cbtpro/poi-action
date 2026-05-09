package com.chenbitao.word.visio;

import com.chenbitao.word.factory.VisioDrawingReaderFactory;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class VisioDrawingReaderTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void readVsdxExtractsPagesShapesAndText() throws Exception {
        File drawing = temporaryFolder.newFile("workflow.vsdx");
        writeSampleVsdx(drawing);

        VisioDrawingInfo info = new VisioDrawingReader().read(drawing);

        assertEquals("vsdx", info.getFormat());
        assertEquals(1, info.getPageCount());
        assertEquals(2, info.getShapeCount());
        assertEquals(2, info.getTextItems().size());
        assertTrue(info.getTextItems().contains("Start review"));
        assertTrue(info.getTextItems().contains("Approve"));

        VisioPageInfo page = info.getPages().get(0);
        assertEquals("Workflow", page.getName());
        assertEquals(8.5D, page.getWidth(), 0.001D);
        assertEquals(11D, page.getHeight(), 0.001D);
        assertEquals(1, page.getConnectionCount());
        assertEquals(2, page.getShapes().size());

        VisioShapeInfo shape = page.getShapes().get(0);
        assertEquals(1L, shape.getId());
        assertEquals("Start", shape.getName());
        assertEquals("Shape", shape.getType());
        assertEquals("Start review", shape.getText());
        assertEquals(2D, shape.getPinX(), 0.001D);
        assertEquals(8D, shape.getPinY(), 0.001D);
        assertEquals(2D, shape.getWidth(), 0.001D);
        assertEquals(1D, shape.getHeight(), 0.001D);
    }

    @Test
    public void readInputStreamUsesProvidedType() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        writeSampleVsdx(output);

        VisioDrawingInfo info = new VisioDrawingReader()
                .read(new ByteArrayInputStream(output.toByteArray()), "vsdx");

        assertEquals(1, info.getPageCount());
        assertEquals(2, info.getShapeCount());
    }

    @Test
    public void factorySupportsVsdAndVsdxTypes() {
        assertTrue(VisioDrawingReaderFactory.get("vsd") instanceof VisioDrawingReader);
        assertTrue(VisioDrawingReaderFactory.get("vsdx") instanceof VisioDrawingReader);
    }

    private void writeSampleVsdx(File file) throws Exception {
        try (FileOutputStream output = new FileOutputStream(file)) {
            writeSampleVsdx(output);
        }
    }

    private void writeSampleVsdx(ByteArrayOutputStream output) throws Exception {
        try (ZipOutputStream zip = new ZipOutputStream(output, StandardCharsets.UTF_8)) {
            writeVsdxEntries(zip);
        }
    }

    private void writeSampleVsdx(FileOutputStream output) throws Exception {
        try (ZipOutputStream zip = new ZipOutputStream(output, StandardCharsets.UTF_8)) {
            writeVsdxEntries(zip);
        }
    }

    private void writeVsdxEntries(ZipOutputStream zip) throws Exception {
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

    private String shape(String id, String name, String pinX, String pinY, String width, String height, String text) {
        return "<Shape ID=\"" + id + "\" Name=\"" + name + "\" NameU=\"" + name + "\" Type=\"Shape\">"
                + "<Cell N=\"PinX\" V=\"" + pinX + "\"/>"
                + "<Cell N=\"PinY\" V=\"" + pinY + "\"/>"
                + "<Cell N=\"Width\" V=\"" + width + "\"/>"
                + "<Cell N=\"Height\" V=\"" + height + "\"/>"
                + "<Text>" + text + "</Text>"
                + "</Shape>";
    }

    private void entry(ZipOutputStream zip, String name, String value) throws Exception {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(value.getBytes(StandardCharsets.UTF_8));
        zip.closeEntry();
    }
}

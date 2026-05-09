package com.chenbitao.word.publisher;

import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.util.StringUtil;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class PublisherDocumentReaderTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void readExtractsMetadataAndText() throws Exception {
        File file = temporaryFolder.newFile("brochure.pub");
        writePublisher(file, "Product Brochure", "Marketing", "poi-action",
                "Publisher document extraction\rSupports text and metadata");

        PublisherDocumentInfo info = new PublisherDocumentReader().read(file);

        assertEquals("Product Brochure", info.getTitle());
        assertEquals("Marketing", info.getSubject());
        assertEquals("poi-action", info.getAuthor());
        assertEquals("publisher,poi", info.getKeywords());
        assertEquals("Generated Publisher sample", info.getComments());
        assertTrue(info.hasText());
        assertTrue(info.getText().contains("Publisher document extraction"));
        assertTrue(info.getText().contains("Supports text and metadata"));
    }

    @Test
    public void readInputStreamExtractsText() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        writePublisher(output, "Newsletter", "Demo", "poi-action", "Publisher newsletter");

        PublisherDocumentInfo info = new PublisherDocumentReader()
                .read(new ByteArrayInputStream(output.toByteArray()));

        assertEquals("Newsletter", info.getTitle());
        assertTrue(info.getText().contains("Publisher newsletter"));
    }

    private void writePublisher(File file, String title, String subject, String author, String text) throws Exception {
        try (FileOutputStream output = new FileOutputStream(file)) {
            writePublisher(output, title, subject, author, text);
        }
    }

    private void writePublisher(ByteArrayOutputStream output, String title, String subject, String author, String text) throws Exception {
        try (POIFSFileSystem fileSystem = createPublisher(title, subject, author, text)) {
            fileSystem.writeFilesystem(output);
        }
    }

    private void writePublisher(FileOutputStream output, String title, String subject, String author, String text) throws Exception {
        try (POIFSFileSystem fileSystem = createPublisher(title, subject, author, text)) {
            fileSystem.writeFilesystem(output);
        }
    }

    private POIFSFileSystem createPublisher(String title, String subject, String author, String text) throws Exception {
        POIFSFileSystem fileSystem = new POIFSFileSystem();
        DirectoryEntry root = fileSystem.getRoot();
        root.createDocument("Contents", new ByteArrayInputStream(new byte[0]));

        DirectoryEntry quill = root.createDirectory("Quill");
        DirectoryEntry quillSub = quill.createDirectory("QuillSub");
        quillSub.createDocument("CONTENTS", new ByteArrayInputStream(quillContents(text)));

        DirectoryEntry escher = root.createDirectory("Escher");
        escher.createDocument("EscherStm", new ByteArrayInputStream(new byte[0]));
        escher.createDocument("EscherDelayStm", new ByteArrayInputStream(new byte[0]));

        SummaryInformation summary = PropertySetFactory.newSummaryInformation();
        summary.setTitle(title);
        summary.setSubject(subject);
        summary.setAuthor(author);
        summary.setKeywords("publisher,poi");
        summary.setComments("Generated Publisher sample");
        summary.write(root, SummaryInformation.DEFAULT_STREAM_NAME);
        return fileSystem;
    }

    private byte[] quillContents(String text) {
        byte[] textBytes = StringUtil.getToUnicodeLE(text);
        int dataOffset = 0x20 + 20 * 24;
        byte[] data = new byte[dataOffset + textBytes.length];
        byte[] header = "CHNKINK ".getBytes();
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

    private void ascii(byte[] data, int offset, String value) {
        byte[] bytes = value.getBytes();
        System.arraycopy(bytes, 0, data, offset, bytes.length);
    }
}

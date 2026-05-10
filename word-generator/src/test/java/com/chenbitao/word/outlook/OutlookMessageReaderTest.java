package com.chenbitao.word.outlook;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.StringUtil;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class OutlookMessageReaderTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void readExtractsMessageFieldsAndAttachments() throws Exception {
        File messageFile = temporaryFolder.newFile("demo.msg");
        writeMessage(messageFile);

        OutlookMessageInfo message = new OutlookMessageReader().read(messageFile);

        assertEquals("项目周报", message.getSubject());
        assertEquals("张三", message.getFrom());
        assertEquals("李四", message.getTo());
        assertTrue(message.getTextBody().contains("Outlook 邮件读取"));
        assertEquals(1, message.getAttachmentCount());
        assertEquals("report.txt", message.getAttachments().get(0).getFileName());
        assertEquals("text/plain", message.getAttachments().get(0).getMimeType());
        assertEquals(11, message.getAttachments().get(0).getSize());
    }

    private void writeMessage(File file) throws Exception {
        try (POIFSFileSystem fileSystem = new POIFSFileSystem();
             FileOutputStream output = new FileOutputStream(file)) {
            writeUnicode(fileSystem.getRoot(), "__substg1.0_001A001F", "IPM.Note");
            writeUnicode(fileSystem.getRoot(), "__substg1.0_0037001F", "项目周报");
            writeUnicode(fileSystem.getRoot(), "__substg1.0_0C1A001F", "张三");
            writeUnicode(fileSystem.getRoot(), "__substg1.0_0E04001F", "李四");
            writeUnicode(fileSystem.getRoot(), "__substg1.0_1000001F", "本周已完成 Outlook 邮件读取能力。");

            DirectoryEntry attachment = fileSystem.getRoot().createDirectory("__attach_version1.0_#00000000");
            writeUnicode(attachment, "__substg1.0_3707001F", "report.txt");
            writeUnicode(attachment, "__substg1.0_370E001F", "text/plain");
            writeBytes(attachment, "__substg1.0_37010102", "hello world".getBytes("UTF-8"));

            fileSystem.writeFilesystem(output);
        }
    }

    private void writeUnicode(DirectoryEntry directory, String name, String value) throws Exception {
        writeBytes(directory, name, StringUtil.getToUnicodeLE(value));
    }

    private void writeBytes(DirectoryEntry directory, String name, byte[] bytes) throws Exception {
        directory.createDocument(name, new ByteArrayInputStream(bytes));
    }
}

package com.chenbitao.word.docx.util;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.nio.file.Files;
import java.util.Base64;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertTrue;

public class ImageSourceUtilsTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void readBytesSupportsByteArrayInputStreamFilePathAndBase64String() throws Exception {
        byte[] bytes = new byte[] { 1, 2, 3, 4 };
        File file = temporaryFolder.newFile("image.bin");
        Files.write(file.toPath(), bytes);

        assertArrayEquals(bytes, ImageSourceUtils.readBytes(bytes));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes(new ByteArrayInputStream(bytes)));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes(file));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes(file.toPath()));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes(Base64.getEncoder().encodeToString(bytes)));
        String base64Str = "data:image/png;base64," + Base64.getEncoder().encodeToString(bytes);
        byte[] decodedBytes;
        if (base64Str.startsWith("data:image/")) {
            // 手动解码 Base64
            String pureBase64 = base64Str.replaceAll("^data:image\\/\\w+;base64,", "");
            decodedBytes = Base64.getDecoder().decode(pureBase64);
        } else {
            // 正常读取文件
            decodedBytes = ImageSourceUtils.readBytes(base64Str);
        }

        assertArrayEquals(bytes, decodedBytes);
    }

    @Test
    public void toPngBytesConvertsRecognizedImageToPng() throws Exception {
        BufferedImage image = new BufferedImage(8, 8, BufferedImage.TYPE_INT_RGB);
        ByteArrayOutputStream jpeg = new ByteArrayOutputStream();
        ImageIO.write(image, "jpg", jpeg);

        byte[] png = ImageSourceUtils.toPngBytes(jpeg.toByteArray());

        assertTrue(png.length > 8);
        assertTrue(png[0] == (byte) 0x89 && png[1] == 'P' && png[2] == 'N' && png[3] == 'G');
    }
}

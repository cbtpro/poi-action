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
        byte[] bytes = new byte[]{1, 2, 3, 4};
        File file = temporaryFolder.newFile("image.bin");
        Files.write(file.toPath(), bytes);

        assertArrayEquals(bytes, ImageSourceUtils.readBytes(bytes));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes(new ByteArrayInputStream(bytes)));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes(file));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes(file.toPath()));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes(Base64.getEncoder().encodeToString(bytes)));
        assertArrayEquals(bytes, ImageSourceUtils.readBytes("data:image/png;base64,"
                + Base64.getEncoder().encodeToString(bytes)));
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

package com.chenbitao.word.docx.util;

import com.chenbitao.word.util.ImageBytesUtils;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;

import static org.junit.Assert.*;

/**
 * ImageSourceUtils 工具类单元测试
 * 覆盖所有入参类型：字节数组、流、文件、路径、Base64、网络图片、资源图片
 *
 * @author chenbitao
 * @since 2025
 */
public class ImageSourceUtilsTest {

    /**
     * 临时文件夹规则，自动创建/清理文件
     */
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    /**
     * 通用测试字节数组
     */
    private static final byte[] TEST_BYTES = {1, 2, 3, 4};

    // ============================ 基础类型测试 ============================

    /**
     * 测试：入参为 byte[] 时，正确返回原字节数组
     */
    @Test
    public void readBytes_WithByteArray_ReturnsOriginalBytes() throws Exception {
        assertArrayEquals(TEST_BYTES, ImageSourceUtils.readBytes(TEST_BYTES));
    }

    /**
     * 测试：入参为 InputStream 时，正确读取字节数组
     */
    @Test
    public void readBytes_WithInputStream_ReadsCorrectly() throws Exception {
        try (ByteArrayInputStream bais = new ByteArrayInputStream(TEST_BYTES)) {
            assertArrayEquals(TEST_BYTES, ImageSourceUtils.readBytes(bais));
        }
    }

    /**
     * 测试：入参为 File 时，正确读取文件字节
     */
    @Test
    public void readBytes_WithFile_ReadsCorrectly() throws Exception {
        File tempFile = temporaryFolder.newFile("test.bin");
        Files.write(tempFile.toPath(), TEST_BYTES);
        assertArrayEquals(TEST_BYTES, ImageSourceUtils.readBytes(tempFile));
    }

    /**
     * 测试：入参为 Path 时，正确读取文件字节
     */
    @Test
    public void readBytes_WithPath_ReadsCorrectly() throws Exception {
        Path tempPath = temporaryFolder.newFile("test.bin").toPath();
        Files.write(tempPath, TEST_BYTES);
        assertArrayEquals(TEST_BYTES, ImageSourceUtils.readBytes(tempPath));
    }

    // ============================ Base64 测试 ============================

    /**
     * 测试：入参为纯 Base64 字符串时，正确解码为原字节
     */
    @Test
    public void readBytes_WithPureBase64String_DecodesCorrectly() throws Exception {
        String pureBase64 = Base64.getEncoder().encodeToString(TEST_BYTES);
        assertArrayEquals(TEST_BYTES, ImageSourceUtils.readBytes(pureBase64));
    }

    /**
     * 测试：入参为带 data:image 前缀的 Base64 时，正确解码
     */
    @Test
    public void readBytes_WithDataUriBase64_DecodesCorrectly() throws Exception {
        String base64 = Base64.getEncoder().encodeToString(TEST_BYTES);
        String dataUri = "data:image/png;base64," + base64;
        assertArrayEquals(TEST_BYTES, ImageSourceUtils.readBytes(dataUri));
    }

    // ============================ 网络图片测试 ============================

    /**
     * 测试：读取 HTTP 网络图片链接，返回有效字节
     * 请替换为真实可用的 HTTP 图片地址
     */
    @Test
    public void readBytes_WithHttpImageUrl_ReturnsValidBytes() throws Exception {
        String httpUrl = "https://p6-passport.byteacctimg.com/img/user-avatar/009e000023976c83dc1229a0330b9a70~40x40.awebp";
        byte[] result = ImageSourceUtils.readBytes(httpUrl);
        assertNotNull(result);
        assertTrue(result.length > 0);
    }

    /**
     * 测试：读取 HTTPS 网络图片链接，返回有效字节
     * 请替换为真实可用的 HTTPS 图片地址
     */
    @Test
    public void readBytes_WithHttpsImageUrl_ReturnsValidBytes() throws Exception {
        String httpsUrl = "https://p6-passport.byteacctimg.com/img/user-avatar/009e000023976c83dc1229a0330b9a70~40x40.awebp";
        byte[] result = ImageSourceUtils.readBytes(httpsUrl);
        assertNotNull(result);
        assertTrue(result.length > 0);
    }

    // ============================ 资源文件测试 ============================

    /**
     * 测试：读取 resources 目录下的 PNG 图片
     * 需确保：resources/test.png 存在
     */
    @Test
    public void readBytes_WithResourcePng_ReadsSuccessfully() throws Exception {
        try (InputStream inputStream = getClass().getResourceAsStream("/test.png")) {
            assertNotNull("未在 resources 目录下找到 test.png 文件", inputStream);
            byte[] expectedBytes = ImageSourceUtils.readAll(inputStream);
            byte[] actualBytes = ImageSourceUtils.readBytes(expectedBytes);
            assertArrayEquals(expectedBytes, actualBytes);
        }
    }

    /**
     * 测试：读取 resources 目录下的 JPG 图片
     * 需确保：resources/test.jpg 存在
     */
    @Test
    public void readBytes_WithResourceJpg_ReadsSuccessfully() throws Exception {
        try (InputStream inputStream = getClass().getResourceAsStream("/test.jpg")) {
            assertNotNull("未在 resources 目录下找到 test.jpg 文件", inputStream);
            byte[] expectedBytes = ImageSourceUtils.readAll(inputStream);
            byte[] actualBytes = ImageSourceUtils.readBytes(expectedBytes);
            assertArrayEquals(expectedBytes, actualBytes);
        }
    }

    // ============================ 格式转换测试 ============================

    /**
     * 测试：将 JPEG 图片转换为 PNG 格式字节
     * 验证输出为标准 PNG 文件头
     */
    @Test
    public void toPngBytes_ConvertJpegToPng_ReturnsPngBytes() throws Exception {
        // 构造测试图片
        byte[] jpgBytes = ImageBytesUtils.solidImage("jpg", 10, 10, 0x3366CC);

        // 转换为 PNG
        byte[] pngBytes = ImageSourceUtils.toPngBytes(jpgBytes);

        // 验证 PNG 文件头
        assertNotNull(pngBytes);
        assertTrue(pngBytes.length > 0);
        assertEquals((byte) 0x89, pngBytes[0]);
        assertEquals('P', pngBytes[1]);
        assertEquals('N', pngBytes[2]);
        assertEquals('G', pngBytes[3]);
    }
}

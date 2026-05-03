package com.chenbitao.word.docx.util;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Base64;

/**
 * 图片来源读取和格式转换工具。
 */
public final class ImageSourceUtils {

    private ImageSourceUtils() {
    }

    /**
     * 将图片来源读取为字节数组。
     *
     * @param source 图片来源，支持
     *               byte[]、InputStream、File、Path、URL、URI、路径字符串、URL字符串和Base64字符串
     * @return 图片字节数组
     * @throws IOException 读取失败时抛出
     */
    public static byte[] readBytes(Object source) throws IOException {
        if (source == null) {
            return new byte[0];
        }
        if (source instanceof byte[]) {
            return (byte[]) source;
        }
        if (source instanceof InputStream) {
            return readAll((InputStream) source);
        }
        if (source instanceof File) {
            return Files.readAllBytes(((File) source).toPath());
        }
        if (source instanceof Path) {
            return Files.readAllBytes((Path) source);
        }
        if (source instanceof URL) {
            return readUrlBytes((URL) source);
        }
        if (source instanceof URI) {
            return readUrlBytes(((URI) source).toURL());
        }
        if (source instanceof String) {
            return readStringSource((String) source);
        }
        throw new IllegalArgumentException("不支持的图片输入类型：" + source.getClass().getName());
    }

    /**
     * 将图片来源转换为 PNG 字节。
     * 如果来源不是 ImageIO 可识别图片，则返回原始字节。
     *
     * @param source 图片来源
     * @return PNG 字节或原始字节
     * @throws IOException 读取或转换失败时抛出
     */
    public static byte[] toPngBytes(Object source) throws IOException {
        byte[] bytes = readBytes(source);
        BufferedImage image = ImageIO.read(new ByteArrayInputStream(bytes));
        if (image == null) {
            return bytes;
        }

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
    }

    /**
     * 读取输入流所有字节，并关闭输入流。
     *
     * @param inputStream 输入流
     * @return 字节数组
     * @throws IOException 读取失败时抛出
     */
    public static byte[] readAll(InputStream inputStream) throws IOException {
        try (InputStream in = inputStream; ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[8192];
            int read;
            while ((read = in.read(buffer)) != -1) {
                output.write(buffer, 0, read);
            }
            return output.toByteArray();
        }
    }

    /**
     * 通过 URL 读取字节。
     *
     * @param url 图片 URL
     * @return 字节数组
     * @throws IOException 读取失败时抛出
     */
    private static byte[] readUrlBytes(URL url) throws IOException {
        try (InputStream inputStream = url.openStream()) {
            return readAll(inputStream);
        }
    }

    /**
     * 处理字符串来源，支持路径、URL和Base64。
     *
     * @param source 字符串来源
     * @return 字节数组
     * @throws IOException 读取失败时抛出
     */
    private static byte[] readStringSource(String source) throws IOException {
        String value = source == null ? "" : source.trim();
        if (value.isEmpty()) {
            return new byte[0];
        }

        // 1. 处理网络地址
        if (value.startsWith("http://") || value.startsWith("https://")) {
            return readUrlBytes(new URL(value));
        }

        // 2. 判断是否是 Base64（带前缀 或 不带前缀）
        if (value.startsWith("data:") || isBase64(value)) {
            int commaIndex = value.indexOf(',');
            String base64Data = commaIndex >= 0 ? value.substring(commaIndex + 1) : value;
            return Base64.getDecoder().decode(base64Data);
        }

        // 3. 处理文件路径
        Path path = Paths.get(value);
        if (Files.exists(path)) {
            return Files.readAllBytes(path);
        }

        throw new IOException("不支持的资源：不是有效路径、URL或Base64字符串：" + value);
    }

    /**
     * 简单判断字符串是否是纯 Base64 格式
     *
     * @param str 待判断字符串
     * @return 如果是纯 Base64 格式返回 true，否则返回 false
     */
    private static boolean isBase64(String str) {
        if (str == null || str.isEmpty()) {
            return false;
        }
        // Base64 只包含：A-Z a-z 0-9 + / =
        return str.matches("[A-Za-z0-9+/=]+");
    }
}

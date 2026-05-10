package com.chenbitao.word.doc;

import com.chenbitao.word.docx.TemplateWordGenerator;
import com.chenbitao.word.exception.WordException;
import org.apache.poi.hwpf.HWPFDocument;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * DOC 模板文档生成器。
 *
 * <p>基于 Apache POI HWPF 渲染 Word 97-2003 的 {@code .doc} 模板，支持
 * {@code ${key}} 文本占位符、点号分隔的嵌套字段，以及列表/数组值的多行输出。</p>
 */
public class TemplateDocWordGenerator {

    /** 普通占位符正则表达式。 */
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    /** DOC 文档对象。 */
    private final HWPFDocument document;

    /** 渲染后的 DOC 二进制内容。 */
    private byte[] renderedBytes;

    /**
     * 构造 DOC 模板文档生成器。
     *
     * @param template 模板文档输入流
     * @throws WordException 如果模板加载失败
     */
    public TemplateDocWordGenerator(InputStream template) {
        try {
            this.renderedBytes = readAll(template);
            this.document = new HWPFDocument(new ByteArrayInputStream(renderedBytes));
        } catch (Exception e) {
            throw new WordException("加载 DOC 模板失败", e);
        }
    }

    /**
     * 渲染模板文档。
     *
     * @param data 模板渲染数据
     */
    public void render(Map<String, Object> data) {
        for (Placeholder placeholder : placeholders(document.getRange().text())) {
            ResolvedValue value = resolveValue(placeholder.key, data);
            if (value.found) {
                replaceInBinary(placeholder.raw, formatValue(value.value));
            }
        }
    }

    /**
     * 提取文档中的占位符。
     *
     * @param text 模板文本
     * @return 占位符列表
     */
    private List<Placeholder> placeholders(String text) {
        List<Placeholder> placeholders = new ArrayList<>();
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(text);
        while (matcher.find()) {
            placeholders.add(new Placeholder(matcher.group(0), matcher.group(1)));
        }
        return placeholders;
    }

    /**
     * 在 DOC 原始二进制中执行等长 Unicode 文本替换。
     *
     * @param placeholder 占位符文本
     * @param replacement 替换文本
     */
    private void replaceInBinary(String placeholder, String replacement) {
        if (replacement.length() > placeholder.length()) {
            throw new WordException("DOC 模板替换值长度不能超过占位符长度：" + placeholder, null);
        }
        byte[] source = placeholder.getBytes(StandardCharsets.UTF_16LE);
        byte[] target = padRight(replacement, placeholder.length()).getBytes(StandardCharsets.UTF_16LE);
        int index = indexOf(renderedBytes, source);
        while (index >= 0) {
            System.arraycopy(target, 0, renderedBytes, index, target.length);
            index = indexOf(renderedBytes, source, index + target.length);
        }
    }

    /**
     * 右侧补空格到指定长度。
     *
     * @param text 原始文本
     * @param length 目标长度
     * @return 补齐后的文本
     */
    private String padRight(String text, int length) {
        StringBuilder builder = new StringBuilder(text == null ? "" : text);
        while (builder.length() < length) {
            builder.append(' ');
        }
        return builder.toString();
    }

    /**
     * 查找字节数组第一次出现的位置。
     *
     * @param source 源字节数组
     * @param target 目标字节数组
     * @return 起始位置，未找到时返回 -1
     */
    private int indexOf(byte[] source, byte[] target) {
        return indexOf(source, target, 0);
    }

    /**
     * 从指定位置开始查找字节数组。
     *
     * @param source 源字节数组
     * @param target 目标字节数组
     * @param from 起始位置
     * @return 起始位置，未找到时返回 -1
     */
    private int indexOf(byte[] source, byte[] target, int from) {
        for (int i = Math.max(0, from); i <= source.length - target.length; i++) {
            if (matchesAt(source, target, i)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 判断目标字节是否在指定位置匹配。
     *
     * @param source 源字节数组
     * @param target 目标字节数组
     * @param index 起始位置
     * @return 如果匹配则返回 true
     */
    private boolean matchesAt(byte[] source, byte[] target, int index) {
        for (int i = 0; i < target.length; i++) {
            if (source[index + i] != target[i]) {
                return false;
            }
        }
        return true;
    }

    /**
     * 保存文档到输出流。
     *
     * @param outputStream 输出流
     * @throws WordException 如果写出失败
     */
    public void save(OutputStream outputStream) {
        try {
            outputStream.write(renderedBytes);
        } catch (Exception e) {
            throw new WordException("导出 DOC 模板文档失败", e);
        }
    }

    /**
     * 保存文档到指定路径。
     *
     * @param path 输出文件路径
     * @throws WordException 如果写出失败
     */
    public void save(String path) {
        try (FileOutputStream output = new FileOutputStream(path)) {
            save(output);
        } catch (Exception e) {
            throw new WordException("导出 DOC 模板文档失败", e);
        }
    }

    /**
     * 解析占位符对应的值。
     *
     * @param key 占位符键名
     * @param data 模板渲染数据
     * @return 解析结果
     */
    private ResolvedValue resolveValue(String key, Map<String, Object> data) {
        if (data == null) {
            return ResolvedValue.missing();
        }
        if (data.containsKey(key)) {
            return ResolvedValue.found(data.get(key));
        }
        return resolveParts(data, key.split("\\."), 0);
    }

    /**
     * 按点号分隔路径递归解析对象值。
     *
     * @param current 当前对象
     * @param parts 路径片段
     * @param index 当前片段索引
     * @return 解析结果
     */
    @SuppressWarnings("unchecked")
    private ResolvedValue resolveParts(Object current, String[] parts, int index) {
        if (index >= parts.length) {
            return ResolvedValue.found(current);
        }
        if (current instanceof Map) {
            Map<String, Object> map = (Map<String, Object>) current;
            String part = parts[index];
            if (!map.containsKey(part)) {
                return ResolvedValue.missing();
            }
            return resolveParts(map.get(part), parts, index + 1);
        }
        if (current instanceof Iterable) {
            return resolveIterable((Iterable<?>) current, parts, index);
        }
        if (current != null && current.getClass().isArray()) {
            return resolveArray(current, parts, index);
        }
        return ResolvedValue.missing();
    }

    /**
     * 解析集合中每个元素的同一路径。
     *
     * @param values 集合值
     * @param parts 路径片段
     * @param index 当前片段索引
     * @return 解析结果
     */
    private ResolvedValue resolveIterable(Iterable<?> values, String[] parts, int index) {
        List<Object> resolved = new ArrayList<>();
        boolean found = false;
        for (Object value : values) {
            ResolvedValue item = resolveParts(value, parts, index);
            if (item.found) {
                found = true;
                resolved.add(item.value);
            }
        }
        return found ? ResolvedValue.found(resolved) : ResolvedValue.missing();
    }

    /**
     * 解析数组中每个元素的同一路径。
     *
     * @param array 数组值
     * @param parts 路径片段
     * @param index 当前片段索引
     * @return 解析结果
     */
    private ResolvedValue resolveArray(Object array, String[] parts, int index) {
        List<Object> resolved = new ArrayList<>();
        boolean found = false;
        int length = Array.getLength(array);
        for (int i = 0; i < length; i++) {
            ResolvedValue item = resolveParts(Array.get(array, i), parts, index);
            if (item.found) {
                found = true;
                resolved.add(item.value);
            }
        }
        return found ? ResolvedValue.found(resolved) : ResolvedValue.missing();
    }

    /**
     * 格式化模板值。
     *
     * @param value 模板值
     * @return 可写入 DOC 的文本
     */
    private String formatValue(Object value) {
        if (value == null) {
            return "";
        }
        if (isPictureValue(value)) {
            return "[图片]";
        }
        if (value instanceof Iterable) {
            return joinIterable((Iterable<?>) value);
        }
        if (value.getClass().isArray()) {
            return joinArray(value);
        }
        return String.valueOf(value);
    }

    /**
     * 判断值是否表示图片。
     *
     * @param value 待判断的值
     * @return 如果是图片值则返回 true
     */
    private boolean isPictureValue(Object value) {
        return value instanceof TemplateWordGenerator.Picture
                || value instanceof InputStream
                || value instanceof byte[]
                || value instanceof File
                || value instanceof Path;
    }

    /**
     * 将集合值拼接为多行文本。
     *
     * @param values 集合值
     * @return 多行文本
     */
    private String joinIterable(Iterable<?> values) {
        StringBuilder text = new StringBuilder();
        for (Object value : values) {
            appendValueLine(text, value);
        }
        return text.toString();
    }

    /**
     * 将数组值拼接为多行文本。
     *
     * @param array 数组值
     * @return 多行文本
     */
    private String joinArray(Object array) {
        StringBuilder text = new StringBuilder();
        int length = Array.getLength(array);
        for (int i = 0; i < length; i++) {
            appendValueLine(text, Array.get(array, i));
        }
        return text.toString();
    }

    /**
     * 向文本缓冲区追加一行值。
     *
     * @param text 文本缓冲区
     * @param value 需要追加的值
     */
    private void appendValueLine(StringBuilder text, Object value) {
        if (text.length() > 0) {
            text.append('\n');
        }
        text.append(value == null ? "" : formatValue(value));
    }

    /**
     * 读取输入流全部字节。
     *
     * @param inputStream 输入流
     * @return 字节数组
     * @throws IOException 如果读取失败
     */
    private byte[] readAll(InputStream inputStream) throws IOException {
        try (InputStream input = inputStream;
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[8192];
            int length;
            while ((length = input.read(buffer)) != -1) {
                output.write(buffer, 0, length);
            }
            return output.toByteArray();
        }
    }

    /**
     * 占位符解析结果。
     */
    private static final class ResolvedValue {
        /** 是否成功解析。 */
        private final boolean found;
        /** 解析得到的值。 */
        private final Object value;

        /**
         * 构造解析结果。
         *
         * @param found 是否成功解析
         * @param value 解析得到的值
         */
        private ResolvedValue(boolean found, Object value) {
            this.found = found;
            this.value = value;
        }

        /**
         * 创建已解析结果。
         *
         * @param value 解析得到的值
         * @return 解析结果
         */
        private static ResolvedValue found(Object value) {
            return new ResolvedValue(true, value);
        }

        /**
         * 创建未解析结果。
         *
         * @return 解析结果
         */
        private static ResolvedValue missing() {
            return new ResolvedValue(false, null);
        }
    }

    /**
     * 占位符信息。
     */
    private static final class Placeholder {
        /** 原始占位符文本。 */
        private final String raw;
        /** 占位符键名。 */
        private final String key;

        /**
         * 构造占位符信息。
         *
         * @param raw 原始占位符文本
         * @param key 占位符键名
         */
        private Placeholder(String raw, String key) {
            this.raw = raw;
            this.key = key;
        }
    }
}

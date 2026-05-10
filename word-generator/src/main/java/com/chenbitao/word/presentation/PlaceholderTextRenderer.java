package com.chenbitao.word.presentation;

import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 占位符文本渲染器。
 *
 * <p>用于演示文稿模板中的 {@code ${key}} 文本替换，支持点号分隔的嵌套字段、
 * 列表字段和数组多行输出。</p>
 */
final class PlaceholderTextRenderer {

    /** 普通占位符正则表达式。 */
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    /**
     * 私有构造方法，避免实例化工具类。
     */
    private PlaceholderTextRenderer() {
    }

    /**
     * 判断文本是否包含占位符。
     *
     * @param text 文本内容
     * @return 如果包含占位符则返回 true
     */
    static boolean containsPlaceholder(String text) {
        return text != null && PLACEHOLDER_PATTERN.matcher(text).find();
    }

    /**
     * 渲染文本中的占位符。
     *
     * @param template 模板文本
     * @param data 模板数据
     * @return 渲染后的文本
     */
    static String render(String template, Map<String, Object> data) {
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(template);
        StringBuffer result = new StringBuffer();
        while (matcher.find()) {
            String key = matcher.group(1);
            ResolvedValue value = resolveValue(key, data);
            String replacement = value.found ? formatValue(value.value) : matcher.group(0);
            matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(result);
        return result.toString();
    }

    /**
     * 解析占位符对应的值。
     *
     * @param key 占位符键名
     * @param data 模板数据
     * @return 解析结果
     */
    private static ResolvedValue resolveValue(String key, Map<String, Object> data) {
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
    private static ResolvedValue resolveParts(Object current, String[] parts, int index) {
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
    private static ResolvedValue resolveIterable(Iterable<?> values, String[] parts, int index) {
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
    private static ResolvedValue resolveArray(Object array, String[] parts, int index) {
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
     * 格式化值对象。
     *
     * @param value 值对象
     * @return 文本值
     */
    private static String formatValue(Object value) {
        if (value == null) {
            return "";
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
     * 将集合拼接为多行文本。
     *
     * @param values 集合值
     * @return 多行文本
     */
    private static String joinIterable(Iterable<?> values) {
        StringBuilder text = new StringBuilder();
        for (Object value : values) {
            appendValueLine(text, value);
        }
        return text.toString();
    }

    /**
     * 将数组拼接为多行文本。
     *
     * @param array 数组值
     * @return 多行文本
     */
    private static String joinArray(Object array) {
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
    private static void appendValueLine(StringBuilder text, Object value) {
        if (text.length() > 0) {
            text.append('\n');
        }
        text.append(value == null ? "" : formatValue(value));
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
}

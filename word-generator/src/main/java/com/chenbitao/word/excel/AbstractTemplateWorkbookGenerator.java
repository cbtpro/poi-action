package com.chenbitao.word.excel;

import com.chenbitao.word.exception.SpreadsheetException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 电子表格模板生成器基类。
 *
 * <p>负责遍历工作簿中的字符串单元格，并使用数据映射替换 {@code ${key}} 占位符。</p>
 */
abstract class AbstractTemplateWorkbookGenerator {

    /** 普通占位符正则表达式。 */
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    /** 模板工作簿对象。 */
    private final Workbook workbook;

    /**
     * 构造电子表格模板生成器。
     *
     * @param workbook 模板工作簿
     */
    AbstractTemplateWorkbookGenerator(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 渲染模板工作簿。
     *
     * @param data 模板数据
     */
    public void render(Map<String, Object> data) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            renderSheet(workbook.getSheetAt(i), data);
        }
        workbook.setForceFormulaRecalculation(true);
    }

    /**
     * 保存工作簿到指定路径。
     *
     * @param path 输出文件路径
     */
    public void save(String path) {
        try (FileOutputStream output = new FileOutputStream(path)) {
            save(output);
        } catch (Exception e) {
            throw new SpreadsheetException("保存模板工作簿失败", e);
        }
    }

    /**
     * 保存工作簿到输出流。
     *
     * @param outputStream 输出流
     */
    public void save(OutputStream outputStream) {
        try {
            workbook.write(outputStream);
        } catch (Exception e) {
            throw new SpreadsheetException("写出模板工作簿失败", e);
        }
    }

    /**
     * 渲染单个工作表。
     *
     * @param sheet 工作表
     * @param data 模板数据
     */
    private void renderSheet(Sheet sheet, Map<String, Object> data) {
        for (Row row : sheet) {
            renderRow(row, data);
        }
    }

    /**
     * 渲染单行。
     *
     * @param row 行对象
     * @param data 模板数据
     */
    private void renderRow(Row row, Map<String, Object> data) {
        for (Cell cell : row) {
            renderCell(cell, data);
        }
    }

    /**
     * 渲染单元格。
     *
     * @param cell 单元格
     * @param data 模板数据
     */
    private void renderCell(Cell cell, Map<String, Object> data) {
        if (cell.getCellType() != CellType.STRING) {
            return;
        }
        String template = cell.getStringCellValue();
        if (!containsPlaceholder(template)) {
            return;
        }
        Placeholder single = singlePlaceholder(template);
        if (single != null) {
            ResolvedValue value = resolveValue(single.key, data);
            if (value.found) {
                writeCellValue(cell, value.value);
            }
            return;
        }
        cell.setCellValue(renderText(template, data));
    }

    /**
     * 判断文本是否包含占位符。
     *
     * @param text 文本内容
     * @return 如果包含占位符则返回 true
     */
    private boolean containsPlaceholder(String text) {
        return text != null && PLACEHOLDER_PATTERN.matcher(text).find();
    }

    /**
     * 判断文本是否只包含一个占位符。
     *
     * @param text 文本内容
     * @return 占位符信息，不满足条件时返回 null
     */
    private Placeholder singlePlaceholder(String text) {
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(text.trim());
        if (!matcher.matches()) {
            return null;
        }
        return new Placeholder(matcher.group(0), matcher.group(1));
    }

    /**
     * 渲染文本中的占位符。
     *
     * @param template 模板文本
     * @param data 模板数据
     * @return 渲染后的文本
     */
    private String renderText(String template, Map<String, Object> data) {
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
     * 写入单元格值。
     *
     * @param cell 单元格
     * @param value 单元格值
     */
    private void writeCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue(formatValue(value));
        }
    }

    /**
     * 解析占位符对应的值。
     *
     * @param key 占位符键名
     * @param data 模板数据
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
     * 格式化值对象。
     *
     * @param value 值对象
     * @return 文本值
     */
    private String formatValue(Object value) {
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
    private String joinIterable(Iterable<?> values) {
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

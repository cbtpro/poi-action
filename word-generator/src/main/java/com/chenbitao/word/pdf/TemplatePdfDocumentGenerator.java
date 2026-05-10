package com.chenbitao.word.pdf;

import com.chenbitao.word.exception.PdfDocumentException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.form.PDAcroForm;
import org.apache.pdfbox.pdmodel.interactive.form.PDField;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * PDF 模板文档生成器。
 *
 * <p>基于带 AcroForm 表单域的 PDF 模板文件生成 PDF，字段名使用 {@code key} 或
 * {@code object.field} 形式，并从渲染数据中读取对应值。</p>
 */
public class TemplatePdfDocumentGenerator {

    /** PDF 模板文档对象。 */
    private final PDDocument document;

    /** PDF 表单对象。 */
    private final PDAcroForm form;

    /**
     * 构造 PDF 模板文档生成器。
     *
     * @param template 模板输入流
     * @throws PdfDocumentException 如果模板读取失败或模板不包含表单
     */
    public TemplatePdfDocumentGenerator(InputStream template) {
        try {
            this.document = PDDocument.load(template);
            this.form = document.getDocumentCatalog().getAcroForm();
            if (form == null) {
                throw new PdfDocumentException("PDF 模板未包含 AcroForm 表单域");
            }
            this.form.setNeedAppearances(true);
        } catch (PdfDocumentException e) {
            throw e;
        } catch (Exception e) {
            throw new PdfDocumentException("加载 PDF 模板失败", e);
        }
    }

    /**
     * 渲染 PDF 模板表单域。
     *
     * @param data 模板数据
     */
    public void render(Map<String, Object> data) {
        for (PDField field : form.getFieldTree()) {
            renderField(field, data);
        }
    }

    /**
     * 保存 PDF 到指定路径。
     *
     * @param path 输出文件路径
     */
    public void save(String path) {
        try {
            saveInternal(new File(path));
        } catch (Exception e) {
            throw new PdfDocumentException("保存 PDF 模板文件失败", e);
        } finally {
            closeQuietly();
        }
    }

    /**
     * 保存 PDF 到输出流。
     *
     * @param outputStream 输出流
     */
    public void save(OutputStream outputStream) {
        try {
            flattenForm();
            document.save(outputStream);
        } catch (Exception e) {
            throw new PdfDocumentException("写出 PDF 模板文件失败", e);
        } finally {
            closeQuietly();
        }
    }

    /**
     * 渲染单个表单域。
     *
     * @param field 表单域
     * @param data 模板数据
     */
    private void renderField(PDField field, Map<String, Object> data) {
        String name = field.getFullyQualifiedName();
        ResolvedValue value = resolveValue(name, data);
        if (!value.found) {
            return;
        }
        try {
            field.setValue(formatValue(value.value));
        } catch (Exception e) {
            throw new PdfDocumentException("渲染 PDF 表单域失败：" + name, e);
        }
    }

    /**
     * 保存 PDF 到指定文件。
     *
     * @param file 输出文件
     * @throws Exception 如果表单扁平化或写出失败
     */
    private void saveInternal(File file) throws Exception {
        flattenForm();
        document.save(file);
    }

    /**
     * 将表单域扁平化为普通页面内容。
     *
     * @throws Exception 如果表单扁平化失败
     */
    private void flattenForm() throws Exception {
        form.refreshAppearances();
        form.setNeedAppearances(false);
        form.flatten();
    }

    /**
     * 解析字段对应的模板数据。
     *
     * @param key 字段名
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
     * 按点号分隔路径递归解析数据值。
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
     * 格式化字段值。
     *
     * @param value 字段值
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
     * @param value 字段值
     */
    private void appendValueLine(StringBuilder text, Object value) {
        if (text.length() > 0) {
            text.append('\n');
        }
        text.append(value == null ? "" : formatValue(value));
    }

    /**
     * 安静关闭 PDF 文档。
     */
    private void closeQuietly() {
        try {
            document.close();
        } catch (Exception ignored) {
            // 释放失败不影响调用方看到原始生成或写出结果。
        }
    }

    /**
     * 字段解析结果。
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

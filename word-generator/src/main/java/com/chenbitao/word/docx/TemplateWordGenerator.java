package com.chenbitao.word.docx;

import com.chenbitao.word.exception.WordException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.lang.reflect.Array;
import java.net.URI;
import java.net.URL;
import java.net.URLConnection;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Word模板文档生成器
 * 支持基于模板的Word文档生成，包括文本替换、表格循环填充和图片插入等功能
 * 使用${key}格式的占位符进行文本替换，使用${list.field}格式进行列表循环填充
 */
public class TemplateWordGenerator {

    /** 普通占位符正则表达式模式 */
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    /** 循环占位符正则表达式模式 */
    private static final Pattern LOOP_PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{(\\w+)\\.\\w+}");

    /** Word文档对象 */
    private final XWPFDocument document;

    /**
     * 创建图片对象（默认尺寸）
     *
     * @param source 图片源，可以是InputStream、byte[]、File、Path、URL、URI或Base64字符串
     * @return 图片对象
     */
    public static Picture picture(Object source) {
        return new Picture(source, null, null);
    }

    /**
     * 创建图片对象（指定尺寸）
     *
     * @param source 图片源，可以是InputStream、byte[]、File、Path、URL、URI或Base64字符串
     * @param widthEmu 图片宽度（EMU单位）
     * @param heightEmu 图片高度（EMU单位）
     * @return 图片对象
     */
    public static Picture picture(Object source, int widthEmu, int heightEmu) {
        return new Picture(source, widthEmu, heightEmu);
    }

    /**
     * 图片内部类
     * 封装图片数据和尺寸信息
     */
    public static class Picture {
        /** 图片源数据 */
        private final Object source;
        /** 图片宽度（EMU单位） */
        private final Integer widthEmu;
        /** 图片高度（EMU单位） */
        private final Integer heightEmu;

        /**
         * 构造图片对象
         *
         * @param source 图片源
         * @param widthEmu 宽度
         * @param heightEmu 高度
         */
        private Picture(Object source, Integer widthEmu, Integer heightEmu) {
            this.source = source;
            this.widthEmu = widthEmu;
            this.heightEmu = heightEmu;
        }

        /**
         * 获取图片源数据。
         *
         * @return 图片源
         */
        public Object getSource() {
            return source;
        }

        /**
         * 获取图片宽度。
         *
         * @return 图片宽度（EMU单位），未指定时返回null
         */
        public Integer getWidthEmu() {
            return widthEmu;
        }

        /**
         * 获取图片高度。
         *
         * @return 图片高度（EMU单位），未指定时返回null
         */
        public Integer getHeightEmu() {
            return heightEmu;
        }
    }

    /**
     * 构造模板文档生成器
     *
     * @param template 模板文档输入流
     * @throws WordException 如果加载模板失败
     */
    public TemplateWordGenerator(InputStream template) {
        try {
            this.document = new XWPFDocument(template);
        } catch (Exception e) {
            throw new WordException("加载模板失败", e);
        }
    }

    // ================== 文本替换 ==================
    /**
     * 渲染模板文档
     * 使用提供的数据替换文档中的占位符，包括普通文本替换和表格循环填充
     *
     * @param data 包含替换数据的Map，key为占位符名称，value为替换值
     */
    public void render(Map<String, Object> data) {
        for (XWPFParagraph p : document.getParagraphs()) {
            replaceInParagraph(p, data);
        }

        for (XWPFTable table : document.getTables()) {
            replaceInTable(table, data);   // 普通填充
            processTable(table, data);     // 循环填充
        }
    }

    /**
     * 替换段落中的占位符
     *
     * @param p 要处理的段落
     * @param data 替换数据
     */
    private void replaceInParagraph(XWPFParagraph p, Map<String, Object> data) {
        List<XWPFRun> runs = p.getRuns();
        if (runs == null || runs.isEmpty()) {
            return;
        }

        String text = paragraphText(runs);
        if (!containsPlaceholder(text)) {
            return;
        }

        Picture picture = singleImagePlaceholder(text, data);
        clearRuns(runs);
        XWPFRun firstRun = runs.get(0);
        if (picture != null) {
            addPicture(firstRun, picture);
            return;
        }

        setRunText(firstRun, renderText(text, data));
    }

    /**
     * 替换表格中的占位符
     *
     * @param table 要处理的表格
     * @param data 替换数据
     */
    private void replaceInTable(XWPFTable table, Map<String, Object> data) {

        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {

                for (XWPFParagraph p : cell.getParagraphs()) {
                    replaceInParagraph(p, data); // 复用已有逻辑
                }
            }
        }
    }

    /**
     * 判断行是否为循环行
     *
     * @param row 表格行
     * @return 如果包含循环占位符则返回true
     */
    private boolean isLoopRow(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            String text = cell.getText();
            if (text != null && LOOP_PLACEHOLDER_PATTERN.matcher(text).find()) {
                return true;
            }
        }
        return false;
    }

    /**
     * 从文本中提取列表名称
     *
     * @param text 包含占位符的文本
     * @return 列表名称，如果未找到则返回null
     */
    private String extractListName(String text) {
        Matcher matcher = LOOP_PLACEHOLDER_PATTERN.matcher(text);

        if (matcher.find()) {
            return matcher.group(1); // 列表名，如 exp、experience、user
        }

        return null;
    }

    /**
     * 获取行的完整文本内容
     *
     * @param row 表格行
     * @return 行的文本内容
     */
    private String getRowText(XWPFTableRow row) {
        StringBuilder sb = new StringBuilder();
        for (XWPFTableCell cell : row.getTableCells()) {
            sb.append(cell.getText());
        }
        return sb.toString();
    }

    /**
     * 判断单元格是否为循环单元格
     *
     * @param cell 表格单元格
     * @param listName 列表名称
     * @return 如果单元格包含指定列表的循环占位符则返回true
     */
    private boolean isLoopCell(XWPFTableCell cell, String listName) {
        String text = cell.getText();
        return text != null && Pattern.compile("\\$\\{" + Pattern.quote(listName) + "\\.\\w+}")
                .matcher(text)
                .find();
    }

    /**
     * 判断是否应该在单元格内渲染循环
     *
     * @param row 表格行
     * @param listName 列表名称
     * @return 如果循环单元格数量不超过1则返回true
     */
    private boolean shouldRenderLoopInCell(XWPFTableRow row, String listName) {
        return loopCellCount(row, listName) <= 1;
    }

    /**
     * 计算循环单元格的数量
     *
     * @param row 表格行
     * @param listName 列表名称
     * @return 循环单元格的数量
     */
    private int loopCellCount(XWPFTableRow row, String listName) {
        int count = 0;
        for (XWPFTableCell cell : row.getTableCells()) {
            if (isLoopCell(cell, listName)) {
                count++;
            }
        }
        return count;
    }

    /**
     * 设置单元格文本内容
     *
     * @param cell 表格单元格
     * @param text 要设置的文本内容
     */
    private void setCellText(XWPFTableCell cell, String text) {
        while (!cell.getParagraphs().isEmpty()) {
            cell.removeParagraph(0);
        }

        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        String[] lines = text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                run.addBreak();
            }
            run.setText(lines[i]);
        }
    }

    /**
     * 渲染循环单元格内容
     *
     * @param template 模板文本
     * @param listName 列表名称
     * @param list 数据列表
     * @return 渲染后的文本内容
     */
    private String renderLoopCell(String template, String listName, List<Map<String, Object>> list) {
        StringBuilder sb = new StringBuilder();

        for (int i = 0; i < list.size(); i++) {
            if (i > 0) {
                sb.append('\n');
            }
            sb.append(renderLoopText(template, listName, list.get(i)));
        }

        return sb.toString();
    }

    /**
     * 渲染循环文本
     *
     * @param template 模板文本
     * @param listName 列表名称
     * @param item 单个数据项
     * @return 渲染后的文本
     */
    private String renderLoopText(String template, String listName, Map<String, Object> item) {
        return renderText(template, java.util.Collections.singletonMap(listName, item));
    }

    /**
     * 在单元格内渲染循环数据
     *
     * @param row 表格行
     * @param listName 列表名称
     * @param list 数据列表
     */
    private void renderLoopInCells(XWPFTableRow row, String listName, List<Map<String, Object>> list) {
        for (XWPFTableCell cell : row.getTableCells()) {
            if (isLoopCell(cell, listName)) {
                setCellText(cell, renderLoopCell(cell.getText(), listName, list));
            }
        }
    }

    /**
     * 在行中渲染循环数据
     *
     * @param table 表格
     * @param rowIndex 行索引
     * @param row 模板行
     * @param listName 列表名称
     * @param list 数据列表
     */
    private void renderLoopInRows(XWPFTable table, int rowIndex, XWPFTableRow row,
                                  String listName, List<Map<String, Object>> list) {
        int offset = 0;
        for (Map<String, Object> item : list) {
            XWPFTableRow newRow = table.insertNewTableRow(rowIndex + 1 + offset);

            for (int j = 0; j < row.getTableCells().size(); j++) {
                XWPFTableCell templateCell = row.getCell(j);
                XWPFTableCell newCell = newRow.createCell();
                if (templateCell.getCTTc().isSetTcPr()) {
                    newCell.getCTTc().setTcPr((CTTcPr) templateCell.getCTTc().getTcPr().copy());
                }

                if (isLoopCell(templateCell, listName)) {
                    setCellText(newCell, renderLoopText(templateCell.getText(), listName, item));
                } else {
                    renderStaticCell(templateCell, newCell, offset);
                }
            }

            offset++;
        }

        table.removeRow(rowIndex);
    }

    /**
     * 渲染静态单元格
     *
     * @param templateCell 模板单元格
     * @param newCell 新单元格
     * @param offset 偏移量
     */
    private void renderStaticCell(XWPFTableCell templateCell, XWPFTableCell newCell, int offset) {
        STMerge.Enum templateMerge = getVerticalMerge(templateCell);
        if (STMerge.CONTINUE.equals(templateMerge)) {
            setCellText(newCell, "");
            return;
        }

        String text = offset == 0 ? templateCell.getText() : "";
        if (templateMerge != null || !templateCell.getText().isEmpty()) {
            setVerticalMerge(newCell, offset == 0);
        }
        setCellText(newCell, text);
    }

    /**
     * 获取单元格的垂直合并状态
     *
     * @param cell 表格单元格
     * @return 合并状态枚举值
     */
    private STMerge.Enum getVerticalMerge(XWPFTableCell cell) {
        if (!cell.getCTTc().isSetTcPr()) {
            return null;
        }
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (!tcPr.isSetVMerge()) {
            return null;
        }
        return tcPr.getVMerge().getVal();
    }

    /**
     * 设置单元格的垂直合并状态
     *
     * @param cell 表格单元格
     * @param restart 是否重新开始合并
     */
    private void setVerticalMerge(XWPFTableCell cell, boolean restart) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr()
                ? cell.getCTTc().getTcPr()
                : cell.getCTTc().addNewTcPr();
        CTVMerge vMerge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
        vMerge.setVal(restart ? STMerge.RESTART : STMerge.CONTINUE);
    }

    /**
     * 处理表格中的循环行
     *
     * @param table 表格
     * @param data 数据Map
     */
    private void processTable(XWPFTable table, Map<String, Object> data) {
        for (int i = 0; i < table.getRows().size(); i++) {
            XWPFTableRow row = table.getRow(i);
            processLoopRowIfPresent(table, i, row, data);
        }
    }

    /**
     * 处理循环行（包含 ${list.field} 格式的占位符）
     *
     * @return true 如果此行是循环行并已处理，false 否则
     */
    private boolean processLoopRowIfPresent(XWPFTable table, int rowIndex,
                                            XWPFTableRow row,
                                            Map<String, Object> data) {
        if (!isLoopRow(row)) {
            return false;
        }

        String listName = extractListName(getRowText(row));
        List<Map<String, Object>> listData = extractListData(listName, data);

        if (listData == null || listData.isEmpty()) {
            return false;
        }

        renderLoopRow(table, rowIndex, row, listName, listData);
        return true;
    }

    /**
     * 从数据map中提取列表数据，进行类型安全的转换
     */
    private List<Map<String, Object>> extractListData(String listName,
                                                       Map<String, Object> data) {
        if (listName == null) {
            return Collections.emptyList();
        }

        Object value = data.get(listName);
        if (value == null) {
            return Collections.emptyList();
        }

        if (!(value instanceof List)) {
            return Collections.emptyList();
        }

        return toListOfMaps((List<?>) value, listName);
    }

    private List<Map<String, Object>> toListOfMaps(List<?> values, String listName) {
        List<Map<String, Object>> result = new ArrayList<>(values.size());
        for (Object item : values) {
            result.add(toStringKeyMap(item, listName));
        }
        return result;
    }

    private Map<String, Object> toStringKeyMap(Object item, String listName) {
        if (!(item instanceof Map)) {
            throw new WordException("循环数据必须是Map类型：" + listName, null);
        }
        Map<?, ?> source = (Map<?, ?>) item;
        Map<String, Object> result = new LinkedHashMap<>(source.size());
        for (Map.Entry<?, ?> entry : source.entrySet()) {
            Object key = entry.getKey();
            if (!(key instanceof String)) {
                throw new WordException("循环数据的字段名必须是String类型：" + listName, null);
            }
            result.put((String) key, entry.getValue());
        }
        return result;
    }

    /**
     * 处理循环行的渲染逻辑
     */
    private void renderLoopRow(XWPFTable table, int rowIndex, XWPFTableRow row,
                               String listName, List<Map<String, Object>> listData) {
        if (shouldRenderLoopInCell(row, listName)) {
            renderLoopInCells(row, listName, listData);
        } else {
            renderLoopInRows(table, rowIndex, row, listName, listData);
        }
    }

    /**
     * 检查文本是否包含占位符
     *
     * @param text 要检查的文本
     * @return 如果包含占位符则返回true
     */
    private boolean containsPlaceholder(String text) {
        return text != null && PLACEHOLDER_PATTERN.matcher(text).find();
    }

    /**
     * 获取段落的所有文本内容
     *
     * @param runs 段落的文本运行列表
     * @return 段落的完整文本内容
     */
    private String paragraphText(List<XWPFRun> runs) {
        StringBuilder text = new StringBuilder();
        for (XWPFRun run : runs) {
            String runText = run.getText(0);
            if (runText != null) {
                text.append(runText);
            }
            for (int i = 0; i < run.getCTR().sizeOfBrArray(); i++) {
                text.append('\n');
            }
        }
        return text.toString();
    }

    /**
     * 渲染文本中的占位符
     *
     * @param template 包含占位符的模板文本
     * @param data 数据Map
     * @return 渲染后的文本
     */
    private String renderText(String template, Map<String, Object> data) {
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(template);
        StringBuffer result = new StringBuffer();
        while (matcher.find()) {
            String key = matcher.group(1);
            Object value = resolveValue(key, data);
            String replacement = value == null && !containsValue(key, data) ? matcher.group(0) : formatValue(value);
            matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(result);
        return result.toString();
    }

    /**
     * 格式化值对象为字符串
     *
     * @param value 要格式化的值
     * @return 格式化后的字符串
     */
    private String formatValue(Object value) {
        if (value == null) {
            return "";
        }

        if (value instanceof Iterable) {
            return joinValues((Iterable<?>) value);
        }

        if (value.getClass().isArray()) {
            return joinArray(value);
        }

        return String.valueOf(value);
    }

    private String joinArray(Object array) {
        StringBuilder text = new StringBuilder();
        int length = Array.getLength(array);
        for (int i = 0; i < length; i++) {
            appendValueLine(text, Array.get(array, i));
        }
        return text.toString();
    }

    private String joinValues(Iterable<?> values) {
        StringBuilder text = new StringBuilder();
        for (Object value : values) {
            appendValueLine(text, value);
        }
        return text.toString();
    }

    private void appendValueLine(StringBuilder text, Object value) {
        if (text.length() > 0) {
            text.append('\n');
        }
        text.append(value == null ? "" : value);
    }

    /**
     * 解析占位符的值
     * 支持点号分隔的嵌套属性访问，如"user.name"
     *
     * @param key 占位符键
     * @param data 数据Map
     * @return 解析得到的值，如果不存在则返回null
     */
    @SuppressWarnings("unchecked")
    private Object resolveValue(String key, Map<String, Object> data) {
        if (data.containsKey(key)) {
            return data.get(key);
        }

        String[] parts = key.split("\\.");
        Object current = data;
        for (String part : parts) {
            if (!(current instanceof Map)) {
                return null;
            }
            current = ((Map<String, Object>) current).get(part);
            if (current == null) {
                return null;
            }
        }
        return current;
    }

    /**
     * 检查数据中是否存在指定的键
     *
     * @param key 要检查的键
     * @param data 数据Map
     * @return 如果存在则返回true
     */
    @SuppressWarnings("unchecked")
    private boolean containsValue(String key, Map<String, Object> data) {
        if (data.containsKey(key)) {
            return true;
        }

        String[] parts = key.split("\\.");
        Object current = data;
        for (String part : parts) {
            if (!(current instanceof Map)) {
                return false;
            }
            Map<String, Object> currentMap = (Map<String, Object>) current;
            if (!currentMap.containsKey(part)) {
                return false;
            }
            current = currentMap.get(part);
        }
        return true;
    }

    /**
     * 检查单个图片占位符
     *
     * @param text 文本内容
     * @param data 数据Map
     * @return 如果是图片占位符则返回Picture对象，否则返回null
     */
    private Picture singleImagePlaceholder(String text, Map<String, Object> data) {
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(text.trim());
        if (!matcher.matches()) {
            return null;
        }

        String key = matcher.group(1);
        Object value = resolveValue(key, data);
        if (value instanceof Picture) {
            return (Picture) value;
        }
        if (isPictureValue(value)) {
            return picture(value);
        }
        return null;
    }

    /**
     * 判断对象是否为图片值
     *
     * @param value 要检查的对象
     * @return 如果是图片类型则返回true
     */
    private boolean isPictureValue(Object value) {
        if (value instanceof InputStream || value instanceof byte[]
                || value instanceof File || value instanceof Path
                || value instanceof URL || value instanceof URI) {
            return true;
        }

        if (!(value instanceof String)) {
            return false;
        }

        String text = ((String) value).trim();
        if (text.isEmpty()) {
            return false;
        }
        if (text.startsWith("data:image/")
                || text.startsWith("http://")
                || text.startsWith("https://")
                || isBase64Image(text)) {
            return true;
        }

        if (containsInvalidPathChars(text)) {
            return false;
        }

        return Files.isRegularFile(new File(text).toPath());
    }
    /**
     * 检查字符串是否包含非法的文件路径字符
     *
     * @param path 要检查的路径字符串
     * @return 如果包含非法字符则返回true
     */
    private boolean containsInvalidPathChars(String path) {
        String osName = System.getProperty("os.name").toLowerCase();
        if (osName.contains("win")) {
            return path.matches(".*[<>:\"/\\\\|?*].*");
        } else {
            return path.contains("\0");
        }
    }

    /**
     * 清空文本运行的内容
     *
     * @param runs 要清空的文本运行列表
     */
    private void clearRuns(List<XWPFRun> runs) {
        for (XWPFRun run : runs) {
            run.setText("", 0);
            run.getCTR().setBrArray(new org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBr[0]);
        }
    }

    /**
     * 设置文本运行的内容
     *
     * @param run 文本运行对象
     * @param text 要设置的文本内容
     */
    private void setRunText(XWPFRun run, String text) {
        String[] lines = text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                run.addBreak();
                run.setText(lines[i]);
            } else {
                run.setText(lines[i], 0);
            }
        }
    }

    /**
     * 检查字符串是否为Base64编码的图片
     *
     * @param text 要检查的字符串
     * @return 如果是Base64图片则返回true
     */
    private boolean isBase64Image(String text) {
        try {
            byte[] bytes = Base64.getDecoder().decode(text);
            return detectPictureType(bytes) != null || ImageIO.read(new ByteArrayInputStream(bytes)) != null;
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * 向文本运行中添加图片
     *
     * @param run 文本运行对象
     * @param picture 图片对象
     * @throws WordException 如果图片添加失败
     */
    private void addPicture(XWPFRun run, Picture picture) {
        try {
            PictureData pictureData = readPictureData(picture.source);
            int[] size = pictureSize(picture, pictureData);
            run.addPicture(new ByteArrayInputStream(pictureData.bytes),
                    pictureData.pictureType,
                    pictureData.fileName,
                    size[0],
                    size[1]);
        } catch (Exception e) {
            throw new WordException("图片替换失败", e);
        }
    }

    /**
     * 读取图片数据
     *
     * @param source 图片源
     * @return 图片数据对象
     * @throws Exception 如果读取失败
     */
    private PictureData readPictureData(Object source) throws Exception {
        byte[] bytes = readPictureBytes(source);
        Integer pictureType = detectPictureType(bytes);
        BufferedImage image = ImageIO.read(new ByteArrayInputStream(bytes));
        if (pictureType != null) {
            return new PictureData(bytes, pictureType, fileName(source, pictureType), imageWidthEmu(image), imageHeightEmu(image));
        }

        if (image == null) {
            throw new WordException("不支持的图片格式", null);
        }

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return new PictureData(output.toByteArray(), Document.PICTURE_TYPE_PNG, fileName(source, Document.PICTURE_TYPE_PNG),
                imageWidthEmu(image), imageHeightEmu(image));
    }

    /**
     * 获取图片尺寸
     *
     * @param picture 图片对象
     * @param pictureData 图片数据
     * @return 包含宽度和高度的数组
     * @throws WordException 如果无法获取尺寸
     */
    private int[] pictureSize(Picture picture, PictureData pictureData) {
        Integer width = picture.widthEmu == null ? pictureData.widthEmu : picture.widthEmu;
        Integer height = picture.heightEmu == null ? pictureData.heightEmu : picture.heightEmu;
        if (width == null || height == null) {
            throw new WordException("无法识别图片尺寸，请在业务数据中指定图片宽高", null);
        }
        return new int[]{width, height};
    }

    /**
     * 获取图片宽度（EMU单位）
     *
     * @param image 图片对象
     * @return 宽度值，如果图片为null则返回null
     */
    private Integer imageWidthEmu(BufferedImage image) {
        return image == null ? null : Units.pixelToEMU(image.getWidth());
    }

    /**
     * 获取图片高度（EMU单位）
     *
     * @param image 图片对象
     * @return 高度值，如果图片为null则返回null
     */
    private Integer imageHeightEmu(BufferedImage image) {
        return image == null ? null : Units.pixelToEMU(image.getHeight());
    }

    /**
     * 读取图片字节数据
     *
     * @param source 图片源
     * @return 图片字节数组
     * @throws Exception 如果读取失败
     */
    private byte[] readPictureBytes(Object source) throws Exception {
        if (source instanceof Picture) {
            return readPictureBytes(((Picture) source).source);
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
            return readFromUrl((URL) source);
        }
        if (source instanceof URI) {
            return readFromUrl(((URI) source).toURL());
        }
        if (source instanceof String) {
            return readPictureBytes((String) source);
        }

        throw new WordException("不支持的图片输入类型：" + source.getClass().getName(), null);
    }

    /**
     * 从字符串读取图片字节数据
     *
     * @param source 字符串形式的图片源
     * @return 图片字节数组
     * @throws IOException 如果读取失败
     */
    private byte[] readPictureBytes(String source) throws IOException {
        String text = source.trim();
        if (text.startsWith("data:image/")) {
            int commaIndex = text.indexOf(',');
            if (commaIndex < 0) {
                throw new WordException("无效的 base64 图片", null);
            }
            return Base64.getDecoder().decode(text.substring(commaIndex + 1));
        }
        if (text.startsWith("http://") || text.startsWith("https://")) {
            return readFromUrl(new URL(text));
        }

        File file = new File(text);
        if (file.exists()) {
            return Files.readAllBytes(file.toPath());
        }

        return Base64.getDecoder().decode(text);
    }

    /**
     * 从URL读取图片数据
     *
     * @param url 图片URL
     * @return 图片字节数组
     * @throws IOException 如果读取失败
     */
    private byte[] readFromUrl(URL url) throws IOException {
        URLConnection connection = url.openConnection();
        connection.setConnectTimeout(10000);
        connection.setReadTimeout(10000);
        try (InputStream inputStream = connection.getInputStream()) {
            return readAll(inputStream);
        }
    }

    /**
     * 读取输入流的所有数据
     *
     * @param inputStream 输入流
     * @return 字节数组
     * @throws IOException 如果读取失败
     */
    private byte[] readAll(InputStream inputStream) throws IOException {
        try (InputStream in = inputStream;
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[8192];
            int length;
            while ((length = in.read(buffer)) != -1) {
                output.write(buffer, 0, length);
            }
            return output.toByteArray();
        }
    }

    /**
     * 检测图片类型
     *
     * @param bytes 图片字节数据
     * @return 图片类型常量，如果无法识别则返回null
     */
    private Integer detectPictureType(byte[] bytes) {
        if (startsWith(bytes, new byte[]{(byte) 0x89, 'P', 'N', 'G'})) {
            return Document.PICTURE_TYPE_PNG;
        }
        if (startsWith(bytes, new byte[]{(byte) 0xFF, (byte) 0xD8})) {
            return Document.PICTURE_TYPE_JPEG;
        }
        if (startsWith(bytes, new byte[]{'G', 'I', 'F'})) {
            return Document.PICTURE_TYPE_GIF;
        }
        if (startsWith(bytes, new byte[]{'B', 'M'})) {
            return Document.PICTURE_TYPE_BMP;
        }
        if (startsWith(bytes, new byte[]{'I', 'I', 42, 0})
                || startsWith(bytes, new byte[]{'M', 'M', 0, 42})) {
            return Document.PICTURE_TYPE_TIFF;
        }
        return null;
    }

    /**
     * 检查字节数组是否以指定前缀开头
     *
     * @param bytes 字节数组
     * @param prefix 前缀字节数组
     * @return 如果匹配则返回true
     */
    private boolean startsWith(byte[] bytes, byte[] prefix) {
        if (bytes == null || bytes.length < prefix.length) {
            return false;
        }
        for (int i = 0; i < prefix.length; i++) {
            if (bytes[i] != prefix[i]) {
                return false;
            }
        }
        return true;
    }

    /**
     * 生成图片文件名
     *
     * @param source 图片源
     * @param pictureType 图片类型
     * @return 文件名
     */
    private String fileName(Object source, int pictureType) {
        String extension = pictureExtension(pictureType);
        if (source instanceof File) {
            return ((File) source).getName();
        }
        if (source instanceof Path) {
            Path fileName = ((Path) source).getFileName();
            return fileName == null ? "image." + extension : fileName.toString();
        }
        return "image." + extension;
    }

    /**
     * 获取图片类型的文件扩展名
     *
     * @param pictureType 图片类型常量
     * @return 文件扩展名
     */
    private String pictureExtension(int pictureType) {
        switch (pictureType) {
            case Document.PICTURE_TYPE_JPEG:
                return "jpg";
            case Document.PICTURE_TYPE_GIF:
                return "gif";
            case Document.PICTURE_TYPE_BMP:
                return "bmp";
            case Document.PICTURE_TYPE_TIFF:
                return "tiff";
            case Document.PICTURE_TYPE_PNG:
            default:
                return "png";
        }
    }

    /**
     * 图片数据内部类
     * 封装图片的字节数据、类型、文件名和尺寸信息
     */
    private static class PictureData {
        /** 图片字节数据 */
        private final byte[] bytes;
        /** POI图片类型 */
        private final int pictureType;
        /** 文件名 */
        private final String fileName;
        /** 宽度（EMU单位） */
        private final Integer widthEmu;
        /** 高度（EMU单位） */
        private final Integer heightEmu;

        /**
         * 构造图片数据对象
         *
         * @param bytes 字节数据
         * @param pictureType 图片类型
         * @param fileName 文件名
         * @param widthEmu 宽度
         * @param heightEmu 高度
         */
        private PictureData(byte[] bytes, int pictureType, String fileName, Integer widthEmu, Integer heightEmu) {
            this.bytes = bytes;
            this.pictureType = pictureType;
            this.fileName = fileName;
            this.widthEmu = widthEmu;
            this.heightEmu = heightEmu;
        }
    }

    /**
     * 保存文档到输出流
     *
     * @param out 输出流
     * @throws WordException 如果保存失败
     */
    public void save(OutputStream out) {
        try {
            document.write(out);
        } catch (Exception e) {
            throw new WordException("导出失败", e);
        }
    }

    /**
     * 保存文档到指定路径
     *
     * @param path 文件路径
     * @throws WordException 如果保存失败
     */
    public void save(String path) {
        try (FileOutputStream out = new FileOutputStream(path)) {
            save(out);
        } catch (Exception e) {
            throw new WordException("导出失败", e);
        }
    }
}

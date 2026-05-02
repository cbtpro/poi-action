package com.chenbitao.word.docx;

import com.chenbitao.word.exception.WordException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TemplateWordGenerator {

    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");
    private static final Pattern LOOP_PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{(\\w+)\\.\\w+}");

    private XWPFDocument document;

    public TemplateWordGenerator(InputStream template) {
        try {
            this.document = new XWPFDocument(template);
        } catch (Exception e) {
            throw new WordException("加载模板失败", e);
        }
    }

    // ================== 文本替换 ==================
    public void render(Map<String, Object> data) {
        for (XWPFParagraph p : document.getParagraphs()) {
            replaceInParagraph(p, data);
        }

        for (XWPFTable table : document.getTables()) {
            replaceInTable(table, data);   // 普通填充
            processTable(table, data);     // 循环填充
        }
    }

    private void replaceInParagraph(XWPFParagraph p, Map<String, Object> data) {
        List<XWPFRun> runs = p.getRuns();
        if (runs == null || runs.isEmpty()) {
            return;
        }

        String text = paragraphText(runs);
        if (!containsPlaceholder(text)) {
            return;
        }

        String imageKey = singleImagePlaceholder(text, data);
        clearRuns(runs);
        XWPFRun firstRun = runs.get(0);
        if (imageKey != null) {
            addPicture(firstRun, imageKey, (InputStream) resolveValue(imageKey, data));
            return;
        }

        setRunText(firstRun, renderText(text, data));
    }

    private void replaceInTable(XWPFTable table, Map<String, Object> data) {

        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {

                for (XWPFParagraph p : cell.getParagraphs()) {
                    replaceInParagraph(p, data); // 复用已有逻辑
                }
            }
        }
    }
    private boolean isLoopRow(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            String text = cell.getText();
            if (text != null && LOOP_PLACEHOLDER_PATTERN.matcher(text).find()) {
                return true;
            }
        }
        return false;
    }
    private String extractListName(String text) {
        Matcher matcher = LOOP_PLACEHOLDER_PATTERN.matcher(text);

        if (matcher.find()) {
            return matcher.group(1); // 列表名，如 exp、experience、user
        }

        return null;
    }
    private String getRowText(XWPFTableRow row) {
        StringBuilder sb = new StringBuilder();
        for (XWPFTableCell cell : row.getTableCells()) {
            sb.append(cell.getText());
        }
        return sb.toString();
    }

    private boolean isLoopCell(XWPFTableCell cell, String listName) {
        String text = cell.getText();
        return text != null && Pattern.compile("\\$\\{" + Pattern.quote(listName) + "\\.\\w+}")
                .matcher(text)
                .find();
    }

    private boolean shouldRenderLoopInCell(XWPFTableRow row, String listName) {
        return loopCellCount(row, listName) <= 1;
    }

    private int loopCellCount(XWPFTableRow row, String listName) {
        int count = 0;
        for (XWPFTableCell cell : row.getTableCells()) {
            if (isLoopCell(cell, listName)) {
                count++;
            }
        }
        return count;
    }

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

    private String renderLoopText(String template, String listName, Map<String, Object> item) {
        return renderText(template, java.util.Collections.<String, Object>singletonMap(listName, item));
    }

    private void renderLoopInCells(XWPFTableRow row, String listName, List<Map<String, Object>> list) {
        for (XWPFTableCell cell : row.getTableCells()) {
            if (isLoopCell(cell, listName)) {
                setCellText(cell, renderLoopCell(cell.getText(), listName, list));
            }
        }
    }

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

    private void renderStaticCell(XWPFTableCell templateCell, XWPFTableCell newCell, int offset) {
        STMerge.Enum templateMerge = getVerticalMerge(templateCell);
        if (STMerge.CONTINUE.equals(templateMerge)) {
            setCellText(newCell, "");
            return;
        }

        String text = offset == 0 ? templateCell.getText() : "";
        if (templateMerge != null || templateCell.getText().length() > 0) {
            setVerticalMerge(newCell, offset == 0);
        }
        setCellText(newCell, text);
    }

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

    private void setVerticalMerge(XWPFTableCell cell, boolean restart) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr()
                ? cell.getCTTc().getTcPr()
                : cell.getCTTc().addNewTcPr();
        CTVMerge vMerge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
        vMerge.setVal(restart ? STMerge.RESTART : STMerge.CONTINUE);
    }

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
            return null;
        }

        Object value = data.get(listName);
        if (value == null) {
            return null;
        }

        if (!(value instanceof List)) {
            return null;
        }

        return safeCastToListOfMaps(value);
    }

    private List<Map<String, Object>> safeCastToListOfMaps(Object value) {
        return (List<Map<String, Object>>) value;
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

    private boolean containsPlaceholder(String text) {
        return text != null && PLACEHOLDER_PATTERN.matcher(text).find();
    }

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

    private String renderText(String template, Map<String, Object> data) {
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(template);
        StringBuffer result = new StringBuffer();
        while (matcher.find()) {
            String key = matcher.group(1);
            Object value = resolveValue(key, data);
            String replacement = value == null ? matcher.group(0) : formatValue(value);
            matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(result);
        return result.toString();
    }

    private String formatValue(Object value) {
        if (value instanceof Iterable) {
            StringBuilder text = new StringBuilder();
            for (Object item : (Iterable<?>) value) {
                if (text.length() > 0) {
                    text.append('\n');
                }
                text.append(item == null ? "" : item);
            }
            return text.toString();
        }

        if (value != null && value.getClass().isArray()) {
            StringBuilder text = new StringBuilder();
            int length = Array.getLength(value);
            for (int i = 0; i < length; i++) {
                if (i > 0) {
                    text.append('\n');
                }
                Object item = Array.get(value, i);
                text.append(item == null ? "" : item);
            }
            return text.toString();
        }

        return String.valueOf(value);
    }

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

    private String singleImagePlaceholder(String text, Map<String, Object> data) {
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(text.trim());
        if (!matcher.matches()) {
            return null;
        }

        String key = matcher.group(1);
        Object value = resolveValue(key, data);
        return value instanceof InputStream ? key : null;
    }

    private void clearRuns(List<XWPFRun> runs) {
        for (XWPFRun run : runs) {
            run.setText("", 0);
            run.getCTR().setBrArray(new org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBr[0]);
        }
    }

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

    private void addPicture(XWPFRun run, String key, InputStream inputStream) {
        try {
            run.addPicture(inputStream,
                    Document.PICTURE_TYPE_PNG,
                    key,
                    Units.toEMU(80),
                    Units.toEMU(80));
        } catch (Exception e) {
            throw new WordException("图片替换失败", e);
        }
    }

    public void save(OutputStream out) {
        try {
            document.write(out);
        } catch (Exception e) {
            throw new WordException("导出失败", e);
        }
    }

    public void save(String path) {
        try (FileOutputStream out = new FileOutputStream(path)) {
            save(out);
        } catch (Exception e) {
            throw new WordException("导出失败", e);
        }
    }
}

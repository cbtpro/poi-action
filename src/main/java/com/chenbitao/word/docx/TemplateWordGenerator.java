package com.chenbitao.word.docx;

import com.chenbitao.word.exception.WordException;
import org.apache.poi.xwpf.usermodel.*;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public class TemplateWordGenerator {

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
        for (XWPFRun run : p.getRuns()) {
            String text = run.getText(0);
            if (text == null) continue;

            for (String key : data.keySet()) {
                String placeholder = "${" + key + "}";
                if (text.contains(placeholder)) {

                    Object value = data.get(key);

                    // 图片处理
                    if (value instanceof InputStream) {
                        try {
                            run.setText("", 0);
                            run.addPicture((InputStream) value,
                                    Document.PICTURE_TYPE_PNG,
                                    key,
                                    200,
                                    200);
                        } catch (Exception e) {
                            throw new WordException("图片替换失败", e);
                        }
                    } else {
                        text = text.replace(placeholder, value.toString());
                        run.setText(text, 0);
                    }
                }
            }
        }
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
    private void replaceImage(XWPFParagraph paragraph, Map<String, Object> data) {

        for (XWPFRun run : paragraph.getRuns()) {

            String text = run.getText(0);
            if (text == null) continue;

            for (String key : data.keySet()) {

                if (text.contains("${" + key + "}")) {

                    Object value = data.get(key);

                    if (value instanceof InputStream) {
                        try {
                            run.setText("", 0);
                            run.addPicture(
                                    (InputStream) value,
                                    XWPFDocument.PICTURE_TYPE_PNG,
                                    key,
                                    200,
                                    200
                            );
                        } catch (Exception e) {
                            throw new RuntimeException(e);
                        }
                    }
                }
            }
        }
    }
    private boolean isLoopRow(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            String text = cell.getText();
            if (text != null && text.matches(".*\\$\\{\\w+\\.\\w+}.*")) {
                return true;
            }
        }
        return false;
    }
    private String extractListName(String text) {
        java.util.regex.Pattern pattern =
                java.util.regex.Pattern.compile("\\$\\{(\\w+)\\.\\w+}");

        java.util.regex.Matcher matcher = pattern.matcher(text);

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
        return text != null && text.matches(".*\\$\\{" + listName + "\\.\\w+}.*");
    }

    private boolean shouldRenderLoopInCell(XWPFTableRow row, String listName) {
        for (XWPFTableCell cell : row.getTableCells()) {
            if (!isLoopCell(cell, listName)) {
                return true;
            }
        }
        return false;
    }

    private void setCellText(XWPFTableCell cell, String text) {
        while (cell.getParagraphs().size() > 0) {
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

    private String renderLoopCell(String template, String listName, List<Map<String, String>> list) {
        StringBuilder sb = new StringBuilder();

        for (int i = 0; i < list.size(); i++) {
            String cellText = template;
            Map<String, String> item = list.get(i);

            for (String key : item.keySet()) {
                cellText = cellText.replace(
                        "${" + listName + "." + key + "}",
                        item.get(key)
                );
            }

            if (i > 0) {
                sb.append('\n');
            }
            sb.append(cellText);
        }

        return sb.toString();
    }

    private String renderLoopText(String template, String listName, Map<String, String> item) {
        String text = template;
        for (String key : item.keySet()) {
            text = text.replace(
                    "${" + listName + "." + key + "}",
                    item.get(key)
            );
        }
        return text;
    }

    private void renderLoopInCells(XWPFTableRow row, String listName, List<Map<String, String>> list) {
        for (XWPFTableCell cell : row.getTableCells()) {
            if (isLoopCell(cell, listName)) {
                setCellText(cell, renderLoopCell(cell.getText(), listName, list));
            }
        }
    }

    private void renderLoopInRows(XWPFTable table, int rowIndex, XWPFTableRow row,
                                  String listName, List<Map<String, String>> list) {
        int offset = 0;
        for (Map<String, String> item : list) {
            XWPFTableRow newRow = table.insertNewTableRow(rowIndex + 1 + offset);

            for (int j = 0; j < row.getTableCells().size(); j++) {
                newRow.createCell();
            }

            for (int j = 0; j < row.getTableCells().size(); j++) {
                newRow.getCell(j).setText(renderLoopText(row.getCell(j).getText(), listName, item));
            }

            offset++;
        }

        table.removeRow(rowIndex);
    }

    private void processTable(XWPFTable table, Map<String, Object> data) {

        List<XWPFTableRow> rows = table.getRows();

        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);

            String rowText = getRowText(row);

            // 判断是否是循环行
            if (isLoopRow(row)) {

                String listName = extractListName(rowText);

                List<Map<String, String>> list =
                        (List<Map<String, String>>) data.get(listName);

                if (shouldRenderLoopInCell(row, listName)) {
                    renderLoopInCells(row, listName, list);
                } else {
                    renderLoopInRows(table, i, row, listName, list);
                }

                break;
            }
        }
    }

    public void save(OutputStream out) {
        try {
            document.write(out);
        } catch (Exception e) {
            throw new WordException("导出失败", e);
        }
    }
}

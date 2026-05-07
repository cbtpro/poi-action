package com.chenbitao.word.presentation;

import com.chenbitao.word.exception.PresentationException;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.sl.usermodel.ShapeType;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * PowerPoint Open XML 演示文稿生成器。
 *
 * <p>基于 Apache POI XSLF 生成 {@code .pptx} 文件，支持标题页、文本页、表格页和图片页。</p>
 */
public class XslfPresentationGenerator implements PresentationGenerator {

    /** 默认页面尺寸：宽屏 16:9 */
    private static final Dimension PAGE_SIZE = new Dimension(960, 540);

    /** 默认字体 */
    private static final String FONT = "微软雅黑";

    /** XSLF 演示文稿对象 */
    private XMLSlideShow slideShow;

    /**
     * 创建新的 PPTX 演示文稿。
     */
    @Override
    public void createPresentation() {
        slideShow = new XMLSlideShow();
        slideShow.setPageSize(PAGE_SIZE);
    }

    /**
     * 添加封面页。
     *
     * @param title 标题
     * @param subtitle 副标题
     */
    @Override
    public void addTitleSlide(String title, String subtitle) {
        XSLFSlide slide = createSlide();
        addTextBox(slide, title, 100, 150, 760, 70, 36D, true, new Color(31, 78, 121));
        addTextBox(slide, subtitle, 120, 235, 720, 45, 20D, false, new Color(89, 89, 89));
    }

    /**
     * 添加文本页。
     *
     * @param title 标题
     * @param bulletItems 项目符号文本
     */
    @Override
    public void addTextSlide(String title, List<String> bulletItems) {
        XSLFSlide slide = createSlide();
        addSlideTitle(slide, title);

        XSLFTextBox content = addTextBox(slide, bulletText(bulletItems), 120, 135, 720, 300, 20D, false, Color.DARK_GRAY);
        content.setLeftInset(8D);
        content.setTopInset(8D);
    }

    /**
     * 添加表格页。
     *
     * @param title 标题
     * @param rows 表格数据，每个内部列表代表一行
     */
    @Override
    public void addTableSlide(String title, List<List<String>> rows) {
        XSLFSlide slide = createSlide();
        addSlideTitle(slide, title);
        if (rows == null || rows.isEmpty()) {
            return;
        }

        int rowCount = rows.size();
        int columnCount = maxColumnCount(rows);
        XSLFTable table = slide.createTable(rowCount, columnCount);
        table.setAnchor(new Rectangle2D.Double(90, 135, 780, Math.max(90, rowCount * 36)));

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            List<String> rowData = rows.get(rowIndex);
            table.setRowHeight(rowIndex, 36D);
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                XSLFTableCell cell = table.getCell(rowIndex, columnIndex);
                String value = rowData != null && columnIndex < rowData.size() ? rowData.get(columnIndex) : "";
                styleTableCell(cell, value, rowIndex == 0);
            }
        }

        for (int i = 0; i < columnCount; i++) {
            table.setColumnWidth(i, 780D / columnCount);
        }
    }

    /**
     * 添加图片页。
     *
     * @param title 标题
     * @param image 图片输入流
     */
    @Override
    public void addImageSlide(String title, InputStream image) {
        XSLFSlide slide = createSlide();
        addSlideTitle(slide, title);
        try {
            XSLFPictureData pictureData = slideShow.addPicture(readAllBytes(image), PictureData.PictureType.PNG);
            XSLFPictureShape picture = slide.createPicture(pictureData);
            picture.setAnchor(new Rectangle2D.Double(330, 145, 300, 240));
        } catch (Exception e) {
            throw new PresentationException("添加 PPTX 图片失败", e);
        }
    }

    /**
     * 保存演示文稿到文件。
     *
     * @param path 输出文件路径
     */
    @Override
    public void save(String path) {
        ensurePresentation();
        try (FileOutputStream output = new FileOutputStream(path)) {
            save(output);
        } catch (Exception e) {
            throw new PresentationException("保存 PPTX 失败", e);
        }
    }

    /**
     * 保存演示文稿到输出流。
     *
     * @param outputStream 输出流
     */
    @Override
    public void save(OutputStream outputStream) {
        ensurePresentation();
        try {
            slideShow.write(outputStream);
        } catch (Exception e) {
            throw new PresentationException("写出 PPTX 失败", e);
        }
    }

    private XSLFSlide createSlide() {
        ensurePresentation();
        return slideShow.createSlide();
    }

    private void addSlideTitle(XSLFSlide slide, String title) {
        addTextBox(slide, title, 60, 35, 840, 55, 28D, true, new Color(31, 78, 121));
        XSLFAutoShape line = slide.createAutoShape();
        line.setShapeType(ShapeType.RECT);
        line.setAnchor(new Rectangle2D.Double(60, 95, 840, 3));
        line.setFillColor(new Color(91, 155, 213));
        line.setLineColor(new Color(91, 155, 213));
    }

    private XSLFTextBox addTextBox(XSLFSlide slide, String text,
                                   double x, double y, double width, double height,
                                   Double fontSize, boolean bold, Color color) {
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(x, y, width, height));
        textBox.setText(text == null ? "" : text);
        textBox.setFillColor(null);
        textBox.setLineColor(null);
        textBox.setWordWrap(true);
        styleText(textBox, fontSize, bold, color);
        return textBox;
    }

    private void styleText(XSLFTextBox textBox, Double fontSize, boolean bold, Color color) {
        for (XSLFTextRun run : textBox.getTextParagraphs().get(0).getTextRuns()) {
            run.setFontFamily(FONT);
            run.setFontSize(fontSize);
            run.setBold(bold);
            run.setFontColor(color);
        }
    }

    private String bulletText(List<String> bulletItems) {
        if (bulletItems == null || bulletItems.isEmpty()) {
            return "";
        }

        StringBuilder text = new StringBuilder();
        for (String item : bulletItems) {
            if (text.length() > 0) {
                text.append('\n');
            }
            text.append("• ").append(item == null ? "" : item);
        }
        return text.toString();
    }

    private int maxColumnCount(List<List<String>> rows) {
        int max = 1;
        for (List<String> row : rows) {
            if (row != null && row.size() > max) {
                max = row.size();
            }
        }
        return max;
    }

    private void styleTableCell(XSLFTableCell cell, String text, boolean header) {
        cell.setText(text == null ? "" : text);
        cell.setFillColor(header ? new Color(91, 155, 213) : Color.WHITE);
        cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
        cell.setTopInset(4D);
        cell.setLeftInset(6D);
        cell.setRightInset(6D);
        setBorders(cell, new Color(191, 191, 191));
        for (XSLFTextRun run : cell.getTextParagraphs().get(0).getTextRuns()) {
            run.setFontFamily(FONT);
            run.setFontSize(header ? 14D : 12D);
            run.setBold(header);
            run.setFontColor(header ? Color.WHITE : Color.DARK_GRAY);
        }
        cell.getTextParagraphs().get(0).setTextAlign(TextParagraph.TextAlign.CENTER);
    }

    private void setBorders(XSLFTableCell cell, Color color) {
        for (TableCell.BorderEdge edge : TableCell.BorderEdge.values()) {
            cell.setBorderColor(edge, color);
            cell.setBorderWidth(edge, 1D);
        }
    }

    private byte[] readAllBytes(InputStream inputStream) throws Exception {
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

    private void ensurePresentation() {
        if (slideShow == null) {
            createPresentation();
        }
    }
}

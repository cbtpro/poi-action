package com.chenbitao.word.pdf;

import com.chenbitao.word.exception.PdfDocumentException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * PDFBox PDF 文档生成器。
 *
 * <p>基于 Apache PDFBox 生成 {@code .pdf} 文件，支持标题、段落、基础表格和图片。</p>
 */
public class PdfBoxDocumentGenerator implements PdfDocumentGenerator {

    private static final float MARGIN = 54F;
    private static final float LINE_HEIGHT = 16F;
    private static final float TITLE_FONT_SIZE = 22F;
    private static final float BODY_FONT_SIZE = 11F;
    private static final float TABLE_ROW_HEIGHT = 24F;

    private PDDocument document;
    private PDPage currentPage;
    private float cursorY;

    /**
     * 创建新的 PDF 文档。
     */
    @Override
    public void createDocument() {
        closeQuietly();
        document = new PDDocument();
        newPage();
    }

    /**
     * 添加标题。
     *
     * @param text 标题文本
     */
    @Override
    public void addTitle(String text) {
        ensureDocument();
        writeLines(wrap(text, 64), PDType1Font.HELVETICA_BOLD, TITLE_FONT_SIZE, 22F, new Color(31, 78, 121));
        cursorY -= 8F;
    }

    /**
     * 添加段落。
     *
     * @param text 段落文本
     */
    @Override
    public void addParagraph(String text) {
        ensureDocument();
        writeLines(wrap(text, 92), PDType1Font.HELVETICA, BODY_FONT_SIZE, LINE_HEIGHT, Color.DARK_GRAY);
        cursorY -= 8F;
    }

    /**
     * 添加表格。
     *
     * @param rows 表格数据，每个内部列表代表一行
     */
    @Override
    public void addTable(List<List<String>> rows) {
        ensureDocument();
        if (rows == null || rows.isEmpty()) {
            return;
        }

        int columns = maxColumnCount(rows);
        float tableWidth = contentWidth();
        float columnWidth = tableWidth / columns;
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            ensureSpace(TABLE_ROW_HEIGHT + 8F);
            List<String> row = rows.get(rowIndex) == null ? Collections.<String>emptyList() : rows.get(rowIndex);
            drawTableRow(row, columns, columnWidth, rowIndex == 0);
            cursorY -= TABLE_ROW_HEIGHT;
        }
        cursorY -= 12F;
    }

    /**
     * 添加图片。
     *
     * @param image 图片输入流
     * @param width 图片宽度
     * @param height 图片高度
     */
    @Override
    public void addImage(InputStream image, float width, float height) {
        ensureDocument();
        try (InputStream input = image) {
            BufferedImage bufferedImage = ImageIO.read(input);
            if (bufferedImage == null) {
                throw new PdfDocumentException("无法识别 PDF 图片内容");
            }
            ensureSpace(height + 12F);
            PDImageXObject imageObject = LosslessFactory.createFromImage(document, bufferedImage);
            try (PDPageContentStream stream = appendStream()) {
                stream.drawImage(imageObject, MARGIN, cursorY - height, width, height);
            }
            cursorY -= height + 12F;
        } catch (PdfDocumentException e) {
            throw e;
        } catch (Exception e) {
            throw new PdfDocumentException("添加 PDF 图片失败", e);
        }
    }

    /**
     * 保存 PDF 到指定路径。
     *
     * @param path 输出文件路径
     */
    @Override
    public void save(String path) {
        ensureDocument();
        try {
            document.save(new File(path));
        } catch (Exception e) {
            throw new PdfDocumentException("保存 PDF 文件失败", e);
        } finally {
            closeQuietly();
        }
    }

    /**
     * 保存 PDF 到输出流。
     *
     * @param outputStream 输出流
     */
    @Override
    public void save(OutputStream outputStream) {
        ensureDocument();
        try {
            document.save(outputStream);
        } catch (Exception e) {
            throw new PdfDocumentException("写出 PDF 失败", e);
        } finally {
            closeQuietly();
        }
    }

    private void writeLines(List<String> lines,
                            PDType1Font font,
                            float fontSize,
                            float lineHeight,
                            Color color) {
        for (String line : lines) {
            ensureSpace(lineHeight);
            try (PDPageContentStream stream = appendStream()) {
                stream.beginText();
                stream.setFont(font, fontSize);
                stream.setNonStrokingColor(color);
                stream.newLineAtOffset(MARGIN, cursorY);
                stream.showText(safeText(line));
                stream.endText();
            } catch (Exception e) {
                throw new PdfDocumentException("写入 PDF 文本失败", e);
            }
            cursorY -= lineHeight;
        }
    }

    private void drawTableRow(List<String> row, int columns, float columnWidth, boolean header) {
        try (PDPageContentStream stream = appendStream()) {
            float y = cursorY - TABLE_ROW_HEIGHT;
            for (int column = 0; column < columns; column++) {
                float x = MARGIN + column * columnWidth;
                if (header) {
                    stream.setNonStrokingColor(new Color(91, 155, 213));
                    stream.addRect(x, y, columnWidth, TABLE_ROW_HEIGHT);
                    stream.fill();
                }
                stream.setStrokingColor(new Color(180, 180, 180));
                stream.addRect(x, y, columnWidth, TABLE_ROW_HEIGHT);
                stream.stroke();

                String value = column < row.size() ? row.get(column) : "";
                stream.beginText();
                stream.setFont(header ? PDType1Font.HELVETICA_BOLD : PDType1Font.HELVETICA, BODY_FONT_SIZE);
                stream.setNonStrokingColor(header ? Color.WHITE : Color.DARK_GRAY);
                stream.newLineAtOffset(x + 6F, y + 8F);
                stream.showText(safeText(value, Math.max(8, (int) (columnWidth / 6F))));
                stream.endText();
            }
        } catch (Exception e) {
            throw new PdfDocumentException("绘制 PDF 表格失败", e);
        }
    }

    private PDPageContentStream appendStream() throws Exception {
        return new PDPageContentStream(document, currentPage, PDPageContentStream.AppendMode.APPEND, true, true);
    }

    private List<String> wrap(String text, int lineLength) {
        String safe = safeText(text);
        if (safe.isEmpty()) {
            return Collections.singletonList("");
        }

        List<String> lines = new ArrayList<>();
        String[] paragraphs = safe.split("\\r?\\n");
        for (String paragraph : paragraphs) {
            for (int start = 0; start < paragraph.length(); start += lineLength) {
                lines.add(paragraph.substring(start, Math.min(paragraph.length(), start + lineLength)));
            }
        }
        return lines;
    }

    private String safeText(String text) {
        return safeText(text, Integer.MAX_VALUE);
    }

    private String safeText(String text, int limit) {
        if (text == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        int max = Math.min(text.length(), limit);
        for (int i = 0; i < max; i++) {
            char value = text.charAt(i);
            builder.append(value >= 32 && value <= 126 ? value : '?');
        }
        return builder.toString();
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

    private float contentWidth() {
        return currentPage.getMediaBox().getWidth() - MARGIN * 2;
    }

    private void ensureSpace(float requiredHeight) {
        if (cursorY - requiredHeight < MARGIN) {
            newPage();
        }
    }

    private void newPage() {
        currentPage = new PDPage(PDRectangle.LETTER);
        document.addPage(currentPage);
        cursorY = currentPage.getMediaBox().getHeight() - MARGIN;
    }

    private void ensureDocument() {
        if (document == null) {
            createDocument();
        }
    }

    private void closeQuietly() {
        if (document != null) {
            try {
                document.close();
            } catch (Exception ignored) {
                // 释放失败不影响调用方看到原始生成或写出结果。
            }
            document = null;
            currentPage = null;
        }
    }
}

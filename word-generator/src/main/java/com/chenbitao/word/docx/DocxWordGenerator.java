package com.chenbitao.word.docx;

import com.chenbitao.word.core.AbstractWordGenerator;
import com.chenbitao.word.core.WordTable;
import com.chenbitao.word.core.WordTableCell;
import com.chenbitao.word.exception.WordException;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.List;

/**
 * DOCX格式Word文档生成器
 * 使用Apache POI的XWPF组件生成现代的DOCX格式Word文档
 * 支持完整的文档功能，包括文本、标题、表格和图片
 */
public class DocxWordGenerator extends AbstractWordGenerator {

    /** DOCX文档对象 */
    private XWPFDocument document;

    /**
     * 创建新的DOCX文档实例
     */
    @Override
    public void createDocument() {
        document = new XWPFDocument();
    }

    /**
     * 添加普通文本段落
     *
     * @param text 段落文本内容
     */
    @Override
    public void addParagraph(String text) {
        XWPFParagraph p = document.createParagraph();
        XWPFRun run = p.createRun();
        run.setText(text);
        run.setFontFamily(font);
    }

    /**
     * 添加标题段落
     * 使用Word内置的标题样式
     *
     * @param text 标题文本内容
     * @param level 标题级别（1-6级）
     */
    @Override
    public void addTitle(String text, int level) {
        XWPFParagraph p = document.createParagraph();
        p.setStyle("Heading" + level);
        XWPFRun run = p.createRun();
        run.setText(text);
    }

    /**
     * 添加表格
     * 创建指定行数和列数的表格，并填充默认内容
     *
     * @param rows 表格行数
     * @param cols 表格列数
     */
    @Override
    public void addTable(int rows, int cols) {
        XWPFTable table = document.createTable(rows, cols);
        fillPageWidth(table);
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < cols; j++) {
                table.getRow(i).getCell(j).setText("cell");
            }
        }
    }

    /**
     * 添加带内容的表格
     *
     * @param rows 表格数据，每个内部列表代表一行
     */
    @Override
    public void addTable(List<List<String>> rows) {
        if (rows == null || rows.isEmpty()) {
            return;
        }

        int rowCount = rows.size();
        int colCount = maxColumnCount(rows);
        XWPFTable table = document.createTable(rowCount, colCount);
        fillPageWidth(table);

        for (int i = 0; i < rowCount; i++) {
            List<String> rowData = rows.get(i);
            for (int j = 0; j < colCount; j++) {
                String text = rowData != null && j < rowData.size() ? rowData.get(j) : "";
                table.getRow(i).getCell(j).setText(text == null ? "" : text);
            }
        }
    }

    /**
     * 添加结构化表格
     *
     * @param wordTable 表格模型
     */
    @Override
    public void addTable(WordTable wordTable) {
        if (wordTable == null || wordTable.getRows().isEmpty()) {
            return;
        }

        List<List<WordTableCell>> rows = wordTable.getRows();
        int rowCount = rows.size();
        int colCount = maxStructuredColumnCount(rows);
        XWPFTable table = document.createTable(rowCount, colCount);
        fillPageWidth(table);
        boolean[][] occupied = new boolean[rowCount][colCount];

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            int columnIndex = 0;
            for (WordTableCell wordCell : rows.get(rowIndex)) {
                columnIndex = nextFreeColumn(occupied, rowIndex, columnIndex);
                if (columnIndex >= colCount) {
                    break;
                }

                WordTableCell cell = wordCell == null ? WordTableCell.empty() : wordCell;
                XWPFTableCell tableCell = table.getRow(rowIndex).getCell(columnIndex);
                fillStructuredCell(tableCell, cell);
                applyHorizontalMerge(tableCell, cell);
                applyVerticalMerge(table, occupied, rowIndex, columnIndex, cell);
                markOccupied(occupied, rowIndex, columnIndex, cell);
                columnIndex += cell.getColSpan();
            }
        }
    }

    private void fillPageWidth(XWPFTable table) {
        CTTblPr tblPr = table.getCTTbl().getTblPr() == null
                ? table.getCTTbl().addNewTblPr()
                : table.getCTTbl().getTblPr();
        CTTblWidth width = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        width.setType(STTblWidth.PCT);
        width.setW(BigInteger.valueOf(5000));
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

    private int maxStructuredColumnCount(List<List<WordTableCell>> rows) {
        int max = 1;
        for (List<WordTableCell> row : rows) {
            int count = 0;
            for (WordTableCell cell : row) {
                count += cell == null ? 1 : cell.getColSpan();
            }
            if (count > max) {
                max = count;
            }
        }
        return max;
    }

    private int nextFreeColumn(boolean[][] occupied, int rowIndex, int columnIndex) {
        int column = columnIndex;
        while (column < occupied[rowIndex].length && occupied[rowIndex][column]) {
            column++;
        }
        return column;
    }

    private void fillStructuredCell(XWPFTableCell tableCell, WordTableCell cell) {
        if (cell.hasImage()) {
            addImageToCell(tableCell, cell);
        } else {
            tableCell.setText(cell.getText() == null ? "" : cell.getText());
        }
    }

    private void addImageToCell(XWPFTableCell tableCell, WordTableCell cell) {
        try {
            tableCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            XWPFParagraph paragraph = firstParagraph(tableCell);
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            run.addPicture(cell.getImage(),
                    Document.PICTURE_TYPE_PNG,
                    "img",
                    cell.getImageWidthEmu(),
                    cell.getImageHeightEmu());
        } catch (Exception e) {
            throw new WordException("表格图片插入失败", e);
        }
    }

    private XWPFParagraph firstParagraph(XWPFTableCell tableCell) {
        if (tableCell.getParagraphs().isEmpty()) {
            return tableCell.addParagraph();
        }
        XWPFParagraph paragraph = tableCell.getParagraphs().get(0);
        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
        return paragraph;
    }

    private void applyHorizontalMerge(XWPFTableCell tableCell, WordTableCell cell) {
        if (cell.getColSpan() <= 1) {
            return;
        }

        CTTcPr tcPr = tableCell.getCTTc().isSetTcPr()
                ? tableCell.getCTTc().getTcPr()
                : tableCell.getCTTc().addNewTcPr();
        tcPr.addNewGridSpan().setVal(BigInteger.valueOf(cell.getColSpan()));
    }

    private void applyVerticalMerge(XWPFTable table, boolean[][] occupied,
                                    int rowIndex, int columnIndex, WordTableCell cell) {
        if (cell.getRowSpan() <= 1) {
            return;
        }

        int endRow = Math.min(occupied.length, rowIndex + cell.getRowSpan());
        table.getRow(rowIndex).getCell(columnIndex).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        setVerticalMerge(table.getRow(rowIndex).getCell(columnIndex), true);
        for (int i = rowIndex + 1; i < endRow; i++) {
            setVerticalMerge(table.getRow(i).getCell(columnIndex), false);
        }
    }

    private void setVerticalMerge(XWPFTableCell cell, boolean restart) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr()
                ? cell.getCTTc().getTcPr()
                : cell.getCTTc().addNewTcPr();
        CTVMerge merge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
        merge.setVal(restart ? STMerge.RESTART : STMerge.CONTINUE);
    }

    private void markOccupied(boolean[][] occupied, int rowIndex, int columnIndex, WordTableCell cell) {
        int endRow = Math.min(occupied.length, rowIndex + cell.getRowSpan());
        int endColumn = Math.min(occupied[rowIndex].length, columnIndex + cell.getColSpan());
        for (int i = rowIndex; i < endRow; i++) {
            for (int j = columnIndex; j < endColumn; j++) {
                occupied[i][j] = true;
            }
        }
    }

    /**
     * 添加图片
     * 将图片插入到文档中
     *
     * @param inputStream 图片输入流
     * @param width 图片宽度（EMU单位）
     * @param height 图片高度（EMU单位）
     * @throws WordException 如果图片插入失败
     */
    @Override
    public void addImage(InputStream inputStream, int width, int height) {
        try {
            XWPFParagraph p = document.createParagraph();
            XWPFRun run = p.createRun();
            run.addPicture(inputStream,
                    Document.PICTURE_TYPE_PNG,
                    "img",
                    width,
                    height);
        } catch (Exception e) {
            throw new WordException("图片插入失败", e);
        }
    }

    /**
     * 保存DOCX文档到指定路径
     *
     * @param path 输出文件路径
     * @throws WordException 如果保存失败
     */
    @Override
    public void save(String path) {
        try (FileOutputStream out = new FileOutputStream(path)) {
            document.write(out);
        } catch (Exception e) {
            throw new WordException("保存失败", e);
        }
    }
}

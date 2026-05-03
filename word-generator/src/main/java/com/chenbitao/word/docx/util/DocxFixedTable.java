package com.chenbitao.word.docx.util;

import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblCellMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.ByteArrayInputStream;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;

/**
 * 固定网格 DOCX 表格渲染工具。
 */
public final class DocxFixedTable {

    private static final int DEFAULT_BORDER_SIZE = 4;

    private DocxFixedTable() {
    }

    /**
     * 渲染固定网格表格。
     *
     * @param document 文档对象
     * @param options 表格选项
     * @param rows 表格行
     * @return 生成的表格对象
     */
    public static XWPFTable render(XWPFDocument document, Options options, List<Row> rows) {
        XWPFTable table = document.createTable();
        configureTable(table, options);

        for (int i = 0; i < rows.size(); i++) {
            Row rowSpec = rows.get(i);
            XWPFTableRow row = prepareRow(table, i, rowSpec.getCells().size());
            if (rowSpec.getHeightDxa() > 0) {
                setRowHeight(row, rowSpec.getHeightDxa());
            }

            int columnIndex = 0;
            for (int j = 0; j < rowSpec.getCells().size(); j++) {
                Cell cellSpec = rowSpec.getCells().get(j);
                XWPFTableCell cell = row.getCell(j);
                int width = options.width(columnIndex, cellSpec.getColSpan());
                configureCell(cell, cellSpec, width, options);
                columnIndex += cellSpec.getColSpan();
            }
        }
        return table;
    }

    public static Row row(Cell... cells) {
        return row(0, cells);
    }

    public static Row row(int heightDxa, Cell... cells) {
        return new Row(heightDxa, Arrays.asList(cells));
    }

    public static Cell text(String text) {
        return new Cell(text, null, 0, 0);
    }

    public static Cell empty() {
        return text("");
    }

    public static Cell image(byte[] imageBytes, int widthEmu, int heightEmu) {
        return new Cell("", imageBytes, widthEmu, heightEmu);
    }

    private static void configureTable(XWPFTable table, Options options) {
        CTTbl ctTbl = table.getCTTbl();
        CTTblPr tblPr = getOrCreateTableProperties(ctTbl);

        configureTableWidth(tblPr, options.getTableWidthDxa());
        configureTableLayout(tblPr);
        configureTableBorders(tblPr);
        configureTableCellMargins(tblPr, options);
        configureTableGrid(ctTbl, options.getColumnWidths());
    }

    private static CTTblPr getOrCreateTableProperties(CTTbl ctTbl) {
        return ctTbl.getTblPr() == null ? ctTbl.addNewTblPr() : ctTbl.getTblPr();
    }

    private static void configureTableWidth(CTTblPr tblPr, int tableWidthDxa) {
        CTTblWidth width = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        width.setW(BigInteger.valueOf(tableWidthDxa));
        width.setType(STTblWidth.DXA);
    }

    private static void configureTableLayout(CTTblPr tblPr) {
        CTTblLayoutType layout = tblPr.isSetTblLayout() ? tblPr.getTblLayout() : tblPr.addNewTblLayout();
        layout.setType(STTblLayoutType.FIXED);
    }

    private static void configureTableBorders(CTTblPr tblPr) {
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();

        setBorder(borders.isSetTop() ? borders.getTop() : borders.addNewTop());
        setBorder(borders.isSetLeft() ? borders.getLeft() : borders.addNewLeft());
        setBorder(borders.isSetBottom() ? borders.getBottom() : borders.addNewBottom());
        setBorder(borders.isSetRight() ? borders.getRight() : borders.addNewRight());
        setBorder(borders.isSetInsideH() ? borders.getInsideH() : borders.addNewInsideH());
        setBorder(borders.isSetInsideV() ? borders.getInsideV() : borders.addNewInsideV());
    }

    private static void configureTableCellMargins(CTTblPr tblPr, Options options) {
        CTTblCellMar cellMar = tblPr.isSetTblCellMar() ? tblPr.getTblCellMar() : tblPr.addNewTblCellMar();

        setCellMargin(cellMar.isSetTop() ? cellMar.getTop() : cellMar.addNewTop(), options.getMarginTopDxa());
        setCellMargin(cellMar.isSetBottom() ? cellMar.getBottom() : cellMar.addNewBottom(), options.getMarginBottomDxa());
        setCellMargin(cellMar.isSetLeft() ? cellMar.getLeft() : cellMar.addNewLeft(), options.getMarginLeftDxa());
        setCellMargin(cellMar.isSetRight() ? cellMar.getRight() : cellMar.addNewRight(), options.getMarginRightDxa());
    }

    private static void configureTableGrid(CTTbl ctTbl, int[] columnWidths) {
        CTTblGrid grid = ctTbl.getTblGrid() == null ? ctTbl.addNewTblGrid() : ctTbl.getTblGrid();
        while (grid.sizeOfGridColArray() > 0) {
            grid.removeGridCol(grid.sizeOfGridColArray() - 1);
        }

        for (int columnWidth : columnWidths) {
            CTTblGridCol column = grid.addNewGridCol();
            column.setW(BigInteger.valueOf(columnWidth));
        }
    }

    private static void configureCell(XWPFTableCell cell, Cell spec, int width, Options options) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr()
                ? cell.getCTTc().getTcPr()
                : cell.getCTTc().addNewTcPr();

        CTTblWidth tcW = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        tcW.setW(BigInteger.valueOf(width));
        tcW.setType(STTblWidth.DXA);

        if (spec.getColSpan() > 1) {
            tcPr.addNewGridSpan().setVal(BigInteger.valueOf(spec.getColSpan()));
        }
        if (spec.getMerge() != Merge.NONE) {
            CTVMerge merge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
            merge.setVal(spec.getMerge() == Merge.RESTART ? STMerge.RESTART : STMerge.CONTINUE);
        }

        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph paragraph = resetParagraph(cell);
        paragraph.setAlignment(spec.getAlignment());
        if (spec.getImageBytes() == null) {
            writeText(paragraph.createRun(), spec.getText(), options);
        } else {
            writeImage(paragraph.createRun(), spec);
        }
    }

    private static void writeText(XWPFRun run, String text, Options options) {
        DocxPageUtils.applyFont(run, options.getFont(), options.getFontSize());
        String[] lines = text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                run.addBreak();
            }
            run.setText(lines[i]);
        }
    }

    private static void writeImage(XWPFRun run, Cell spec) {
        try {
            run.addPicture(new ByteArrayInputStream(spec.getImageBytes()),
                    Document.PICTURE_TYPE_PNG,
                    "image.png",
                    spec.getImageWidthEmu(),
                    spec.getImageHeightEmu());
        } catch (Exception e) {
            throw new IllegalStateException("表格图片插入失败", e);
        }
    }

    private static XWPFParagraph resetParagraph(XWPFTableCell cell) {
        while (cell.getParagraphs().size() > 1) {
            cell.removeParagraph(1);
        }

        XWPFParagraph paragraph = cell.getParagraphs().isEmpty()
                ? cell.addParagraph()
                : cell.getParagraphs().get(0);
        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
        return paragraph;
    }

    private static XWPFTableRow prepareRow(XWPFTable table, int rowIndex, int cellCount) {
        XWPFTableRow row = rowIndex == 0 ? table.getRow(0) : table.createRow();
        while (row.getTableCells().size() > cellCount) {
            int position = row.getTableCells().size() - 1;
            row.removeCell(position);
            row.getCtRow().removeTc(position);
        }
        while (row.getTableCells().size() < cellCount) {
            row.addNewTableCell();
        }
        return row;
    }

    private static void setRowHeight(XWPFTableRow row, int heightDxa) {
        CTTrPr trPr = row.getCtRow().isSetTrPr() ? row.getCtRow().getTrPr() : row.getCtRow().addNewTrPr();
        CTHeight height = trPr.sizeOfTrHeightArray() > 0 ? trPr.getTrHeightArray(0) : trPr.addNewTrHeight();
        height.setVal(BigInteger.valueOf(heightDxa));
        height.setHRule(STHeightRule.AT_LEAST);
    }

    private static void setBorder(CTBorder border) {
        border.setVal(STBorder.SINGLE);
        border.setColor("auto");
        border.setSz(BigInteger.valueOf(DEFAULT_BORDER_SIZE));
        border.setSpace(BigInteger.ZERO);
    }

    private static void setCellMargin(CTTblWidth margin, int width) {
        margin.setW(BigInteger.valueOf(width));
        margin.setType(STTblWidth.DXA);
    }

    public enum Merge {
        NONE,
        RESTART,
        CONTINUE
    }

    public static final class Row {
        private final int heightDxa;
        private final List<Cell> cells;

        private Row(int heightDxa, List<Cell> cells) {
            this.heightDxa = heightDxa;
            this.cells = cells;
        }

        public int getHeightDxa() {
            return heightDxa;
        }

        public List<Cell> getCells() {
            return cells;
        }
    }

    public static final class Cell {
        private final String text;
        private final byte[] imageBytes;
        private final int imageWidthEmu;
        private final int imageHeightEmu;
        private int colSpan = 1;
        private Merge merge = Merge.NONE;
        private ParagraphAlignment alignment = ParagraphAlignment.LEFT;

        private Cell(String text, byte[] imageBytes, int imageWidthEmu, int imageHeightEmu) {
            this.text = text == null ? "" : text;
            this.imageBytes = imageBytes;
            this.imageWidthEmu = imageWidthEmu;
            this.imageHeightEmu = imageHeightEmu;
        }

        public Cell span(int colSpan) {
            this.colSpan = Math.max(1, colSpan);
            return this;
        }

        public Cell vRestart() {
            this.merge = Merge.RESTART;
            return this;
        }

        public Cell vContinue() {
            this.merge = Merge.CONTINUE;
            return this;
        }

        public Cell verticalMerge(boolean restart) {
            this.merge = restart ? Merge.RESTART : Merge.CONTINUE;
            return this;
        }

        public Cell center() {
            this.alignment = ParagraphAlignment.CENTER;
            return this;
        }

        public String getText() {
            return text;
        }

        public byte[] getImageBytes() {
            return imageBytes;
        }

        public int getImageWidthEmu() {
            return imageWidthEmu;
        }

        public int getImageHeightEmu() {
            return imageHeightEmu;
        }

        public int getColSpan() {
            return colSpan;
        }

        public Merge getMerge() {
            return merge;
        }

        public ParagraphAlignment getAlignment() {
            return alignment;
        }
    }

    public static final class Options {
        private final int[] columnWidths;
        private int tableWidthDxa;
        private int marginTopDxa;
        private int marginRightDxa = 108;
        private int marginBottomDxa;
        private int marginLeftDxa = 108;
        private String font = "宋体";
        private int fontSize = 12;

        public Options(int[] columnWidths) {
            this.columnWidths = Arrays.copyOf(columnWidths, columnWidths.length);
            this.tableWidthDxa = sum(columnWidths);
        }

        public Options tableWidthDxa(int tableWidthDxa) {
            this.tableWidthDxa = tableWidthDxa;
            return this;
        }

        public Options cellMargins(int topDxa, int rightDxa, int bottomDxa, int leftDxa) {
            this.marginTopDxa = topDxa;
            this.marginRightDxa = rightDxa;
            this.marginBottomDxa = bottomDxa;
            this.marginLeftDxa = leftDxa;
            return this;
        }

        public Options font(String font, int fontSize) {
            this.font = font;
            this.fontSize = fontSize;
            return this;
        }

        public int width(int startColumn, int colSpan) {
            int width = 0;
            for (int i = startColumn; i < startColumn + colSpan && i < columnWidths.length; i++) {
                width += columnWidths[i];
            }
            return width;
        }

        public int[] getColumnWidths() {
            return Arrays.copyOf(columnWidths, columnWidths.length);
        }

        public int getTableWidthDxa() {
            return tableWidthDxa;
        }

        public int getMarginTopDxa() {
            return marginTopDxa;
        }

        public int getMarginRightDxa() {
            return marginRightDxa;
        }

        public int getMarginBottomDxa() {
            return marginBottomDxa;
        }

        public int getMarginLeftDxa() {
            return marginLeftDxa;
        }

        public String getFont() {
            return font;
        }

        public int getFontSize() {
            return fontSize;
        }

        private static int sum(int[] values) {
            int result = 0;
            for (int value : values) {
                result += value;
            }
            return result;
        }
    }
}

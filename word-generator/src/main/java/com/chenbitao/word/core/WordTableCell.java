package com.chenbitao.word.core;

import java.io.InputStream;

/**
 * 编程式Word表格单元格模型。
 * 单元格可以承载文本或图片，并支持跨行、跨列。
 */
public class WordTableCell {

    /** 单元格文本 */
    private final String text;
    /** 图片输入流 */
    private final InputStream image;
    /** 图片宽度（EMU单位） */
    private final int imageWidthEmu;
    /** 图片高度（EMU单位） */
    private final int imageHeightEmu;
    /** 跨行数 */
    private int rowSpan = 1;
    /** 跨列数 */
    private int colSpan = 1;

    private WordTableCell(String text, InputStream image, int imageWidthEmu, int imageHeightEmu) {
        this.text = text;
        this.image = image;
        this.imageWidthEmu = imageWidthEmu;
        this.imageHeightEmu = imageHeightEmu;
    }

    /**
     * 创建文本单元格。
     *
     * @param text 单元格文本
     * @return 文本单元格
     */
    public static WordTableCell text(String text) {
        return new WordTableCell(text, null, 0, 0);
    }

    /**
     * 创建空单元格。
     *
     * @return 空单元格
     */
    public static WordTableCell empty() {
        return text("");
    }

    /**
     * 创建图片单元格。
     *
     * @param image 图片输入流
     * @param widthEmu 图片宽度（EMU单位）
     * @param heightEmu 图片高度（EMU单位）
     * @return 图片单元格
     */
    public static WordTableCell image(InputStream image, int widthEmu, int heightEmu) {
        return new WordTableCell("", image, widthEmu, heightEmu);
    }

    /**
     * 设置跨行数。
     *
     * @param rowSpan 跨行数，最小为1
     * @return 当前单元格
     */
    public WordTableCell rowSpan(int rowSpan) {
        this.rowSpan = Math.max(1, rowSpan);
        return this;
    }

    /**
     * 设置跨列数。
     *
     * @param colSpan 跨列数，最小为1
     * @return 当前单元格
     */
    public WordTableCell colSpan(int colSpan) {
        this.colSpan = Math.max(1, colSpan);
        return this;
    }

    public String getText() {
        return text;
    }

    public InputStream getImage() {
        return image;
    }

    public int getImageWidthEmu() {
        return imageWidthEmu;
    }

    public int getImageHeightEmu() {
        return imageHeightEmu;
    }

    public int getRowSpan() {
        return rowSpan;
    }

    public int getColSpan() {
        return colSpan;
    }

    public boolean hasImage() {
        return image != null;
    }
}

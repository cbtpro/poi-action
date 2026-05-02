package com.chenbitao.word.docx;

import com.chenbitao.word.core.AbstractWordGenerator;
import com.chenbitao.word.exception.WordException;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.InputStream;

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
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < cols; j++) {
                table.getRow(i).getCell(j).setText("cell");
            }
        }
    }

    /**
     * 添加图片
     * 将图片插入到文档中
     *
     * @param inputStream 图片输入流
     * @param width 图片宽度（像素）
     * @param height 图片高度（像素）
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
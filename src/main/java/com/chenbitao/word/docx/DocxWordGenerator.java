package com.chenbitao.word.docx;

import com.chenbitao.word.core.AbstractWordGenerator;
import com.chenbitao.word.exception.WordException;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.InputStream;

public class DocxWordGenerator extends AbstractWordGenerator {

    private XWPFDocument document;

    @Override
    public void createDocument() {
        document = new XWPFDocument();
    }

    @Override
    public void addParagraph(String text) {
        XWPFParagraph p = document.createParagraph();
        XWPFRun run = p.createRun();
        run.setText(text);
        run.setFontFamily(font);
    }

    @Override
    public void addTitle(String text, int level) {
        XWPFParagraph p = document.createParagraph();
        p.setStyle("Heading" + level);
        XWPFRun run = p.createRun();
        run.setText(text);
    }

    @Override
    public void addTable(int rows, int cols) {
        XWPFTable table = document.createTable(rows, cols);
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < cols; j++) {
                table.getRow(i).getCell(j).setText("cell");
            }
        }
    }

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

    @Override
    public void save(String path) {
        try (FileOutputStream out = new FileOutputStream(path)) {
            document.write(out);
        } catch (Exception e) {
            throw new WordException("保存失败", e);
        }
    }
}
package com.chenbitao.word.doc;

import com.chenbitao.word.core.AbstractWordGenerator;
import com.chenbitao.word.exception.WordException;
import org.apache.poi.hwpf.HWPFDocument;

import java.io.FileOutputStream;
import java.io.InputStream;

public class DocWordGenerator extends AbstractWordGenerator {

    private HWPFDocument document;

    @Override
    public void createDocument() {
        try (InputStream is = getClass().getResourceAsStream("/template.doc")) {
            if (is == null) {
                throw new RuntimeException("template.doc 不存在");
            }
            document = new HWPFDocument(is);
        } catch (Exception e) {
            throw new WordException("创建DOC失败", e);
        }
    }

    @Override
    public void addParagraph(String text) {
        try {
            document.getRange().insertAfter(text + "\n");
        } catch (Exception e) {
            throw new WordException("写入失败", e);
        }
    }

    @Override
    public void addTitle(String text, int level) {
        addParagraph(text); // 简化处理
    }

    @Override
    public void addTable(int rows, int cols) {
        addParagraph("[DOC 不支持复杂表格]");
    }

    @Override
    public void addImage(InputStream inputStream, int width, int height) {
        addParagraph("[DOC 图片支持有限]");
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
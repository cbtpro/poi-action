package com.chenbitao.word.doc;

import com.chenbitao.word.core.AbstractWordGenerator;
import com.chenbitao.word.exception.WordException;
import org.apache.poi.hwpf.HWPFDocument;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * DOC格式Word文档生成器
 * 使用Apache POI的HWPF组件生成传统的DOC格式Word文档
 * 支持基本的文本添加功能，但对表格和图片支持有限
 */
public class DocWordGenerator extends AbstractWordGenerator {

    /** DOC文档对象 */
    private HWPFDocument document;

    /**
     * 创建新的DOC文档实例
     * 从模板文件template.doc创建文档对象
     *
     * @throws WordException 如果模板文件不存在或读取失败
     */
    @Override
    public void createDocument() {
        try (InputStream is = getClass().getResourceAsStream("/template.doc")) {
            if (is == null) {
                throw new FileNotFoundException("template.doc 不存在");
            }
            document = new HWPFDocument(is);
        } catch (Exception e) {
            throw new WordException("创建DOC失败", e);
        }
    }

    /**
     * 添加普通文本段落
     *
     * @param text 段落文本内容
     * @throws WordException 如果写入失败
     */
    @Override
    public void addParagraph(String text) {
        try {
            document.getRange().insertAfter(text + "\n");
        } catch (Exception e) {
            throw new WordException("写入失败", e);
        }
    }

    /**
     * 添加标题段落
     * DOC格式简化处理，直接作为普通段落添加
     *
     * @param text 标题文本内容
     * @param level 标题级别（忽略）
     */
    @Override
    public void addTitle(String text, int level) {
        addParagraph(text); // 简化处理
    }

    /**
     * 添加表格
     * DOC格式不支持复杂表格，添加提示信息
     *
     * @param rows 表格行数（忽略）
     * @param cols 表格列数（忽略）
     */
    @Override
    public void addTable(int rows, int cols) {
        addParagraph("[DOC 不支持复杂表格]");
    }

    /**
     * 添加图片
     * DOC格式图片支持有限，添加提示信息
     *
     * @param inputStream 图片输入流（忽略）
     * @param width 图片宽度（忽略）
     * @param height 图片高度（忽略）
     */
    @Override
    public void addImage(InputStream inputStream, int width, int height) {
        addParagraph("[DOC 图片支持有限]");
    }

    /**
     * 保存DOC文档到指定路径
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
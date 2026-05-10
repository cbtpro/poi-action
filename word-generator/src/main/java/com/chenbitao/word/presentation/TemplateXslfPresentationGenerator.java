package com.chenbitao.word.presentation;

import com.chenbitao.word.exception.PresentationException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

/**
 * PPTX 模板演示文稿生成器。
 *
 * <p>基于 Apache POI XSLF 渲染 PowerPoint Open XML 的 {@code .pptx} 模板。</p>
 */
public class TemplateXslfPresentationGenerator {

    /** PPTX 模板演示文稿对象。 */
    private final XMLSlideShow slideShow;

    /**
     * 构造 PPTX 模板演示文稿生成器。
     *
     * @param template 模板输入流
     * @throws PresentationException 如果模板加载失败
     */
    public TemplateXslfPresentationGenerator(InputStream template) {
        try {
            this.slideShow = new XMLSlideShow(template);
        } catch (Exception e) {
            throw new PresentationException("加载 PPTX 模板失败", e);
        }
    }

    /**
     * 渲染模板演示文稿。
     *
     * @param data 模板数据
     */
    public void render(Map<String, Object> data) {
        for (XSLFSlide slide : slideShow.getSlides()) {
            renderSlide(slide, data);
        }
    }

    /**
     * 保存演示文稿到指定路径。
     *
     * @param path 输出文件路径
     */
    public void save(String path) {
        try (FileOutputStream output = new FileOutputStream(path)) {
            save(output);
        } catch (Exception e) {
            throw new PresentationException("保存 PPTX 模板失败", e);
        }
    }

    /**
     * 保存演示文稿到输出流。
     *
     * @param outputStream 输出流
     */
    public void save(OutputStream outputStream) {
        try {
            slideShow.write(outputStream);
        } catch (Exception e) {
            throw new PresentationException("写出 PPTX 模板失败", e);
        }
    }

    /**
     * 渲染单页幻灯片。
     *
     * @param slide 幻灯片
     * @param data 模板数据
     */
    private void renderSlide(XSLFSlide slide, Map<String, Object> data) {
        for (XSLFShape shape : slide.getShapes()) {
            renderShape(shape, data);
        }
    }

    /**
     * 渲染单个图形。
     *
     * @param shape 图形对象
     * @param data 模板数据
     */
    private void renderShape(XSLFShape shape, Map<String, Object> data) {
        if (shape instanceof XSLFTable) {
            renderTable((XSLFTable) shape, data);
            return;
        }
        if (shape instanceof XSLFTextShape) {
            renderTextShape((XSLFTextShape) shape, data);
            return;
        }
        if (shape instanceof XSLFGroupShape) {
            renderGroup((XSLFGroupShape) shape, data);
        }
    }

    /**
     * 渲染组合图形。
     *
     * @param group 组合图形
     * @param data 模板数据
     */
    private void renderGroup(XSLFGroupShape group, Map<String, Object> data) {
        for (XSLFShape child : group.getShapes()) {
            renderShape(child, data);
        }
    }

    /**
     * 渲染表格。
     *
     * @param table 表格图形
     * @param data 模板数据
     */
    private void renderTable(XSLFTable table, Map<String, Object> data) {
        for (int row = 0; row < table.getNumberOfRows(); row++) {
            for (int column = 0; column < table.getNumberOfColumns(); column++) {
                renderTextShape(table.getCell(row, column), data);
            }
        }
    }

    /**
     * 渲染文本图形。
     *
     * @param textShape 文本图形
     * @param data 模板数据
     */
    private void renderTextShape(XSLFTextShape textShape, Map<String, Object> data) {
        String text = textShape.getText();
        if (PlaceholderTextRenderer.containsPlaceholder(text)) {
            textShape.setText(PlaceholderTextRenderer.render(text, data));
        }
    }
}

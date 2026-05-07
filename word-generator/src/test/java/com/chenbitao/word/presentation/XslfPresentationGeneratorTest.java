package com.chenbitao.word.presentation;

import com.chenbitao.word.util.ImageBytesUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class XslfPresentationGeneratorTest {

    @Test
    public void saveWritesPptxWithTextTableAndPictureSlides() throws Exception {
        XslfPresentationGenerator generator = new XslfPresentationGenerator();
        generator.createPresentation();
        generator.addTitleSlide("季度复盘", "poi-action");
        generator.addTextSlide("亮点", Arrays.asList("支持 PPTX 生成", "支持表格", "支持图片"));
        generator.addTableSlide("指标", Arrays.asList(
                Arrays.asList("类型", "数量"),
                Arrays.asList("Word", "2"),
                Arrays.asList("PowerPoint", "2")
        ));
        generator.addImageSlide("架构图", new ByteArrayInputStream(demoImage()));

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (XMLSlideShow slideShow = new XMLSlideShow(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals(4, slideShow.getSlides().size());
            assertTrue(slideContainsText(slideShow.getSlides().get(0), "季度复盘"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "支持表格"));
            assertTrue(slideContainsTable(slideShow.getSlides().get(2)));
            assertTrue(tableContainsText(slideShow.getSlides().get(2), "PowerPoint"));
            assertTrue(slideContainsPicture(slideShow.getSlides().get(3)));
            assertEquals(1, slideShow.getPictureData().size());
        }
    }

    private boolean slideContainsText(XSLFSlide slide, String expected) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
                String text = ((XSLFTextShape) shape).getText();
                if (text != null && text.contains(expected)) {
                    return true;
                }
            }
        }
        return false;
    }

    private boolean slideContainsTable(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTable) {
                return true;
            }
        }
        return false;
    }

    private boolean tableContainsText(XSLFSlide slide, String expected) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTable) {
                XSLFTable table = (XSLFTable) shape;
                for (int row = 0; row < table.getNumberOfRows(); row++) {
                    for (int column = 0; column < table.getNumberOfColumns(); column++) {
                        String text = table.getCell(row, column).getText();
                        if (text != null && text.contains(expected)) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    private boolean slideContainsPicture(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFPictureShape) {
                return true;
            }
        }
        return false;
    }

    private byte[] demoImage() throws Exception {
        return ImageBytesUtils.solidPng(160, 100, 0x70AD47);
    }
}

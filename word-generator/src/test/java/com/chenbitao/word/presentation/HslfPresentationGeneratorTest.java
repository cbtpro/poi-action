package com.chenbitao.word.presentation;

import com.chenbitao.word.util.ImageBytesUtils;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.hslf.usermodel.HSLFTextShape;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class HslfPresentationGeneratorTest {

    @Test
    public void saveWritesPptWithTextTableAndPictureSlides() throws Exception {
        HslfPresentationGenerator generator = new HslfPresentationGenerator();
        generator.createPresentation();
        generator.addTitleSlide("季度复盘", "poi-action");
        generator.addTextSlide("亮点", Arrays.asList("支持 PPT 生成", "支持表格", "支持图片"));
        generator.addTableSlide("指标", Arrays.asList(
                Arrays.asList("类型", "数量"),
                Arrays.asList("Word", "2"),
                Arrays.asList("Excel", "2")
        ));
        generator.addImageSlide("架构图", new ByteArrayInputStream(demoImage()));

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        generator.save(output);

        try (HSLFSlideShow slideShow = new HSLFSlideShow(new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals(4, slideShow.getSlides().size());
            assertTrue(slideContainsText(slideShow.getSlides().get(0), "季度复盘"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "支持表格"));
            assertTrue(slideContainsTable(slideShow.getSlides().get(2)));
            assertTrue(tableContainsText(slideShow.getSlides().get(2), "Excel"));
            assertTrue(slideContainsPicture(slideShow.getSlides().get(3)));
            assertEquals(1, slideShow.getPictureData().size());
        }
    }

    private boolean slideContainsText(HSLFSlide slide, String expected) {
        for (HSLFShape shape : slide.getShapes()) {
            if (shape instanceof HSLFTextShape) {
                String text = ((HSLFTextShape) shape).getText();
                if (text != null && text.contains(expected)) {
                    return true;
                }
            }
        }
        return false;
    }

    private boolean slideContainsTable(HSLFSlide slide) {
        for (HSLFShape shape : slide.getShapes()) {
            if (shape instanceof HSLFTable) {
                return true;
            }
        }
        return false;
    }

    private boolean tableContainsText(HSLFSlide slide, String expected) {
        for (HSLFShape shape : slide.getShapes()) {
            if (shape instanceof HSLFTable) {
                HSLFTable table = (HSLFTable) shape;
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

    private boolean slideContainsPicture(HSLFSlide slide) {
        for (HSLFShape shape : slide.getShapes()) {
            if (shape instanceof HSLFPictureShape) {
                return true;
            }
        }
        return false;
    }

    private byte[] demoImage() throws Exception {
        return ImageBytesUtils.solidPng(160, 100, 0x5B9BD5);
    }
}

package com.chenbitao.word.playground.demo.presentation;

import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.hslf.usermodel.HSLFTextShape;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class PptProjectReportDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesProjectReportPresentation() throws Exception {
        File output = temporaryFolder.newFile("project-report-demo.ppt");

        PptProjectReportDemo.generate(output.toPath());

        assertTrue(output.length() > 0);
        try (HSLFSlideShow slideShow = new HSLFSlideShow(Files.newInputStream(output.toPath()))) {
            assertEquals(4, slideShow.getSlides().size());
            assertTrue(slideContainsText(slideShow.getSlides().get(0), "poi-action 项目汇报"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "支持 HSLF"));
            assertTrue(slideContainsTable(slideShow.getSlides().get(2)));
            assertTrue(slideContainsPicture(slideShow.getSlides().get(3)));
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

    private boolean slideContainsPicture(HSLFSlide slide) {
        for (HSLFShape shape : slide.getShapes()) {
            if (shape instanceof HSLFPictureShape) {
                return true;
            }
        }
        return false;
    }
}

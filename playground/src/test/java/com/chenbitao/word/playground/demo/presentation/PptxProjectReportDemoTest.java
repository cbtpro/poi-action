package com.chenbitao.word.playground.demo.presentation;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.nio.file.Files;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class PptxProjectReportDemoTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void generateCreatesProjectReportPresentation() throws Exception {
        File output = temporaryFolder.newFile("project-report-demo.pptx");

        PptxProjectReportDemo.generate(output.toPath());

        assertTrue(output.length() > 0);
        try (XMLSlideShow slideShow = new XMLSlideShow(Files.newInputStream(output.toPath()))) {
            assertEquals(4, slideShow.getSlides().size());
            assertTrue(slideContainsText(slideShow.getSlides().get(0), "poi-action 项目汇报"));
            assertTrue(slideContainsText(slideShow.getSlides().get(1), "支持 XSLF"));
            assertTrue(slideContainsTable(slideShow.getSlides().get(2)));
            assertTrue(slideContainsPicture(slideShow.getSlides().get(3)));
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

    private boolean slideContainsPicture(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFPictureShape) {
                return true;
            }
        }
        return false;
    }
}

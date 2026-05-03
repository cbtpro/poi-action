package com.chenbitao.word.docx.util;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import org.w3c.dom.Element;

import static org.junit.Assert.assertEquals;

public class DocxPageUtilsTest {

    @Test
    public void setupPageWritesPageSizeAndMargins() {
        XWPFDocument document = new XWPFDocument();

        DocxPageUtils.setupPage(document, 11906, 16838,
                new DocxPageUtils.PageMargin(1440, 1800, 1440, 1800, 851, 992, 0));

        Element section = (Element) document.getDocument().getBody().getSectPr().getDomNode();
        Element pageSize = (Element) section.getElementsByTagNameNS(
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "pgSz").item(0);
        Element pageMargin = (Element) section.getElementsByTagNameNS(
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "pgMar").item(0);

        assertEquals("11906", pageSize.getAttributeNS(
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w"));
        assertEquals("16838", pageSize.getAttributeNS(
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "h"));
        assertEquals("1800", pageMargin.getAttributeNS(
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "left"));
        assertEquals("992", pageMargin.getAttributeNS(
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "footer"));
    }

    @Test
    public void addCenteredTitleWritesTextAndFont() {
        XWPFDocument document = new XWPFDocument();

        DocxPageUtils.addCenteredTitle(document, "测试标题", "宋体", 22);

        assertEquals("测试标题", document.getParagraphArray(0).getText());
        assertEquals("宋体", document.getParagraphArray(0).getRuns().get(0).getFontFamily());
        assertEquals(22, document.getParagraphArray(0).getRuns().get(0).getFontSize());
    }
}

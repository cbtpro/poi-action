package com.chenbitao.word.docx.util;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

/**
 * DOCX 页面和字体相关工具。
 */
public final class DocxPageUtils {

    private static final String W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private DocxPageUtils() {
    }

    /**
     * 设置页面大小和边距。
     *
     * @param document 文档对象
     * @param pageWidthDxa 页面宽度（DXA）
     * @param pageHeightDxa 页面高度（DXA）
     * @param margin 页面边距
     */
    public static void setupPage(XWPFDocument document, int pageWidthDxa, int pageHeightDxa, PageMargin margin) {
        CTBody body = document.getDocument().getBody();
        CTSectPr section = body.isSetSectPr() ? body.getSectPr() : body.addNewSectPr();
        Node sectionNode = section.getDomNode();
        removeChild(sectionNode, "pgSz");
        removeChild(sectionNode, "pgMar");

        Element pageSize = sectionNode.getOwnerDocument().createElementNS(W_NS, "w:pgSz");
        pageSize.setAttributeNS(W_NS, "w:w", String.valueOf(pageWidthDxa));
        pageSize.setAttributeNS(W_NS, "w:h", String.valueOf(pageHeightDxa));
        sectionNode.appendChild(pageSize);

        Element pageMargin = sectionNode.getOwnerDocument().createElementNS(W_NS, "w:pgMar");
        pageMargin.setAttributeNS(W_NS, "w:top", String.valueOf(margin.getTopDxa()));
        pageMargin.setAttributeNS(W_NS, "w:right", String.valueOf(margin.getRightDxa()));
        pageMargin.setAttributeNS(W_NS, "w:bottom", String.valueOf(margin.getBottomDxa()));
        pageMargin.setAttributeNS(W_NS, "w:left", String.valueOf(margin.getLeftDxa()));
        pageMargin.setAttributeNS(W_NS, "w:header", String.valueOf(margin.getHeaderDxa()));
        pageMargin.setAttributeNS(W_NS, "w:footer", String.valueOf(margin.getFooterDxa()));
        pageMargin.setAttributeNS(W_NS, "w:gutter", String.valueOf(margin.getGutterDxa()));
        sectionNode.appendChild(pageMargin);
    }

    /**
     * 添加居中标题。
     *
     * @param document 文档对象
     * @param text 标题文本
     * @param font 字体名称
     * @param fontSize 字号
     */
    public static void addCenteredTitle(XWPFDocument document, String text, String font, int fontSize) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        applyFont(run, font, fontSize);
        run.setText(text);
    }

    /**
     * 设置运行对象字体。
     *
     * @param run 运行对象
     * @param font 字体名称
     * @param fontSize 字号
     */
    public static void applyFont(XWPFRun run, String font, int fontSize) {
        run.setFontFamily(font);
        run.setFontSize(fontSize);

        CTFonts fonts = getOrCreateFonts(run);
        fonts.setAscii(font);
        fonts.setHAnsi(font);
        fonts.setEastAsia(font);
        fonts.setCs(font);
    }

    private static void removeChild(Node parent, String localName) {
        for (int i = parent.getChildNodes().getLength() - 1; i >= 0; i--) {
            Node child = parent.getChildNodes().item(i);
            if (localName.equals(child.getLocalName())) {
                parent.removeChild(child);
            }
        }
    }

    private static CTFonts getOrCreateFonts(XWPFRun run) {
        if (!run.getCTR().isSetRPr()) {
            return run.getCTR().addNewRPr().addNewRFonts();
        }

        if (!run.getCTR().getRPr().isSetRFonts()) {
            return run.getCTR().getRPr().addNewRFonts();
        }

        return run.getCTR().getRPr().getRFonts();
    }

    /**
     * 页面边距配置。
     */
    public static final class PageMargin {
        private final int topDxa;
        private final int rightDxa;
        private final int bottomDxa;
        private final int leftDxa;
        private final int headerDxa;
        private final int footerDxa;
        private final int gutterDxa;

        public PageMargin(int topDxa, int rightDxa, int bottomDxa, int leftDxa,
                          int headerDxa, int footerDxa, int gutterDxa) {
            this.topDxa = topDxa;
            this.rightDxa = rightDxa;
            this.bottomDxa = bottomDxa;
            this.leftDxa = leftDxa;
            this.headerDxa = headerDxa;
            this.footerDxa = footerDxa;
            this.gutterDxa = gutterDxa;
        }

        public int getTopDxa() {
            return topDxa;
        }

        public int getRightDxa() {
            return rightDxa;
        }

        public int getBottomDxa() {
            return bottomDxa;
        }

        public int getLeftDxa() {
            return leftDxa;
        }

        public int getHeaderDxa() {
            return headerDxa;
        }

        public int getFooterDxa() {
            return footerDxa;
        }

        public int getGutterDxa() {
            return gutterDxa;
        }
    }
}

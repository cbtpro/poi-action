package com.chenbitao.word.docx.util;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.math.BigInteger;
import java.util.Arrays;

import static com.chenbitao.word.docx.util.DocxFixedTable.empty;
import static com.chenbitao.word.docx.util.DocxFixedTable.image;
import static com.chenbitao.word.docx.util.DocxFixedTable.row;
import static com.chenbitao.word.docx.util.DocxFixedTable.text;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public class DocxFixedTableTest {

    @Test
    public void renderWritesFixedGridSpansVerticalMergeTextAndImage() throws Exception {
        XWPFDocument document = new XWPFDocument();
        int[] widths = {1000, 2000, 3000};
        DocxFixedTable.Options options = new DocxFixedTable.Options(widths)
                .tableWidthDxa(6000)
                .cellMargins(0, 120, 0, 120)
                .font("宋体", 12);

        XWPFTable table = DocxFixedTable.render(document, options, Arrays.asList(
                row(500,
                        text("姓名"),
                        text("张三").center(),
                        image(pngBytes(), Units.toEMU(20), Units.toEMU(30)).vRestart()),
                row(text("跨列"), text("跨列内容").span(2)),
                row(text("出生年月"), text("2008-01"), empty().vContinue())
        ));

        assertEquals(BigInteger.valueOf(6000), table.getCTTbl().getTblPr().getTblW().getW());
        assertEquals(3, table.getCTTbl().getTblGrid().sizeOfGridColArray());
        assertEquals(BigInteger.valueOf(1000), table.getCTTbl().getTblGrid().getGridColArray(0).getW());
        assertEquals(3, table.getRow(0).getTableCells().size());
        assertEquals("姓名", table.getRow(0).getCell(0).getText());
        assertEquals("张三", table.getRow(0).getCell(1).getText());
        assertEquals(BigInteger.valueOf(2),
                table.getRow(1).getCell(1).getCTTc().getTcPr().getGridSpan().getVal());
        assertNotNull(table.getRow(0).getCell(2).getCTTc().getTcPr().getVMerge());
        assertNotNull(table.getRow(2).getCell(2).getCTTc().getTcPr().getVMerge());
        assertEquals(1, document.getAllPictures().size());
    }

    private byte[] pngBytes() throws Exception {
        BufferedImage image = new BufferedImage(10, 10, BufferedImage.TYPE_INT_RGB);
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
    }
}

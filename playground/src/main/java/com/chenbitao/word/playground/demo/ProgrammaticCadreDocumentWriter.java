package com.chenbitao.word.playground.demo;

import com.chenbitao.word.docx.TemplateWordGenerator;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblCellMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigInteger;
import java.net.URI;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Base64;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * 干部任免审批表的编程式写入器。
 * 这里承载的是demo业务版式，公共Word生成器只保留通用能力。
 */
final class ProgrammaticCadreDocumentWriter {

    /** 字体名称：宋体 */
    private static final String FONT = "宋体";

    /** Word XML 命名空间 */
    private static final String W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /** 表格总宽度（DXA单位） */
    private static final int TABLE_WIDTH_DXA = 9752;

    /** 表格列宽度数组（DXA单位） */
    private static final int[] COLUMN_WIDTHS = {986, 1434, 1132, 1384, 1500, 1384, 1932};

    /**
     * 生成文档。
     *
     * @param data 业务数据
     * @param output 输出路径
     * @throws Exception 如果文档生成失败
     */
    void write(Map<String, Object> data, Path output) throws Exception {
        try (XWPFDocument document = new XWPFDocument()) {
            setupPage(document);
            addTitle(document);
            addMainTable(document, data);
            addReporter(document);

            try (ByteArrayOutputStream buffer = new ByteArrayOutputStream()) {
                document.write(buffer);
                Files.write(output, buffer.toByteArray());
            }
        }
    }

    /**
     * 设置页面格式（页面大小和边距）。
     *
     * @param document Word文档对象
     */
    private void setupPage(XWPFDocument document) {
        CTBody body = document.getDocument().getBody();
        CTSectPr section = body.isSetSectPr() ? body.getSectPr() : body.addNewSectPr();
        Node sectionNode = section.getDomNode();
        removeChild(sectionNode, "pgSz");
        removeChild(sectionNode, "pgMar");

        Element pageSize = sectionNode.getOwnerDocument().createElementNS(W_NS, "w:pgSz");
        pageSize.setAttributeNS(W_NS, "w:w", "11906");
        pageSize.setAttributeNS(W_NS, "w:h", "16838");
        sectionNode.appendChild(pageSize);

        Element pageMargin = sectionNode.getOwnerDocument().createElementNS(W_NS, "w:pgMar");
        pageMargin.setAttributeNS(W_NS, "w:top", "1440");
        pageMargin.setAttributeNS(W_NS, "w:right", "1800");
        pageMargin.setAttributeNS(W_NS, "w:bottom", "1440");
        pageMargin.setAttributeNS(W_NS, "w:left", "1800");
        pageMargin.setAttributeNS(W_NS, "w:header", "851");
        pageMargin.setAttributeNS(W_NS, "w:footer", "992");
        pageMargin.setAttributeNS(W_NS, "w:gutter", "0");
        sectionNode.appendChild(pageMargin);
    }

    /**
     * 从父节点中移除指定名称的子节点。
     *
     * @param parent 父节点
     * @param localName 子节点本地名称
     */
    private void removeChild(Node parent, String localName) {
        for (int i = parent.getChildNodes().getLength() - 1; i >= 0; i--) {
            Node child = parent.getChildNodes().item(i);
            if (localName.equals(child.getLocalName())) {
                parent.removeChild(child);
            }
        }
    }

    /**
     * 添加文档标题。
     *
     * @param document Word文档对象
     */
    private void addTitle(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        applyFont(run, 22);
        run.setText("干部任免审批表");
    }

    /**
     * 添加主表格到文档。
     *
     * @param document Word文档对象
     * @param data 业务数据
     * @throws Exception 如果表格创建失败
     */
    private void addMainTable(XWPFDocument document, Map<String, Object> data) throws Exception {
        List<RowSpec> rows = rows(data);
        XWPFTable table = document.createTable();
        configureTable(table);

        for (int i = 0; i < rows.size(); i++) {
            RowSpec rowSpec = rows.get(i);
            XWPFTableRow row = prepareRow(table, i, rowSpec.cells.size());
            if (rowSpec.heightDxa > 0) {
                setRowHeight(row, rowSpec.heightDxa);
            }

            int columnIndex = 0;
            for (int j = 0; j < rowSpec.cells.size(); j++) {
                CellSpec cellSpec = rowSpec.cells.get(j);
                XWPFTableCell cell = row.getCell(j);
                int width = width(columnIndex, cellSpec.colSpan);
                configureCell(cell, cellSpec, width);
                columnIndex += cellSpec.colSpan;
            }
        }
    }

    /**
     * 构建表格行规格列表。
     *
     * @param data 业务数据
     * @return 行规格列表
     * @throws Exception 如果行构建失败
     */
    private List<RowSpec> rows(Map<String, Object> data) throws Exception {
        List<RowSpec> rows = new ArrayList<>();

        rows.add(row(981,
                text("姓名"), text(text(data, "nameSc")),
                text("性别"), text(text(data, "sex")),
                text("出生年月（岁）"), text(text(data, "birthday") + "\n" + text(data, "age")),
                photoCell(data).vRestart()));
        rows.add(row(882,
                text("民族"), text(text(data, "nationality")),
                text("籍贯"), text(text(data, "nativePlace")),
                text("出生地"), text(text(data, "birthPlace")),
                empty().vContinue()));
        rows.add(row(998,
                text("入党时间"), text(text(data, "joinPartyDate")),
                text("参加工作时间"), text(text(data, "firstJobDate")),
                text("健康状况"), text(text(data, "healthCondition")),
                empty().vContinue()));
        rows.add(row(1231,
                text("专业技术职务"), text(text(data, "qualificationName")).span(2),
                text("熟悉专业有何专长"), text(text(data, "majorExpertise")).span(2),
                empty().vContinue()));

        addEducationRows(rows, data);
        rows.add(row(text("现任职务").span(1), text(text(data, "currentPosition")).span(6)));
        rows.add(row(text("拟任职务").span(1), text(text(data, "proposedPosition")).span(6)));
        rows.add(row(text("拟免职务").span(1), text(text(data, "proposedRemovedPosition")).span(6)));
        rows.add(row(text("简历"), text(joinLines(toStringList(data.get("workExperience")))).span(6)));
        rows.add(row(text("奖惩情况"), text(text(data, "rewardPunishmentRecord")).span(6)));
        rows.add(row(text("年度考核结果"), text(joinLines(toStringList(data.get("annualAssessmentResult")))).span(6)));
        rows.add(row(text("任免理由"), text(text(data, "appointmentRemovalReason")).span(6)));
        addFamilyRows(rows, data);
        rows.add(row(text("呈报单位"), text(text(data, "reportingUnit")).span(6)));
        rows.add(row(1991,
                text("审批机关意见"),
                text("\n\n\n（盖章）\n年   月   日").span(2).center(),
                text("行政机关任免意见"),
                text("\n\n\n（盖章）\n       年   月   日").span(3).center()));

        return rows;
    }

    /**
     * 添加学历学位行到表格。
     *
     * @param rows 行规格列表
     * @param data 业务数据
     */
    private void addEducationRows(List<RowSpec> rows, Map<String, Object> data) {
        List<Map<?, ?>> educationList = toMapList(data.get("education"));
        int count = Math.max(1, educationList.size());

        for (int i = 0; i < count; i++) {
            Map<?, ?> education = mapAt(educationList, i);
            rows.add(row(
                    text(i == 0 ? "学历学位" : "").verticalMerge(i == 0),
                    text(mapText(education, "type")),
                    text(mapText(education, "degree")).span(2),
                    text(i == 0 ? "毕业院校系及专业" : "").verticalMerge(i == 0),
                    text(mapText(education, "department") + mapText(education, "major")).span(2)
            ));
        }
    }

    /**
     * 添加家庭成员行到表格。
     *
     * @param rows 行规格列表
     * @param data 业务数据
     */
    private void addFamilyRows(List<RowSpec> rows, Map<String, Object> data) {
        List<Map<?, ?>> relations = toMapList(data.get("familyAndSocialRelations"));
        int relationRows = Math.max(1, relations.size());
        rows.add(row(
                text("家庭主要成员及重要社会关系").vRestart(),
                text("称谓"), text("姓名"), text("年龄"), text("政治面貌"), text("工作单位及职务").span(2)
        ));

        for (int i = 0; i < relationRows; i++) {
            Map<?, ?> relation = mapAt(relations, i);
            rows.add(row(
                    empty().vContinue(),
                    text(mapText(relation, "appellation")),
                    text(mapText(relation, "name")),
                    text(mapText(relation, "age")),
                    text(mapText(relation, "political")),
                    text(mapText(relation, "workUnit")).span(2)
            ));
        }
    }

    /**
     * 创建照片单元格规格。
     *
     * @param data 业务数据
     * @return 单元格规格
     * @throws Exception 如果图片处理失败
     */
    private CellSpec photoCell(Map<String, Object> data) throws Exception {
        Object value = data.get("photo");
        if (!(value instanceof TemplateWordGenerator.Picture)) {
            return empty();
        }

        TemplateWordGenerator.Picture picture = (TemplateWordGenerator.Picture) value;
        if (picture.getSource() == null || picture.getWidthEmu() == null || picture.getHeightEmu() == null) {
            return empty();
        }
        return image(toPngBytes(picture.getSource()), picture.getWidthEmu(), picture.getHeightEmu()).center();
    }

    /**
     * 配置主表格的整体属性。
     *
     * @param table 目标表格
     */
    /**
     * 初始化并配置主表格的属性、宽度、边框、单元格边距和网格列。
     *
     * @param table 目标表格
     */
    private void configureTable(XWPFTable table) {
        CTTbl ctTbl = table.getCTTbl();
        CTTblPr tblPr = getOrCreateTableProperties(ctTbl);

        configureTableWidth(tblPr);
        configureTableLayout(tblPr);
        configureTableBorders(tblPr);
        configureTableCellMargins(tblPr);
        configureTableGrid(ctTbl);
    }

    /**
     * 获取或创建表格属性对象。
     *
     * @param ctTbl 表格 CT 对象
     * @return 表格属性对象
     */
    private CTTblPr getOrCreateTableProperties(CTTbl ctTbl) {
        return ctTbl.getTblPr() == null ? ctTbl.addNewTblPr() : ctTbl.getTblPr();
    }

    /**
     * 设置表格宽度为预定义值。
     *
     * @param tblPr 表格属性对象
     */
    private void configureTableWidth(CTTblPr tblPr) {
        CTTblWidth width = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        width.setW(BigInteger.valueOf(TABLE_WIDTH_DXA));
        width.setType(STTblWidth.DXA);
    }

    /**
     * 设置表格的布局模式为固定宽度。
     *
     * @param tblPr 表格属性对象
     */
    private void configureTableLayout(CTTblPr tblPr) {
        CTTblLayoutType layout = tblPr.isSetTblLayout() ? tblPr.getTblLayout() : tblPr.addNewTblLayout();
        layout.setType(STTblLayoutType.FIXED);
    }

    /**
     * 为表格添加边框样式。
     *
     * @param tblPr 表格属性对象
     */
    private void configureTableBorders(CTTblPr tblPr) {
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();

        setBorder(borders.isSetTop() ? borders.getTop() : borders.addNewTop());
        setBorder(borders.isSetLeft() ? borders.getLeft() : borders.addNewLeft());
        setBorder(borders.isSetBottom() ? borders.getBottom() : borders.addNewBottom());
        setBorder(borders.isSetRight() ? borders.getRight() : borders.addNewRight());
        setBorder(borders.isSetInsideH() ? borders.getInsideH() : borders.addNewInsideH());
        setBorder(borders.isSetInsideV() ? borders.getInsideV() : borders.addNewInsideV());
    }

    /**
     * 设置表格单元格的默认边距。
     *
     * @param tblPr 表格属性对象
     */
    private void configureTableCellMargins(CTTblPr tblPr) {
        CTTblCellMar cellMar = tblPr.isSetTblCellMar() ? tblPr.getTblCellMar() : tblPr.addNewTblCellMar();

        setCellMargin(cellMar.isSetTop() ? cellMar.getTop() : cellMar.addNewTop(), 0);
        setCellMargin(cellMar.isSetBottom() ? cellMar.getBottom() : cellMar.addNewBottom(), 0);
        setCellMargin(cellMar.isSetLeft() ? cellMar.getLeft() : cellMar.addNewLeft(), 108);
        setCellMargin(cellMar.isSetRight() ? cellMar.getRight() : cellMar.addNewRight(), 108);
    }

    /**
     * 配置表格网格列宽度。
     *
     * @param ctTbl 表格 CT 对象
     */
    private void configureTableGrid(CTTbl ctTbl) {
        CTTblGrid grid = ctTbl.getTblGrid() == null ? ctTbl.addNewTblGrid() : ctTbl.getTblGrid();

        // Clear existing columns
        while (grid.sizeOfGridColArray() > 0) {
            grid.removeGridCol(grid.sizeOfGridColArray() - 1);
        }

        // Add new columns with specified widths
        for (int columnWidth : COLUMN_WIDTHS) {
            CTTblGridCol column = grid.addNewGridCol();
            column.setW(BigInteger.valueOf(columnWidth));
        }
    }

    /**
     * 为单元格应用规格并写入内容。
     *
     * @param cell 单元格对象
     * @param spec 单元格规格
     * @param width 单元格宽度（DXA）
     */
    private void configureCell(XWPFTableCell cell, CellSpec spec, int width) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr()
                ? cell.getCTTc().getTcPr()
                : cell.getCTTc().addNewTcPr();

        CTTblWidth tcW = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        tcW.setW(BigInteger.valueOf(width));
        tcW.setType(STTblWidth.DXA);

        if (spec.colSpan > 1) {
            tcPr.addNewGridSpan().setVal(BigInteger.valueOf(spec.colSpan));
        }
        if (spec.merge != Merge.NONE) {
            CTVMerge merge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
            merge.setVal(spec.merge == Merge.RESTART ? STMerge.RESTART : STMerge.CONTINUE);
        }

        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph paragraph = resetParagraph(cell);
        paragraph.setAlignment(spec.alignment);
        if (spec.imageBytes == null) {
            writeText(paragraph.createRun(), spec.text);
        } else {
            writeImage(paragraph.createRun(), spec);
        }
    }

    /**
     * 将文本写入段落，并保持换行。
     *
     * @param run 运行对象
     * @param text 文本内容
     */
    private void writeText(XWPFRun run, String text) {
        applyFont(run, 12);
        String[] lines = text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                run.addBreak();
            }
            run.setText(lines[i]);
        }
    }

    /**
     * 将图片字节流插入到单元格中。
     *
     * @param run 运行对象
     * @param spec 图片单元格规格
     */
    private void writeImage(XWPFRun run, CellSpec spec) {
        try {
            run.addPicture(new ByteArrayInputStream(spec.imageBytes),
                    Document.PICTURE_TYPE_PNG,
                    "photo.png",
                    spec.imageWidthEmu,
                    spec.imageHeightEmu);
        } catch (Exception e) {
            throw new IllegalStateException("照片插入失败", e);
        }
    }

    /**
     * 重置单元格中的段落，确保使用单个空段落进行写入。
     *
     * @param cell 单元格对象
     * @return 可写入的段落对象
     */
    private XWPFParagraph resetParagraph(XWPFTableCell cell) {
        while (cell.getParagraphs().size() > 1) {
            cell.removeParagraph(1);
        }

        XWPFParagraph paragraph = cell.getParagraphs().isEmpty()
                ? cell.addParagraph()
                : cell.getParagraphs().get(0);
        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
        return paragraph;
    }

    /**
     * 确保表格行包含正确数量的单元格。
     *
     * @param table 表格对象
     * @param rowIndex 行索引
     * @param cellCount 目标单元格数量
     * @return 已准备好的表格行
     */
    private XWPFTableRow prepareRow(XWPFTable table, int rowIndex, int cellCount) {
        XWPFTableRow row = rowIndex == 0 ? table.getRow(0) : table.createRow();
        while (row.getTableCells().size() > cellCount) {
            int position = row.getTableCells().size() - 1;
            row.removeCell(position);
            row.getCtRow().removeTc(position);
        }
        while (row.getTableCells().size() < cellCount) {
            row.addNewTableCell();
        }
        return row;
    }

    /**
     * 设置行高。
     *
     * @param row 行对象
     * @param heightDxa 行高，DXA 单位
     */
    private void setRowHeight(XWPFTableRow row, int heightDxa) {
        CTTrPr trPr = row.getCtRow().isSetTrPr() ? row.getCtRow().getTrPr() : row.getCtRow().addNewTrPr();
        CTHeight height = trPr.sizeOfTrHeightArray() > 0 ? trPr.getTrHeightArray(0) : trPr.addNewTrHeight();
        height.setVal(BigInteger.valueOf(heightDxa));
        height.setHRule(STHeightRule.AT_LEAST);
    }

    /**
     * 设置边框样式。
     *
     * @param border 边框对象
     */
    private void setBorder(CTBorder border) {
        border.setVal(STBorder.SINGLE);
        border.setColor("auto");
        border.setSz(BigInteger.valueOf(4));
        border.setSpace(BigInteger.ZERO);
    }

    /**
     * 设置单元格边距属性。
     *
     * @param margin 单元格宽度对象
     * @param width 边距宽度（DXA）
     */
    private void setCellMargin(CTTblWidth margin, int width) {
        margin.setW(BigInteger.valueOf(width));
        margin.setType(STTblWidth.DXA);
    }

    /**
     * 应用默认字体样式到运行对象。
     *
     * @param run 运行对象
     * @param fontSize 字号
     */
    private void applyFont(XWPFRun run, int fontSize) {
        run.setFontFamily(FONT);
        run.setFontSize(fontSize);

        CTFonts fonts = getOrCreateFonts(run);
        fonts.setAscii(FONT);
        fonts.setHAnsi(FONT);
        fonts.setEastAsia(FONT);
        fonts.setCs(FONT);
    }

    /**
     * 获取或创建运行对象的字体配置。
     *
     * @param run 运行对象
     * @return 字体配置对象
     */
    private CTFonts getOrCreateFonts(XWPFRun run) {
        if (!run.getCTR().isSetRPr()) {
            return run.getCTR().addNewRPr().addNewRFonts();
        }

        if (!run.getCTR().getRPr().isSetRFonts()) {
            return run.getCTR().getRPr().addNewRFonts();
        }

        return run.getCTR().getRPr().getRFonts();
    }

    /**
     * 添加填表人信息段落。
     *
     * @param document 文档对象
     */
    private void addReporter(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        applyFont(run, 12);
        run.setText("填表人：");
        document.createParagraph();
    }

    /**
     * 计算跨列单元格的总宽度。
     *
     * @param startColumn 起始列号
     * @param colSpan 列跨度
     * @return 宽度（DXA）
     */
    private int width(int startColumn, int colSpan) {
        int width = 0;
        for (int i = startColumn; i < startColumn + colSpan && i < COLUMN_WIDTHS.length; i++) {
            width += COLUMN_WIDTHS[i];
        }
        return width;
    }

    /**
     * 将图片来源转换为 PNG 字节数组。
     *
     * @param source 图片来源
     * @return PNG 图像字节数组
     * @throws Exception 转换失败时抛出
     */
    private byte[] toPngBytes(Object source) throws Exception {
        byte[] bytes = readBytes(source);
        BufferedImage image = ImageIO.read(new ByteArrayInputStream(bytes));
        if (image == null) {
            return bytes;
        }

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
    }

    /**
     * 从多种来源读取图片字节。
     *
     * @param source 图片来源
     * @return 图片字节数组
     * @throws Exception 读取失败时抛出
     */
    private byte[] readBytes(Object source) throws Exception {
        if (source instanceof byte[]) {
            return (byte[]) source;
        }
        if (source instanceof InputStream) {
            return readAll((InputStream) source);
        }
        if (source instanceof File) {
            return readFileBytes((File) source);
        }
        if (source instanceof Path) {
            return readPathBytes((Path) source);
        }
        if (source instanceof URL) {
            return readUrlBytes((URL) source);
        }
        if (source instanceof URI) {
            return readUriBytes((URI) source);
        }
        if (source instanceof String) {
            return readStringSource((String) source);
        }
        throw new IllegalArgumentException("不支持的图片输入类型：" + source.getClass().getName());
    }

    /**
     * 从文件读取字节数组。
     *
     * @param file 文件对象
     * @return 文件字节数组
     * @throws IOException 读取失败时抛出
     */
    private byte[] readFileBytes(File file) throws IOException {
        return Files.readAllBytes(file.toPath());
    }

    /**
     * 从路径读取字节数组。
     *
     * @param path 路径对象
     * @return 字节数组
     * @throws IOException 读取失败时抛出
     */
    private byte[] readPathBytes(Path path) throws IOException {
        return Files.readAllBytes(path);
    }

    /**
     * 从 URL 读取字节数组。
     *
     * @param url URL 对象
     * @return 字节数组
     * @throws IOException 读取失败时抛出
     */
    private byte[] readUrlBytes(URL url) throws IOException {
        try (InputStream inputStream = url.openStream()) {
            return readAll(inputStream);
        }
    }

    /**
     * 从 URI 读取字节数组。
     *
     * @param uri URI 对象
     * @return 字节数组
     * @throws IOException 读取失败时抛出
     */
    private byte[] readUriBytes(URI uri) throws IOException {
        return readUrlBytes(uri.toURL());
    }

    /**
     * 读取字符串形式的图片来源，支持 URL、路径和 Base64。
     *
     * @param source 图片来源字符串
     * @return 图片字节数组
     * @throws Exception 解析或读取失败时抛出
     */
    private byte[] readStringSource(String source) throws Exception {
        String value = source == null ? "" : source.trim();
        if (value.isEmpty()) {
            return new byte[0];
        }
        if (value.startsWith("http://") || value.startsWith("https://")) {
            return readBytes(new URL(value));
        }

        Path path = Paths.get(value);
        if (Files.exists(path)) {
            return Files.readAllBytes(path);
        }

        int commaIndex = value.indexOf(',');
        String base64 = commaIndex >= 0 ? value.substring(commaIndex + 1) : value;
        return Base64.getDecoder().decode(base64);
    }

    /**
     * 读取输入流中的所有字节。
     *
     * @param inputStream 输入流
     * @return 字节数组
     * @throws IOException 读取失败时抛出
     */
    private byte[] readAll(InputStream inputStream) throws IOException {
        try (InputStream in = inputStream; ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[8192];
            int read;
            while ((read = in.read(buffer)) != -1) {
                output.write(buffer, 0, read);
            }
            return output.toByteArray();
        }
    }

    /**
     * 从数据映射中读取字符串值。
     *
     * @param data 数据映射
     * @param key 键名
     * @return 字符串值或空字符串
     */
    private String text(Map<String, Object> data, String key) {
        Object value = data.get(key);
        return value == null ? "" : value.toString();
    }

    /**
     * 从通用映射中读取字符串值。
     *
     * @param data 数据映射
     * @param key 键名
     * @return 字符串值或空字符串
     */
    private String mapText(Map<?, ?> data, String key) {
        Object value = data.get(key);
        return value == null ? "" : value.toString();
    }

    /**
     * 将可迭代对象中的映射元素收集为列表。
     *
     * @param value 可迭代对象
     * @return 映射列表
     */
    private List<Map<?, ?>> toMapList(Object value) {
        List<Map<?, ?>> result = new ArrayList<>();
        if (!(value instanceof Iterable<?>)) {
            return result;
        }

        for (Object item : (Iterable<?>) value) {
            if (item instanceof Map<?, ?>) {
                result.add((Map<?, ?>) item);
            }
        }
        return result;
    }

    /**
     * 获取列表中指定索引的映射对象。
     *
     * @param list 映射列表
     * @param index 索引
     * @return 映射对象或空映射
     */
    private Map<?, ?> mapAt(List<Map<?, ?>> list, int index) {
        return index < list.size() ? list.get(index) : Collections.emptyMap();
    }

    /**
     * 将可迭代对象转换为字符串列表。
     *
     * @param value 可迭代对象
     * @return 字符串列表
     */
    private List<String> toStringList(Object value) {
        List<String> result = new ArrayList<>();
        if (!(value instanceof Iterable<?>)) {
            return result;
        }

        for (Object item : (Iterable<?>) value) {
            result.add(item == null ? "" : item.toString());
        }
        return result;
    }

    /**
     * 将字符串列表拼接为多行文本。
     *
     * @param values 字符串列表
     * @return 多行文本
     */
    private String joinLines(List<String> values) {
        StringBuilder text = new StringBuilder();
        for (String value : values) {
            if (text.length() > 0) {
                text.append('\n');
            }
            text.append(value == null ? "" : value);
        }
        return text.toString();
    }

    /**
     * 构造默认行规格。
     *
     * @param cells 单元格规格
     * @return 行规格对象
     */
    private RowSpec row(CellSpec... cells) {
        return row(0, cells);
    }

    /**
     * 构造带高度的行规格。
     *
     * @param heightDxa 行高度（DXA）
     * @param cells 单元格规格
     * @return 行规格对象
     */
    private RowSpec row(int heightDxa, CellSpec... cells) {
        return new RowSpec(heightDxa, Arrays.asList(cells));
    }

    /**
     * 创建文本单元格规格。
     *
     * @param text 文本内容
     * @return 单元格规格
     */
    private CellSpec text(String text) {
        return new CellSpec(text, null, 0, 0);
    }

    /**
     * 创建空单元格规格。
     *
     * @return 单元格规格
     */
    private CellSpec empty() {
        return text("");
    }

    /**
     * 创建图片单元格规格。
     *
     * @param imageBytes 图片字节数组
     * @param widthEmu 图片宽度（EMU）
     * @param heightEmu 图片高度（EMU）
     * @return 单元格规格
     */
    private CellSpec image(byte[] imageBytes, int widthEmu, int heightEmu) {
        return new CellSpec("", imageBytes, widthEmu, heightEmu);
    }

    private enum Merge {
        NONE,
        RESTART,
        CONTINUE
    }

    private static final class RowSpec {
        private final int heightDxa;
        private final List<CellSpec> cells;

        private RowSpec(int heightDxa, List<CellSpec> cells) {
            this.heightDxa = heightDxa;
            this.cells = cells;
        }
    }

    private static final class CellSpec {
        private final String text;
        private final byte[] imageBytes;
        private final int imageWidthEmu;
        private final int imageHeightEmu;
        private int colSpan = 1;
        private Merge merge = Merge.NONE;
        private ParagraphAlignment alignment = ParagraphAlignment.LEFT;

        private CellSpec(String text, byte[] imageBytes, int imageWidthEmu, int imageHeightEmu) {
            this.text = text == null ? "" : text;
            this.imageBytes = imageBytes;
            this.imageWidthEmu = imageWidthEmu;
            this.imageHeightEmu = imageHeightEmu;
        }

        private CellSpec span(int colSpan) {
            this.colSpan = Math.max(1, colSpan);
            return this;
        }

        private CellSpec vRestart() {
            this.merge = Merge.RESTART;
            return this;
        }

        private CellSpec vContinue() {
            this.merge = Merge.CONTINUE;
            return this;
        }

        private CellSpec verticalMerge(boolean restart) {
            this.merge = restart ? Merge.RESTART : Merge.CONTINUE;
            return this;
        }

        private CellSpec center() {
            this.alignment = ParagraphAlignment.CENTER;
            return this;
        }
    }
}

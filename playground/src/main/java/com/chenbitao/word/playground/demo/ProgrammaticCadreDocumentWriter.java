package com.chenbitao.word.playground.demo;

import com.chenbitao.word.docx.TemplateWordGenerator;
import com.chenbitao.word.docx.util.DocxFixedTable;
import com.chenbitao.word.docx.util.DocxPageUtils;
import com.chenbitao.word.docx.util.ImageSourceUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static com.chenbitao.word.docx.util.DocxFixedTable.empty;
import static com.chenbitao.word.docx.util.DocxFixedTable.image;
import static com.chenbitao.word.docx.util.DocxFixedTable.row;
import static com.chenbitao.word.docx.util.DocxFixedTable.text;

/**
 * 干部任免审批表的编程式写入器。
 * 这里承载的是demo业务版式，公共Word生成器只保留通用能力。
 */
final class ProgrammaticCadreDocumentWriter {

    /** 字体名称：宋体 */
    private static final String FONT = "宋体";

    /** 表格总宽度（DXA单位） */
    private static final int TABLE_WIDTH_DXA = 9752;

    /** 表格列宽度数组（DXA单位） */
    private static final int[] COLUMN_WIDTHS = {986, 1434, 1132, 1384, 1500, 1384, 1932};

    /** 页面边距 */
    private static final DocxPageUtils.PageMargin PAGE_MARGIN =
            new DocxPageUtils.PageMargin(1440, 1800, 1440, 1800, 851, 992, 0);

    /** 主表渲染选项 */
    private static final DocxFixedTable.Options TABLE_OPTIONS = new DocxFixedTable.Options(COLUMN_WIDTHS)
            .tableWidthDxa(TABLE_WIDTH_DXA)
            .cellMargins(0, 108, 0, 108)
            .font(FONT, 12);

    /**
     * 生成文档。
     *
     * @param data 业务数据
     * @param output 输出路径
     * @throws IOException 如果文档生成失败
     */
    void write(Map<String, Object> data, Path output) throws IOException {
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
        DocxPageUtils.setupPage(document, 11906, 16838, PAGE_MARGIN);
    }

    /**
     * 添加文档标题。
     *
     * @param document Word文档对象
     */
    private void addTitle(XWPFDocument document) {
        DocxPageUtils.addCenteredTitle(document, "干部任免审批表", FONT, 22);
    }

    /**
     * 添加主表格到文档。
     *
     * @param document Word文档对象
     * @param data 业务数据
     * @throws IOException 如果表格创建失败
     */
    private void addMainTable(XWPFDocument document, Map<String, Object> data) throws IOException {
        DocxFixedTable.render(document, TABLE_OPTIONS, rows(data));
    }

    /**
     * 构建表格行规格列表。
     *
     * @param data 业务数据
     * @return 行规格列表
     * @throws IOException 如果行构建失败
     */
    private List<DocxFixedTable.Row> rows(Map<String, Object> data) throws IOException {
        List<DocxFixedTable.Row> rows = new ArrayList<>();

        rows.add(row(981,
                text("姓名"), text(dataText(data, "nameSc")),
                text("性别"), text(dataText(data, "sex")),
                text("出生年月（岁）"), text(dataText(data, "birthday") + "\n" + dataText(data, "age")),
                photoCell(data).vRestart()));
        rows.add(row(882,
                text("民族"), text(dataText(data, "nationality")),
                text("籍贯"), text(dataText(data, "nativePlace")),
                text("出生地"), text(dataText(data, "birthPlace")),
                empty().vContinue()));
        rows.add(row(998,
                text("入党时间"), text(dataText(data, "joinPartyDate")),
                text("参加工作时间"), text(dataText(data, "firstJobDate")),
                text("健康状况"), text(dataText(data, "healthCondition")),
                empty().vContinue()));
        rows.add(row(1231,
                text("专业技术职务"), text(dataText(data, "qualificationName")).span(2),
                text("熟悉专业有何专长"), text(dataText(data, "majorExpertise")).span(2),
                empty().vContinue()));

        addEducationRows(rows, data);
        rows.add(row(text("现任职务").span(1), text(dataText(data, "currentPosition")).span(6)));
        rows.add(row(text("拟任职务").span(1), text(dataText(data, "proposedPosition")).span(6)));
        rows.add(row(text("拟免职务").span(1), text(dataText(data, "proposedRemovedPosition")).span(6)));
        rows.add(row(text("简历"), text(joinLines(toStringList(data.get("workExperience")))).span(6)));
        rows.add(row(text("奖惩情况"), text(dataText(data, "rewardPunishmentRecord")).span(6)));
        rows.add(row(text("年度考核结果"), text(joinLines(toStringList(data.get("annualAssessmentResult")))).span(6)));
        rows.add(row(text("任免理由"), text(dataText(data, "appointmentRemovalReason")).span(6)));
        addFamilyRows(rows, data);
        rows.add(row(text("呈报单位"), text(dataText(data, "reportingUnit")).span(6)));
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
    private void addEducationRows(List<DocxFixedTable.Row> rows, Map<String, Object> data) {
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
    private void addFamilyRows(List<DocxFixedTable.Row> rows, Map<String, Object> data) {
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
     * @throws IOException 如果图片处理失败
     */
    private DocxFixedTable.Cell photoCell(Map<String, Object> data) throws IOException {
        Object value = data.get("photo");
        if (!(value instanceof TemplateWordGenerator.Picture)) {
            return empty();
        }

        TemplateWordGenerator.Picture picture = (TemplateWordGenerator.Picture) value;
        if (picture.getSource() == null || picture.getWidthEmu() == null || picture.getHeightEmu() == null) {
            return empty();
        }
        return image(ImageSourceUtils.toPngBytes(picture.getSource()), picture.getWidthEmu(), picture.getHeightEmu()).center();
    }

    /**
     * 添加填表人信息段落。
     *
     * @param document 文档对象
     */
    private void addReporter(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        DocxPageUtils.applyFont(run, FONT, 12);
        run.setText("填表人：");
        document.createParagraph();
    }

    /**
     * 从数据映射中读取字符串值。
     *
     * @param data 数据映射
     * @param key 键名
     * @return 字符串值或空字符串
     */
    private String dataText(Map<String, Object> data, String key) {
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

}

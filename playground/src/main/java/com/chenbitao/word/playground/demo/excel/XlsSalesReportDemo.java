package com.chenbitao.word.playground.demo.excel;

import com.chenbitao.word.excel.ExcelGenerator;
import com.chenbitao.word.factory.SpreadsheetGeneratorFactory;
import lombok.extern.slf4j.Slf4j;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

/**
 * Excel 97-2003 销售报表示例。
 *
 * <p>该演示使用 {@link SpreadsheetGeneratorFactory} 创建 XLS 生成器，展示表头、
 * 明细数据、公式和自动列宽的基础用法。</p>
 */
@Slf4j
public class XlsSalesReportDemo {

    private static final String SHEET_NAME = "销售数据";
    private static final String OUTPUT_FILE_NAME = "excel-sales-demo.xls";

    /**
     * 生成 Excel 演示文件。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 如果目录创建或文件写出失败
     */
    public static void main(String[] args) throws Exception {
        Path outputPath = defaultOutputPath();
        generate(outputPath);
        log.info("Excel 演示文件生成成功：{}", outputPath.toAbsolutePath());
    }

    /**
     * 生成销售报表 XLS 文件。
     *
     * @param outputPath 输出路径
     * @throws Exception 如果目录创建或文件写出失败
     */
    public static void generate(Path outputPath) throws Exception {
        Files.createDirectories(outputPath.getParent());

        ExcelGenerator generator = SpreadsheetGeneratorFactory.get("xls");
        generator.createWorkbook();
        generator.addHeaderRow(SHEET_NAME, Arrays.asList("产品", "数量", "单价", "小计"));
        generator.addRows(SHEET_NAME, salesRows());

        // HSSF 公式使用 Excel 坐标，第一行是表头，所以明细从第 2 行开始。
        generator.setFormula(SHEET_NAME, 1, 3, "B2*C2");
        generator.setFormula(SHEET_NAME, 2, 3, "B3*C3");
        generator.setFormula(SHEET_NAME, 3, 3, "B4*C4");
        generator.setFormula(SHEET_NAME, 4, 3, "SUM(D2:D4)");
        generator.autoSizeColumns(SHEET_NAME);
        generator.save(outputPath.toString());
    }

    private static List<List<?>> salesRows() {
        return Arrays.<List<?>>asList(
                Arrays.asList("模板服务", 12, 199D),
                Arrays.asList("批量生成", 5, 699D),
                Arrays.asList("定制版式", 3, 1299D),
                Arrays.asList("合计", "", "")
        );
    }

    private static Path defaultOutputPath() {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", OUTPUT_FILE_NAME);
        }
        // 从仓库根目录执行 mvn -pl playground 时，user.dir 是根目录。
        return Paths.get("playground", "target", OUTPUT_FILE_NAME);
    }
}

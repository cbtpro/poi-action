package com.chenbitao.word.playground.demo.excel;

import com.chenbitao.word.excel.ExcelGenerator;
import com.chenbitao.word.excel.XlsxWorkbookGenerator;
import lombok.extern.slf4j.Slf4j;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Excel Open XML 大数据量销售报表示例。
 *
 * <p>该演示使用 SXSSF 流式写出 {@code .xlsx} 文件，适合验证大量明细行导出时的内存友好生成方式。</p>
 */
@Slf4j
public class XlsxLargeSalesReportDemo {

    private static final String SHEET_NAME = "销售明细";
    private static final String OUTPUT_FILE_NAME = "excel-large-sales-demo.xlsx";
    private static final int DEFAULT_ROW_COUNT = 1000;
    private static final int STREAMING_WINDOW_SIZE = 100;

    /**
     * 生成 XLSX 演示文件。
     *
     * @param args 第一个参数可指定明细行数，未指定时生成 1000 行
     * @throws Exception 如果目录创建或文件写出失败
     */
    public static void main(String[] args) throws Exception {
        int rowCount = parseRowCount(args);
        Path outputPath = defaultOutputPath();

        generate(outputPath, rowCount);

        log.info("XLSX 演示文件生成成功：{}，明细行数：{}", outputPath.toAbsolutePath(), rowCount);
    }

    /**
     * 生成大数据量销售报表 XLSX 文件。
     *
     * @param outputPath 输出路径
     * @param rowCount 明细行数
     * @throws Exception 如果目录创建或文件写出失败
     */
    public static void generate(Path outputPath, int rowCount) throws Exception {
        Files.createDirectories(outputPath.getParent());

        ExcelGenerator generator = XlsxWorkbookGenerator.streaming(STREAMING_WINDOW_SIZE);
        generator.createWorkbook();
        generator.addHeaderRow(SHEET_NAME, Arrays.asList("序号", "产品", "数量", "单价", "总计"));
        generator.addRows(SHEET_NAME, salesRows(Math.max(0, rowCount)));
        addTotalFormula(generator, rowCount);
        generator.autoSizeColumns(SHEET_NAME);
        generator.save(outputPath.toString());
    }

    private static List<List<?>> salesRows(int rowCount) {
        List<List<?>> rows = new ArrayList<>();
        for (int i = 1; i <= rowCount; i++) {
            int quantity = i % 9 + 1;
            double unitPrice = 99D + (i % 5) * 50D;
            rows.add(Arrays.asList(i, "商品-" + i, quantity, unitPrice, ""));
        }
        rows.add(Arrays.asList("", "合计", "", "", ""));
        return rows;
    }

    private static void addTotalFormula(ExcelGenerator generator, int rowCount) {
        int totalRowIndex = rowCount + 1;
        int totalExcelRow = rowCount + 2;
        // 只在最后一行写总计公式，避免 SXSSF 已刷出的明细行被二次修改。
        String formula = rowCount == 0 ? "0" : "SUMPRODUCT(C2:C" + (totalExcelRow - 1) + ",D2:D" + (totalExcelRow - 1) + ")";
        generator.setFormula(SHEET_NAME, totalRowIndex, 4, formula);
    }

    private static int parseRowCount(String[] args) {
        if (args == null || args.length == 0 || args[0] == null || args[0].trim().isEmpty()) {
            return DEFAULT_ROW_COUNT;
        }
        return Math.max(0, Integer.parseInt(args[0]));
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

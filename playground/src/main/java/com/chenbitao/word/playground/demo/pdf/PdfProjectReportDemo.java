package com.chenbitao.word.playground.demo.pdf;

import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 内容丰富的 100 页 PDF 项目报告演示。
 */
@Slf4j
public class PdfProjectReportDemo {

    private static final String OUTPUT_FILE_NAME = "pdf-report-demo.pdf";
    private static final int PAGE_COUNT = 100;
    private static final float MARGIN = 42F;
    private static final float PAGE_WIDTH = PDRectangle.LETTER.getWidth();
    private static final float PAGE_HEIGHT = PDRectangle.LETTER.getHeight();
    private static final float CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2;
    private static final DecimalFormat ONE_DECIMAL = new DecimalFormat("0.0");
    private static final DecimalFormat INTEGER = new DecimalFormat("#,##0");

    private static final Color INK = new Color(38, 50, 56);
    private static final Color MUTED = new Color(100, 116, 139);
    private static final Color BLUE = new Color(31, 78, 121);
    private static final Color TEAL = new Color(0, 128, 128);
    private static final Color GREEN = new Color(46, 125, 50);
    private static final Color AMBER = new Color(230, 126, 34);
    private static final Color RED = new Color(183, 28, 28);
    private static final Color LIGHT_BLUE = new Color(229, 241, 252);
    private static final Color LIGHT_GREEN = new Color(232, 245, 233);
    private static final Color LIGHT_AMBER = new Color(255, 243, 224);

    /**
     * 生成 PDF 演示文件。
     *
     * @param args 命令行参数，当前未使用
     * @throws Exception 目录创建、绘制或文件写入失败时抛出
     */
    public static void main(String[] args) throws Exception {
        Path outputPath = defaultOutputPath();
        generate(outputPath);
        log.info("PDF 演示文件生成完成：{}", outputPath.toAbsolutePath());
    }

    /**
     * 生成一份 100 页的真实感项目报告 PDF。
     *
     * @param outputPath 输出路径
     * @throws Exception 目录创建、绘制或文件写入失败时抛出
     */
    public static void generate(Path outputPath) throws Exception {
        PdfBoxLoggingConfigurer.configure();
        Files.createDirectories(outputPath.getParent());

        try (PDDocument document = new PDDocument()) {
            for (int i = 1; i <= PAGE_COUNT; i++) {
                ReportPage page = ReportPage.of(i);
                drawPage(document, page);
            }
            document.save(outputPath.toFile());
        }
    }

    private static void drawPage(PDDocument document, ReportPage reportPage) throws Exception {
        PDPage page = new PDPage(PDRectangle.LETTER);
        document.addPage(page);
        PDImageXObject image = LosslessFactory.createFromImage(document, reportImage(reportPage));

        try (PDPageContentStream stream = new PDPageContentStream(document, page)) {
            drawPageBackground(stream);
            drawHeader(stream, reportPage);
            drawTitleBlock(stream, reportPage);
            drawHeroImage(stream, image, reportPage);
            drawWrappedNarrative(stream, reportPage);
            drawMetricCards(stream, reportPage);
            drawTimeline(stream, reportPage);
            drawRiskPanel(stream, reportPage);
            drawFooter(stream, reportPage);
        }
    }

    private static void drawPageBackground(PDPageContentStream stream) throws Exception {
        stream.setNonStrokingColor(Color.WHITE);
        stream.addRect(0, 0, PAGE_WIDTH, PAGE_HEIGHT);
        stream.fill();

        stream.setNonStrokingColor(new Color(248, 250, 252));
        stream.addRect(0, PAGE_HEIGHT - 72F, PAGE_WIDTH, 72F);
        stream.fill();
    }

    private static void drawHeader(PDPageContentStream stream, ReportPage page) throws Exception {
        stream.setNonStrokingColor(BLUE);
        stream.addRect(MARGIN, PAGE_HEIGHT - 54F, 11F, 11F);
        stream.fill();
        text(stream, "poi-action PDF Report", MARGIN + 18F, PAGE_HEIGHT - 54F, PDType1Font.HELVETICA_BOLD, 10F, BLUE);
        text(stream, page.department + " / " + page.region, PAGE_WIDTH - MARGIN - 150F, PAGE_HEIGHT - 54F,
                PDType1Font.HELVETICA, 9F, MUTED);
    }

    private static void drawTitleBlock(PDPageContentStream stream, ReportPage page) throws Exception {
        text(stream, page.title, MARGIN, 704F, PDType1Font.HELVETICA_BOLD, 20F, INK);
        text(stream, "Quarter: " + page.quarter + "    Owner: " + page.owner + "    Status: " + page.status,
                MARGIN, 684F, PDType1Font.HELVETICA, 10F, MUTED);

        Color statusColor = page.health >= 88 ? GREEN : page.health >= 76 ? AMBER : RED;
        stream.setNonStrokingColor(statusColor);
        stream.addRect(MARGIN, 666F, Math.max(60F, page.health * 2.2F), 5F);
        stream.fill();
        text(stream, "Health score " + page.health + "/100", MARGIN, 651F, PDType1Font.HELVETICA_BOLD, 9F, statusColor);
    }

    private static void drawHeroImage(PDPageContentStream stream, PDImageXObject image, ReportPage page) throws Exception {
        float imageX = PAGE_WIDTH - MARGIN - 190F;
        float imageY = 483F;
        float imageW = 190F;
        float imageH = 128F;

        stream.setNonStrokingColor(new Color(236, 239, 241));
        stream.addRect(imageX - 5F, imageY - 5F, imageW + 10F, imageH + 10F);
        stream.fill();
        stream.drawImage(image, imageX, imageY, imageW, imageH);
        text(stream, "Figure " + page.pageNumber + ": " + page.imageCaption,
                imageX, imageY - 13F, PDType1Font.HELVETICA_OBLIQUE, 8F, MUTED);
    }

    private static void drawWrappedNarrative(PDPageContentStream stream, ReportPage page) throws Exception {
        float imageLeft = PAGE_WIDTH - MARGIN - 200F;
        float imageBottom = 470F;
        float imageTop = 620F;
        float y = 630F;

        for (String paragraph : page.paragraphs) {
            List<String> sourceLines = wrap(paragraph, PDType1Font.HELVETICA, 10.2F, CONTENT_WIDTH);
            for (String sourceLine : sourceLines) {
                float usableWidth = y <= imageTop && y >= imageBottom ? imageLeft - MARGIN - 12F : CONTENT_WIDTH;
                for (String line : wrap(sourceLine, PDType1Font.HELVETICA, 10.2F, usableWidth)) {
                    Color color = line.contains("risk") || line.contains("delayed") ? RED
                            : line.contains("improved") || line.contains("ahead") ? GREEN : INK;
                    text(stream, line, MARGIN, y, PDType1Font.HELVETICA, 10.2F, color);
                    y -= 14F;
                }
            }
            y -= 6F;
        }

        text(stream, page.callout, MARGIN, 430F, PDType1Font.HELVETICA_BOLD, 11F,
                page.callout.contains("Watch") ? AMBER : TEAL);
    }

    private static void drawMetricCards(PDPageContentStream stream, ReportPage page) throws Exception {
        float y = 365F;
        float cardW = (CONTENT_WIDTH - 20F) / 3F;
        metricCard(stream, MARGIN, y, cardW, "Revenue", "$" + ONE_DECIMAL.format(page.revenueMillions) + "M",
                page.revenueDelta >= 0 ? "+" + ONE_DECIMAL.format(page.revenueDelta) + "%" : ONE_DECIMAL.format(page.revenueDelta) + "%",
                LIGHT_BLUE, BLUE);
        metricCard(stream, MARGIN + cardW + 10F, y, cardW, "Active users", INTEGER.format(page.activeUsers),
                "+" + ONE_DECIMAL.format(page.userGrowth) + "%", LIGHT_GREEN, GREEN);
        metricCard(stream, MARGIN + (cardW + 10F) * 2F, y, cardW, "Cycle time", page.cycleDays + " days",
                page.cycleDelta <= 0 ? ONE_DECIMAL.format(page.cycleDelta) + " days" : "+" + ONE_DECIMAL.format(page.cycleDelta) + " days",
                LIGHT_AMBER, AMBER);
    }

    private static void metricCard(PDPageContentStream stream,
                                   float x,
                                   float y,
                                   float width,
                                   String label,
                                   String value,
                                   String delta,
                                   Color background,
                                   Color accent) throws Exception {
        stream.setNonStrokingColor(background);
        stream.addRect(x, y, width, 58F);
        stream.fill();
        stream.setNonStrokingColor(accent);
        stream.addRect(x, y + 54F, width, 4F);
        stream.fill();
        text(stream, label, x + 10F, y + 39F, PDType1Font.HELVETICA, 8.5F, MUTED);
        text(stream, value, x + 10F, y + 21F, PDType1Font.HELVETICA_BOLD, 15F, INK);
        text(stream, delta + " vs prior period", x + 10F, y + 8F, PDType1Font.HELVETICA, 8F, accent);
    }

    private static void drawTimeline(PDPageContentStream stream, ReportPage page) throws Exception {
        float x = MARGIN;
        float y = 280F;
        text(stream, "Delivery timeline", x, y + 38F, PDType1Font.HELVETICA_BOLD, 12F, INK);
        String[] labels = {"Discovery", "Build", "Validation", "Launch"};
        for (int i = 0; i < labels.length; i++) {
            float nodeX = x + i * 135F;
            stream.setStrokingColor(new Color(203, 213, 225));
            if (i < labels.length - 1) {
                stream.moveTo(nodeX + 16F, y + 15F);
                stream.lineTo(nodeX + 135F, y + 15F);
                stream.stroke();
            }
            stream.setNonStrokingColor(i <= page.phaseIndex ? TEAL : new Color(226, 232, 240));
            stream.addRect(nodeX, y, 30F, 30F);
            stream.fill();
            text(stream, String.valueOf(i + 1), nodeX + 10F, y + 9F, PDType1Font.HELVETICA_BOLD, 10F,
                    i <= page.phaseIndex ? Color.WHITE : MUTED);
            text(stream, labels[i], nodeX, y - 14F, PDType1Font.HELVETICA, 8.8F, INK);
        }
    }

    private static void drawRiskPanel(PDPageContentStream stream, ReportPage page) throws Exception {
        float y = 168F;
        text(stream, "Operational notes", MARGIN, y + 44F, PDType1Font.HELVETICA_BOLD, 12F, INK);
        stream.setNonStrokingColor(new Color(250, 250, 250));
        stream.addRect(MARGIN, y - 6F, CONTENT_WIDTH, 40F);
        stream.fill();
        text(stream, "Risk: " + page.risk, MARGIN + 10F, y + 19F, PDType1Font.HELVETICA_BOLD, 9F,
                page.risk.startsWith("Low") ? GREEN : page.risk.startsWith("Medium") ? AMBER : RED);
        text(stream, "Next action: " + page.nextAction, MARGIN + 10F, y + 5F, PDType1Font.HELVETICA, 9F, INK);
    }

    private static void drawFooter(PDPageContentStream stream, ReportPage page) throws Exception {
        stream.setStrokingColor(new Color(226, 232, 240));
        stream.moveTo(MARGIN, 64F);
        stream.lineTo(PAGE_WIDTH - MARGIN, 64F);
        stream.stroke();
        text(stream, "Generated demo data. Page " + page.pageNumber + " of " + PAGE_COUNT,
                MARGIN, 44F, PDType1Font.HELVETICA, 8.5F, MUTED);
        text(stream, "Confidential sample report", PAGE_WIDTH - MARGIN - 120F, 44F,
                PDType1Font.HELVETICA, 8.5F, MUTED);
    }

    private static BufferedImage reportImage(ReportPage page) {
        BufferedImage image = new BufferedImage(760, 512, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        Color base = page.pageNumber % 3 == 0 ? new Color(20, 116, 107)
                : page.pageNumber % 3 == 1 ? new Color(32, 82, 149) : new Color(142, 75, 16);
        g.setColor(new Color(245, 247, 250));
        g.fillRect(0, 0, image.getWidth(), image.getHeight());
        g.setColor(base);
        g.fillRect(0, 0, image.getWidth(), 112);
        g.setColor(Color.WHITE);
        g.setFont(new java.awt.Font("SansSerif", java.awt.Font.BOLD, 34));
        g.drawString(page.department, 40, 68);
        g.setFont(new java.awt.Font("SansSerif", java.awt.Font.PLAIN, 20));
        g.drawString(page.region + " dashboard", 40, 96);

        int[] bars = page.chartValues;
        for (int i = 0; i < bars.length; i++) {
            int barHeight = 60 + bars[i] * 3;
            int x = 60 + i * 85;
            g.setColor(new Color(base.getRed(), base.getGreen(), base.getBlue(), 180));
            g.fillRoundRect(x, 430 - barHeight, 48, barHeight, 10, 10);
            g.setColor(new Color(70, 70, 70));
            g.setFont(new java.awt.Font("SansSerif", java.awt.Font.PLAIN, 14));
            g.drawString("M" + (i + 1), x + 11, 456);
        }

        g.setColor(new Color(255, 255, 255, 230));
        g.fillRoundRect(430, 152, 270, 160, 18, 18);
        g.setColor(base);
        g.setFont(new java.awt.Font("SansSerif", java.awt.Font.BOLD, 44));
        g.drawString(page.health + "%", 468, 232);
        g.setFont(new java.awt.Font("SansSerif", java.awt.Font.PLAIN, 18));
        g.drawString("portfolio health", 468, 262);
        g.dispose();
        return image;
    }

    private static void text(PDPageContentStream stream,
                             String value,
                             float x,
                             float y,
                             PDType1Font font,
                             float size,
                             Color color) throws Exception {
        stream.beginText();
        stream.setFont(font, size);
        stream.setNonStrokingColor(color);
        stream.newLineAtOffset(x, y);
        stream.showText(safe(value));
        stream.endText();
    }

    private static List<String> wrap(String text, PDType1Font font, float fontSize, float maxWidth) throws Exception {
        List<String> lines = new ArrayList<>();
        String[] words = safe(text).split(" ");
        StringBuilder line = new StringBuilder();
        for (String word : words) {
            String candidate = line.length() == 0 ? word : line + " " + word;
            if (textWidth(candidate, font, fontSize) <= maxWidth) {
                line.setLength(0);
                line.append(candidate);
            } else {
                if (line.length() > 0) {
                    lines.add(line.toString());
                }
                line.setLength(0);
                line.append(word);
            }
        }
        if (line.length() > 0) {
            lines.add(line.toString());
        }
        return lines;
    }

    private static float textWidth(String text, PDType1Font font, float fontSize) throws Exception {
        return font.getStringWidth(safe(text)) / 1000F * fontSize;
    }

    private static String safe(String value) {
        if (value == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder(value.length());
        for (int i = 0; i < value.length(); i++) {
            char ch = value.charAt(i);
            builder.append(ch >= 32 && ch <= 126 ? ch : ' ');
        }
        return builder.toString();
    }

    private static Path defaultOutputPath() {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", OUTPUT_FILE_NAME);
        }
        return Paths.get("playground", "target", OUTPUT_FILE_NAME);
    }

    private static final class ReportPage {
        private final int pageNumber;
        private final String title;
        private final String quarter;
        private final String owner;
        private final String status;
        private final String department;
        private final String region;
        private final String imageCaption;
        private final List<String> paragraphs;
        private final String callout;
        private final double revenueMillions;
        private final double revenueDelta;
        private final int activeUsers;
        private final double userGrowth;
        private final int cycleDays;
        private final double cycleDelta;
        private final int health;
        private final int phaseIndex;
        private final String risk;
        private final String nextAction;
        private final int[] chartValues;

        private ReportPage(int pageNumber,
                           String title,
                           String quarter,
                           String owner,
                           String status,
                           String department,
                           String region,
                           String imageCaption,
                           List<String> paragraphs,
                           String callout,
                           double revenueMillions,
                           double revenueDelta,
                           int activeUsers,
                           double userGrowth,
                           int cycleDays,
                           double cycleDelta,
                           int health,
                           int phaseIndex,
                           String risk,
                           String nextAction,
                           int[] chartValues) {
            this.pageNumber = pageNumber;
            this.title = title;
            this.quarter = quarter;
            this.owner = owner;
            this.status = status;
            this.department = department;
            this.region = region;
            this.imageCaption = imageCaption;
            this.paragraphs = paragraphs;
            this.callout = callout;
            this.revenueMillions = revenueMillions;
            this.revenueDelta = revenueDelta;
            this.activeUsers = activeUsers;
            this.userGrowth = userGrowth;
            this.cycleDays = cycleDays;
            this.cycleDelta = cycleDelta;
            this.health = health;
            this.phaseIndex = phaseIndex;
            this.risk = risk;
            this.nextAction = nextAction;
            this.chartValues = chartValues;
        }

        private static ReportPage of(int pageNumber) {
            String[] departments = {"Customer Success", "Payments", "Logistics", "Data Platform", "Retail Growth"};
            String[] regions = {"North America", "Europe", "Asia Pacific", "Latin America", "Global"};
            String[] owners = {"Avery Chen", "Morgan Lee", "Sam Rivera", "Jordan Patel", "Taylor Kim"};
            String[] statuses = {"On track", "Ahead", "Focused watch", "Stabilizing"};
            String department = departments[(pageNumber - 1) % departments.length];
            String region = regions[(pageNumber + 1) % regions.length];
            int health = 70 + (pageNumber * 7) % 27;
            double revenue = 8.2 + pageNumber * 0.42 + (pageNumber % 5) * 1.15;
            double revenueDelta = -2.5 + (pageNumber % 12) * 0.72;
            int activeUsers = 42000 + pageNumber * 1370 + (pageNumber % 9) * 2300;
            double userGrowth = 2.1 + (pageNumber % 10) * 0.85;
            int cycleDays = 18 - (pageNumber % 6);
            double cycleDelta = -3.0 + (pageNumber % 7) * 0.6;
            String status = statuses[pageNumber % statuses.length];
            String risk = pageNumber % 11 == 0 ? "High: vendor dependency may delay rollout"
                    : pageNumber % 4 == 0 ? "Medium: adoption variance across regional teams"
                    : "Low: delivery indicators remain within tolerance";
            String callout = pageNumber % 4 == 0
                    ? "Watch item: field enablement should be reinforced before launch."
                    : "Momentum note: adoption improved and delivery remains ahead of the baseline.";
            List<String> paragraphs = Arrays.asList(
                    "This page reviews " + department + " performance for " + region
                            + ". The operating plan combines pipeline quality, delivery execution, customer outcomes, and support readiness into a single view for leadership review.",
                    "Revenue conversion improved in priority accounts while renewal conversations remained disciplined. The team used weekly account health reviews, faster escalation paths, and tighter implementation planning to keep delivery ahead of the baseline.",
                    "The main risk is monitored through customer readiness, staffing coverage, and integration quality. When a metric moves outside the expected range, the next action is assigned to a named owner and reviewed in the following operating meeting."
            );
            int[] chartValues = new int[7];
            for (int i = 0; i < chartValues.length; i++) {
                chartValues[i] = 18 + ((pageNumber * 5 + i * 9) % 62);
            }
            return new ReportPage(
                    pageNumber,
                    "Portfolio Review " + pageNumber + ": " + department,
                    "FY2026 Q" + (((pageNumber - 1) / 25) + 1),
                    owners[pageNumber % owners.length],
                    status,
                    department,
                    region,
                    "monthly trend and health summary",
                    paragraphs,
                    callout,
                    revenue,
                    revenueDelta,
                    activeUsers,
                    userGrowth,
                    cycleDays,
                    cycleDelta,
                    health,
                    pageNumber % 4,
                    risk,
                    "Confirm owner plan for milestone " + (pageNumber % 8 + 1) + " and publish the Friday update.",
                    chartValues
            );
        }
    }
}

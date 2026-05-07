package com.chenbitao.word.pdf;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;

/**
 * PDFBox 测试日志配置加载器。
 */
final class PdfBoxTestLoggingConfigurer {

    private static final String RESOURCE_NAME = "application-logging-test.yml";
    private static final String LOG_LEVEL_OVERRIDE = "pdfbox.log.level";

    private PdfBoxTestLoggingConfigurer() {
    }

    static void configure() {
        String pdfboxLevel = readLevel("org.apache.pdfbox");
        String fontboxLevel = readLevel("org.apache.fontbox");
        setLevel("org.apache.pdfbox", override(pdfboxLevel));
        setLevel("org.apache.fontbox", override(fontboxLevel));
    }

    private static String override(String configuredLevel) {
        String override = System.getProperty(LOG_LEVEL_OVERRIDE);
        if (override != null && !override.trim().isEmpty()) {
            return override.trim();
        }
        return configuredLevel == null || configuredLevel.trim().isEmpty() ? "WARN" : configuredLevel.trim();
    }

    private static String readLevel(String loggerName) {
        InputStream input = PdfBoxTestLoggingConfigurer.class.getClassLoader().getResourceAsStream(RESOURCE_NAME);
        if (input == null) {
            return "WARN";
        }

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(input, StandardCharsets.UTF_8))) {
            String line;
            while ((line = reader.readLine()) != null) {
                String trimmed = line.trim();
                if (trimmed.startsWith(loggerName + ":")) {
                    return trimmed.substring(trimmed.indexOf(':') + 1).trim();
                }
            }
        } catch (Exception ignored) {
            return "WARN";
        }
        return "WARN";
    }

    private static void setLevel(String loggerName, String level) {
        Logger logger = (Logger) LoggerFactory.getLogger(loggerName);
        logger.setLevel(Level.toLevel(level, Level.WARN));
    }
}

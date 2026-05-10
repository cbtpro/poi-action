package com.chenbitao.word.playground.demo.pdf;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

/**
 * PDFBox 日志配置加载器。
 *
 * <p>playground demo 是普通 main 方法，不会自动启动 Spring 环境，
 * 因此这里主动读取 application YAML 中的 PDFBox 相关日志级别配置。</p>
 */
public final class PdfBoxLoggingConfigurer {

    private static final String LOG_LEVEL_OVERRIDE = "pdfbox.log.level";
    private static final String ACTIVE_PROFILE = "spring.profiles.active";
    private static final String ACTIVE_PROFILE_ENV = "SPRING_PROFILES_ACTIVE";

    private PdfBoxLoggingConfigurer() {
    }

    public static void configure() {
        Properties properties = load("application-logging.yml");
        String profile = activeProfile();
        if (!profile.isEmpty()) {
            properties.putAll(load("application-logging-" + profile + ".yml"));
        }

        apply(properties, "org.apache.pdfbox");
        apply(properties, "org.apache.fontbox");
        applyOverride();
    }

    private static void apply(Properties properties, String loggerName) {
        String level = properties.getProperty("logging.level." + loggerName);
        if (level != null && !level.trim().isEmpty()) {
            setLevel(loggerName, level);
        }
    }

    private static void applyOverride() {
        String level = System.getProperty(LOG_LEVEL_OVERRIDE);
        if (level != null && !level.trim().isEmpty()) {
            setLevel("org.apache.pdfbox", level);
            setLevel("org.apache.fontbox", level);
        }
    }

    private static void setLevel(String loggerName, String level) {
        Logger logger = (Logger) LoggerFactory.getLogger(loggerName);
        logger.setLevel(Level.toLevel(level.trim(), Level.WARN));
    }

    private static Properties load(String resourceName) {
        Properties properties = new Properties();
        InputStream input = PdfBoxLoggingConfigurer.class.getClassLoader().getResourceAsStream(resourceName);
        if (input == null) {
            return properties;
        }

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(input, StandardCharsets.UTF_8))) {
            String line;
            while ((line = reader.readLine()) != null) {
                readLevel(properties, line, "org.apache.pdfbox");
                readLevel(properties, line, "org.apache.fontbox");
            }
        } catch (Exception ignored) {
            return new Properties();
        }
        return properties;
    }

    private static void readLevel(Properties properties, String line, String loggerName) {
        String trimmed = line.trim();
        if (trimmed.startsWith(loggerName + ":")) {
            String level = trimmed.substring(trimmed.indexOf(':') + 1).trim();
            properties.setProperty("logging.level." + loggerName, level);
        }
    }

    private static String activeProfile() {
        String profile = System.getProperty(ACTIVE_PROFILE);
        if (profile == null || profile.trim().isEmpty()) {
            profile = System.getenv(ACTIVE_PROFILE_ENV);
        }
        if (profile == null || profile.trim().isEmpty()) {
            return "";
        }
        return profile.split(",")[0].trim();
    }
}

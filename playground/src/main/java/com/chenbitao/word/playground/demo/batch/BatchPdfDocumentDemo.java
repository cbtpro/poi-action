package com.chenbitao.word.playground.demo.batch;

import com.chenbitao.word.playground.demo.pdf.PdfBoxLoggingConfigurer;
import com.chenbitao.word.playground.demo.pdf.PdfProjectReportDemo;
import lombok.extern.slf4j.Slf4j;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

/**
 * 批量生成 PDF 演示。
 */
@Slf4j
public class BatchPdfDocumentDemo {

    private static final int DEFAULT_COUNT = 32000;
    private static final int DEFAULT_THREADS = Math.max(1, Runtime.getRuntime().availableProcessors() / 2);
    private static final int PROGRESS_INTERVAL_SECONDS = 5;
    private static final String OUTPUT_DIR_NAME = "pdf-batch-out";

    /**
     * 批量生成 PDF 入口。
     *
     * <p>第一个参数是生成数量，第二个参数是工作线程数。
     * 每个输出文件都会复用 {@link PdfProjectReportDemo} 生成完整的 100 页 PDF。</p>
     *
     * @param args 命令行参数
     * @throws Exception 生成失败时抛出
     */
    public static void main(String[] args) throws Exception {
        int count = parseCount(args);
        int threads = parseThreads(args);
        Path outputDirectory = defaultOutputDirectory();
        long start = System.currentTimeMillis();

        generate(count, threads, outputDirectory);

        long cost = System.currentTimeMillis() - start;
        double average = count == 0 ? 0D : cost * 1.0D / count;
        log.info("批量 100 页 PDF 生成完成：{} 个文件，线程数：{}，输出目录：{}，耗时：{} ms，平均：{} ms/文件",
                count,
                threads,
                outputDirectory.toAbsolutePath(),
                cost,
                String.format("%.3f", average));
    }

    /**
     * 使用默认输出目录生成 100 页 PDF。
     *
     * @param count 文件数量
     * @param threads 工作线程数
     * @throws Exception 生成失败时抛出
     */
    public static void generate(int count, int threads) throws Exception {
        generate(count, threads, defaultOutputDirectory());
    }

    /**
     * 使用指定输出目录生成 100 页 PDF。
     *
     * @param count 文件数量
     * @param threads 工作线程数
     * @param outputDirectory 输出目录
     * @throws Exception 生成失败时抛出
     */
    public static void generate(int count, int threads, Path outputDirectory) throws Exception {
        PdfBoxLoggingConfigurer.configure();
        Files.createDirectories(outputDirectory);
        if (count <= 0) {
            log.info("PDF 生成数量为 0，无需执行");
            return;
        }

        BatchProgress progress = new BatchProgress(count);
        AtomicInteger nextIndex = new AtomicInteger(1);
        AtomicReference<Exception> error = new AtomicReference<>();
        int workerCount = Math.max(1, threads);

        ExecutorService executor = Executors.newFixedThreadPool(workerCount);
        ScheduledExecutorService progressLogger = Executors.newSingleThreadScheduledExecutor();
        progressLogger.scheduleAtFixedRate(() -> logProgress(progress),
                PROGRESS_INTERVAL_SECONDS,
                PROGRESS_INTERVAL_SECONDS,
                TimeUnit.SECONDS);

        for (int i = 0; i < workerCount; i++) {
            executor.execute(() -> writeDocuments(outputDirectory, count, nextIndex, progress, error));
        }

        executor.shutdown();
        while (!executor.awaitTermination(1, TimeUnit.MINUTES)) {
            logProgress(progress);
        }
        progressLogger.shutdownNow();
        logProgress(progress);

        if (error.get() != null) {
            throw error.get();
        }
    }

    private static void writeDocuments(Path outputDirectory,
                                       int count,
                                       AtomicInteger nextIndex,
                                       BatchProgress progress,
                                       AtomicReference<Exception> error) {
        while (error.get() == null) {
            int index = nextIndex.getAndIncrement();
            if (index > count) {
                return;
            }

            try {
                PdfProjectReportDemo.generate(outputDirectory.resolve(fileName(index)));
                progress.completed.incrementAndGet();
            } catch (Exception e) {
                progress.failed.incrementAndGet();
                error.compareAndSet(null, e);
                return;
            }
        }
    }

    private static void logProgress(BatchProgress progress) {
        int completed = progress.completed.get();
        int failed = progress.failed.get();
        long elapsedMillis = Math.max(1, System.currentTimeMillis() - progress.startMillis);
        double percent = completed * 100.0 / progress.total;
        double speed = completed * 1000.0 / elapsedMillis;
        long remaining = Math.max(0, progress.total - completed);
        long etaSeconds = speed <= 0 ? -1 : Math.round(remaining / speed);

        log.info("100 页 PDF 生成进度：{}/{}（{}%），失败：{}，速度：{} 个/秒，预计剩余：{}",
                completed,
                progress.total,
                String.format("%.2f", percent),
                failed,
                String.format("%.2f", speed),
                etaSeconds < 0 ? "未知" : etaSeconds + " 秒");
    }

    private static String fileName(int index) {
        return String.format("batch-pdf-%05d.pdf", index);
    }

    private static int parseCount(String[] args) {
        if (args == null || args.length == 0 || args[0] == null || args[0].trim().isEmpty()) {
            return DEFAULT_COUNT;
        }
        return Integer.parseInt(args[0]);
    }

    private static int parseThreads(String[] args) {
        if (args == null || args.length < 2 || args[1] == null || args[1].trim().isEmpty()) {
            return DEFAULT_THREADS;
        }
        return Math.max(1, Integer.parseInt(args[1]));
    }

    private static Path defaultOutputDirectory() {
        Path workingDirectory = Paths.get("").toAbsolutePath();
        if ("playground".equals(workingDirectory.getFileName().toString())) {
            return Paths.get("target", OUTPUT_DIR_NAME);
        }
        return Paths.get("playground", "target", OUTPUT_DIR_NAME);
    }

    private static final class BatchProgress {
        private final int total;
        private final long startMillis;
        private final AtomicInteger completed = new AtomicInteger();
        private final AtomicInteger failed = new AtomicInteger();

        private BatchProgress(int total) {
            this.total = total;
            this.startMillis = System.currentTimeMillis();
        }
    }
}

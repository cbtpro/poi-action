package com.chenbitao.word.demo;

import com.chenbitao.word.docx.TemplateWordGenerator;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

@Slf4j
public class TemplateBatchDemo {

    private static final int DEFAULT_COUNT = 1000;
    private static final int DEFAULT_THREADS = Math.max(1, Runtime.getRuntime().availableProcessors());
    private static final int PROGRESS_INTERVAL_SECONDS = 5;
    private static final Path OUTPUT_DIR = Paths.get("target", "out");

    /**
     * 批量生成入口。
     *
     * <p>第一个参数是生成数量，第二个参数是线程数；未传时分别使用默认值。</p>
     */
    public static void main(String[] args) throws Exception {
        int count = parseCount(args);
        int threads = parseThreads(args);
        long start = System.currentTimeMillis();

        generate(count, threads);

        long cost = System.currentTimeMillis() - start;
        log.info("批量生成完成：{} 个文档，线程数：{}，输出目录：{}，耗时：{} ms，平均：{} ms/个",
                count,
                threads,
                OUTPUT_DIR.toAbsolutePath(),
                cost,
                count == 0 ? 0 : cost / count);
    }

    /**
     * 使用默认线程数批量生成文档。
     *
     * @param count 需要生成的文档数量
     */
    public static void generate(int count) throws Exception {
        generate(count, DEFAULT_THREADS);
    }

    /**
     * 批量生成文档。
     *
     * <p>性能优化点：模板只读取一次，并且只渲染一次；后续多线程复用渲染后的
     * docx 字节写入不同文件。适合 demo 场景下验证大批量文件写出性能。</p>
     *
     * @param count 需要生成的文档数量
     * @param threads 并发写文件的线程数
     */
    public static void generate(int count, int threads) throws Exception {
        Files.createDirectories(OUTPUT_DIR);
        if (count <= 0) {
            log.info("生成数量为 0，无需处理");
            return;
        }

        byte[] templateBytes = readTemplateBytes();
        log.info("模板读取完成：{} bytes", templateBytes.length);

        byte[] documentBytes = renderDocumentBytes(templateBytes, TemplateDemoData.create());
        log.info("模板渲染完成：{} bytes，开始多线程写入文件", documentBytes.length);

        BatchProgress progress = new BatchProgress(count);
        AtomicInteger nextIndex = new AtomicInteger(1);
        AtomicReference<Exception> error = new AtomicReference<>();
        int workerCount = Math.max(1, threads);

        ExecutorService executor = Executors.newFixedThreadPool(workerCount);
        ScheduledExecutorService progressLogger = Executors.newSingleThreadScheduledExecutor();
        progressLogger.scheduleAtFixedRate(() -> logProgress(progress), PROGRESS_INTERVAL_SECONDS, PROGRESS_INTERVAL_SECONDS, TimeUnit.SECONDS);

        for (int i = 0; i < workerCount; i++) {
            executor.execute(() -> writeDocuments(documentBytes, count, nextIndex, progress, error));
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

    /**
     * 将模板字节和演示数据渲染成最终 Word 文档字节。
     *
     * @param templateBytes 模板文件字节
     * @param data 模板渲染数据
     * @return 已完成渲染的 docx 字节
     */
    private static byte[] renderDocumentBytes(byte[] templateBytes, Map<String, Object> data) throws IOException {
        TemplateWordGenerator generator = new TemplateWordGenerator(new ByteArrayInputStream(templateBytes));
        generator.render(data);
        try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            generator.save(output);
            return output.toByteArray();
        }
    }

    /**
     * worker 循环领取文件序号并写出文档。
     *
     * <p>使用 {@link AtomicInteger} 分配序号，避免一次性创建大量任务占用内存；
     * 任一 worker 失败后会记录异常，其它 worker 会尽快停止。</p>
     */
    private static void writeDocuments(byte[] documentBytes,
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
                writeOne(documentBytes, index);
                progress.completed.incrementAndGet();
            } catch (Exception e) {
                progress.failed.incrementAndGet();
                error.compareAndSet(null, e);
                return;
            }
        }
    }

    /**
     * 写出单个 Word 文件。
     *
     * @param documentBytes 已渲染好的 docx 字节
     * @param index 文件序号
     */
    private static void writeOne(byte[] documentBytes, int index) throws IOException {
        Files.write(OUTPUT_DIR.resolve(fileName(index)), documentBytes);
    }

    /**
     * 输出当前生成进度，包括完成数量、失败数量、速度和预计剩余时间。
     */
    private static void logProgress(BatchProgress progress) {
        int completed = progress.completed.get();
        int failed = progress.failed.get();
        long elapsedMillis = Math.max(1, System.currentTimeMillis() - progress.startMillis);
        double percent = completed * 100.0 / progress.total;
        double speed = completed * 1000.0 / elapsedMillis;
        long remaining = Math.max(0, progress.total - completed);
        long etaSeconds = speed <= 0 ? -1 : Math.round(remaining / speed);

        log.info("生成进度：{}/{}，{}%，失败：{}，速度：{} 个/秒，预计剩余：{}",
                completed,
                progress.total,
                String.format("%.2f", percent),
                failed,
                String.format("%.2f", speed),
                etaSeconds < 0 ? "未知" : etaSeconds + " 秒");
    }

    /**
     * 读取模板文件为字节数组。
     *
     * <p>模板只读取一次，后续生成过程都复用这份字节，避免重复读取 classpath 资源。</p>
     */
    private static byte[] readTemplateBytes() throws IOException {
        try (InputStream template = TemplateDemo.loadTemplate();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[8192];
            int length;
            while ((length = template.read(buffer)) != -1) {
                output.write(buffer, 0, length);
            }
            return output.toByteArray();
        }
    }

    /**
     * 根据序号生成输出文件名。
     */
    private static String fileName(int index) {
        return String.format("template-demo-%06d.docx", index);
    }

    /**
     * 解析生成数量参数。
     */
    private static int parseCount(String[] args) {
        if (args == null || args.length == 0 || args[0] == null || args[0].trim().isEmpty()) {
            return DEFAULT_COUNT;
        }
        return Integer.parseInt(args[0]);
    }

    /**
     * 解析线程数参数。
     */
    private static int parseThreads(String[] args) {
        if (args == null || args.length < 2 || args[1] == null || args[1].trim().isEmpty()) {
            return DEFAULT_THREADS;
        }
        return Math.max(1, Integer.parseInt(args[1]));
    }

    /**
     * 批量生成进度状态。
     */
    private static class BatchProgress {
        private final int total;
        private final long startMillis;
        private final AtomicInteger completed;
        private final AtomicInteger failed;

        private BatchProgress(int total) {
            this.total = total;
            this.startMillis = System.currentTimeMillis();
            this.completed = new AtomicInteger();
            this.failed = new AtomicInteger();
        }
    }
}

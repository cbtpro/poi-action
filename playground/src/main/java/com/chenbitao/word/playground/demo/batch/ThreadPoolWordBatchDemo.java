package com.chenbitao.word.playground.demo.batch;

import com.chenbitao.word.builder.WordBuilder;
import com.chenbitao.word.docx.DocxWordGenerator;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

@Slf4j
public class ThreadPoolWordBatchDemo {

    private static final int DEFAULT_COUNT = 320000;
    private static final int DEFAULT_THREADS = Math.max(1, Runtime.getRuntime().availableProcessors() / 2);
    private static final int PROGRESS_INTERVAL_SECONDS = 3;
    private static final Path OUTPUT_DIR = Paths.get("target", "thread-pool-out");

    /**
     * 线程池批量生成演示入口。
     *
     * <p>该演示每个文件都重新构建一个简单文档，用来展示生成器实例隔离的并发写入方式；
     * 模板复用型批量场景见 {@link TemplateBatchDocumentDemo}。</p>
     */
    public static void main(String[] args) throws Exception {
        int count = parseCount(args);
        int threads = parseThreads(args);
        long start = System.currentTimeMillis();

        generate(count, threads);

        long cost = System.currentTimeMillis() - start;
        double average = count == 0 ? 0D : cost * 1.0D / count;
        log.info("线程池批量生成完成：{} 个文档，线程数：{}，输出目录：{}，总耗时：{} ms，平均耗时：{} ms/个",
                count,
                threads,
                OUTPUT_DIR.toAbsolutePath(),
                cost,
                String.format("%.3f", average));
    }

    private static void generate(int count, int threads) throws Exception {
        Files.createDirectories(OUTPUT_DIR);
        if (count <= 0) {
            log.info("生成数量为 0，无需处理");
            return;
        }

        final AtomicInteger nextIndex = new AtomicInteger(1);
        final AtomicInteger completed = new AtomicInteger();
        final AtomicInteger failed = new AtomicInteger();
        final AtomicReference<Exception> error = new AtomicReference<>();

        ExecutorService executor = Executors.newFixedThreadPool(threads);
        ScheduledExecutorService progressLogger = Executors.newSingleThreadScheduledExecutor();
        progressLogger.scheduleAtFixedRate(() -> logProgress(count, completed.get(), failed.get()),
                PROGRESS_INTERVAL_SECONDS,
                PROGRESS_INTERVAL_SECONDS,
                TimeUnit.SECONDS);

        for (int i = 0; i < threads; i++) {
            executor.execute(() -> {
                while (error.get() == null) {
                    int index = nextIndex.getAndIncrement();
                    if (index > count) {
                        return;
                    }
                    try {
                        writeDocument(index);
                        completed.incrementAndGet();
                    } catch (Exception e) {
                        failed.incrementAndGet();
                        error.compareAndSet(null, e);
                        return;
                    }
                }
            });
        }

        executor.shutdown();
        while (!executor.awaitTermination(1, TimeUnit.MINUTES)) {
            logProgress(count, completed.get(), failed.get());
        }
        progressLogger.shutdownNow();
        logProgress(count, completed.get(), failed.get());

        if (error.get() != null) {
            throw error.get();
        }
    }

    private static void writeDocument(int index) throws IOException {
        String fileName = String.format("thread-pool-document-%04d.docx", index);
        Path output = OUTPUT_DIR.resolve(fileName);

        WordBuilder builder = new WordBuilder(new DocxWordGenerator());
        // 每个任务独立持有 DocxWordGenerator，避免多线程共享可变文档对象。
        builder.title("批量文档示例 " + index)
                .paragraph("这是第 " + index + " 个批量生成的 Word 文档。")
                .paragraph("线程池并发写入可以显著提高大批量文档生成的吞吐量。")
                .paragraphList(sampleBulletPoints())
                .table(Arrays.asList(
                        Arrays.asList("属性", "值"),
                        Arrays.asList("文档编号", String.valueOf(index)),
                        Arrays.asList("线程数", String.valueOf(Runtime.getRuntime().availableProcessors())),
                        Arrays.asList("生成时间", String.valueOf(System.currentTimeMillis()))
                ))
                .build(output.toString());
    }

    private static List<String> sampleBulletPoints() {
        return Arrays.asList(
                "支持并发生成多个 Word 文档",
                "每个文档使用独立 DocxWordGenerator 实例",
                "避免多个线程复用同一个生成器实例导致数据冲突"
        );
    }

    private static void logProgress(int total, int completed, int failed) {
        double percentage = total == 0 ? 100D : completed * 100.0D / total;
        log.info("线程池生成进度：{}/{}，{}% 完成，失败：{}",
                completed,
                total,
                String.format("%.2f", percentage),
                failed);
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
}

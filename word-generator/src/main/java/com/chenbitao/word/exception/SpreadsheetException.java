package com.chenbitao.word.exception;

/**
 * 电子表格处理异常。
 *
 * <p>用于封装 Excel 生成、写出和样式处理过程中的运行时异常。</p>
 */
public class SpreadsheetException extends RuntimeException {

    /**
     * 构造电子表格处理异常。
     *
     * @param message 异常信息
     * @param cause 异常原因
     */
    public SpreadsheetException(String message, Throwable cause) {
        super(message, cause);
    }
}

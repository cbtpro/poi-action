package com.chenbitao.word.exception;

/**
 * Word文档处理异常
 * Word文档生成和处理过程中发生的运行时异常
 */
public class WordException extends RuntimeException {

    /**
     * 构造Word文档处理异常
     *
     * @param message 异常信息
     * @param cause 异常原因
     */
    public WordException(String message, Throwable cause) {
        super(message, cause);
    }
}
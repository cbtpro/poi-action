package com.chenbitao.word.exception;

/**
 * 演示文稿处理异常。
 *
 * <p>用于封装 PowerPoint 生成、图片插入和写出过程中的运行时异常。</p>
 */
public class PresentationException extends RuntimeException {

    /**
     * 构造演示文稿处理异常。
     *
     * @param message 异常信息
     * @param cause 异常原因
     */
    public PresentationException(String message, Throwable cause) {
        super(message, cause);
    }
}

package com.chenbitao.word.exception;

/**
 * Outlook 邮件读取异常。
 *
 * <p>用于封装 {@code .msg} 文件读取、解析过程中的底层异常。</p>
 */
public class OutlookMessageException extends RuntimeException {

    public OutlookMessageException(String message, Throwable cause) {
        super(message, cause);
    }
}

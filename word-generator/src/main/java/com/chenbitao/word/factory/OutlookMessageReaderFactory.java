package com.chenbitao.word.factory;

import com.chenbitao.word.outlook.OutlookMessageReader;

/**
 * Outlook 邮件读取器工厂。
 *
 * <p>当前支持 Outlook {@code msg} 类型。</p>
 */
public class OutlookMessageReaderFactory {

    private OutlookMessageReaderFactory() {
    }

    /**
     * 获取指定类型的 Outlook 邮件读取器。
     *
     * @param type 文件类型，目前支持 {@code msg}
     * @return 新的 Outlook 邮件读取器实例
     */
    public static OutlookMessageReader get(String type) {
        if ("msg".equalsIgnoreCase(type)) {
            return new OutlookMessageReader();
        }
        throw new IllegalArgumentException("不支持 Outlook 邮件类型: " + type);
    }
}

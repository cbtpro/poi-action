package com.chenbitao.word.outlook;

/**
 * Outlook 邮件附件摘要信息。
 */
public class OutlookAttachmentInfo {

    private final String fileName;
    private final String mimeType;
    private final String contentId;
    private final long size;
    private final boolean embeddedMessage;

    public OutlookAttachmentInfo(String fileName, String mimeType, String contentId, long size, boolean embeddedMessage) {
        this.fileName = fileName;
        this.mimeType = mimeType;
        this.contentId = contentId;
        this.size = size;
        this.embeddedMessage = embeddedMessage;
    }

    public String getFileName() {
        return fileName;
    }

    public String getMimeType() {
        return mimeType;
    }

    public String getContentId() {
        return contentId;
    }

    public long getSize() {
        return size;
    }

    public boolean isEmbeddedMessage() {
        return embeddedMessage;
    }
}

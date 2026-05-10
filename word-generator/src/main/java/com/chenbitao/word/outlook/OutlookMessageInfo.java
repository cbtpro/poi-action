package com.chenbitao.word.outlook;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Outlook 邮件摘要信息。
 */
public class OutlookMessageInfo {

    private final String subject;
    private final String from;
    private final String to;
    private final String cc;
    private final String bcc;
    private final String textBody;
    private final String htmlBody;
    private final List<String> recipientNames;
    private final List<String> recipientEmailAddresses;
    private final List<OutlookAttachmentInfo> attachments;

    public OutlookMessageInfo(String subject,
                              String from,
                              String to,
                              String cc,
                              String bcc,
                              String textBody,
                              String htmlBody,
                              List<String> recipientNames,
                              List<String> recipientEmailAddresses,
                              List<OutlookAttachmentInfo> attachments) {
        this.subject = subject;
        this.from = from;
        this.to = to;
        this.cc = cc;
        this.bcc = bcc;
        this.textBody = textBody;
        this.htmlBody = htmlBody;
        this.recipientNames = unmodifiableCopy(recipientNames);
        this.recipientEmailAddresses = unmodifiableCopy(recipientEmailAddresses);
        this.attachments = unmodifiableCopy(attachments);
    }

    public String getSubject() {
        return subject;
    }

    public String getFrom() {
        return from;
    }

    public String getTo() {
        return to;
    }

    public String getCc() {
        return cc;
    }

    public String getBcc() {
        return bcc;
    }

    public String getTextBody() {
        return textBody;
    }

    public String getHtmlBody() {
        return htmlBody;
    }

    public List<String> getRecipientNames() {
        return recipientNames;
    }

    public List<String> getRecipientEmailAddresses() {
        return recipientEmailAddresses;
    }

    public List<OutlookAttachmentInfo> getAttachments() {
        return attachments;
    }

    public int getAttachmentCount() {
        return attachments.size();
    }

    private static <T> List<T> unmodifiableCopy(List<T> values) {
        if (values == null || values.isEmpty()) {
            return Collections.emptyList();
        }
        return Collections.unmodifiableList(new ArrayList<>(values));
    }
}

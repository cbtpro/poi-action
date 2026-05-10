package com.chenbitao.word.outlook;

import com.chenbitao.word.exception.OutlookMessageException;
import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.datatypes.ByteChunk;
import org.apache.poi.hsmf.datatypes.StringChunk;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * Outlook {@code .msg} 邮件读取器。
 *
 * <p>基于 Apache POI HSMF 提取邮件主题、发件人、收件人、正文和附件摘要信息。</p>
 */
public class OutlookMessageReader {

    /**
     * 从文件路径读取 Outlook 邮件。
     *
     * @param path 邮件文件路径
     * @return 邮件摘要信息
     */
    public OutlookMessageInfo read(String path) {
        return read(new File(path));
    }

    /**
     * 从文件路径读取 Outlook 邮件。
     *
     * @param path 邮件文件路径
     * @return 邮件摘要信息
     */
    public OutlookMessageInfo read(Path path) {
        return read(path.toFile());
    }

    /**
     * 从文件读取 Outlook 邮件。
     *
     * @param file 邮件文件
     * @return 邮件摘要信息
     */
    public OutlookMessageInfo read(File file) {
        try (MAPIMessage message = new MAPIMessage(file)) {
            return extract(message);
        } catch (Exception e) {
            throw new OutlookMessageException("读取 Outlook MSG 文件失败", e);
        }
    }

    /**
     * 从输入流读取 Outlook 邮件。
     *
     * @param inputStream 邮件输入流
     * @return 邮件摘要信息
     */
    public OutlookMessageInfo read(InputStream inputStream) {
        try (MAPIMessage message = new MAPIMessage(inputStream)) {
            return extract(message);
        } catch (Exception e) {
            throw new OutlookMessageException("读取 Outlook MSG 输入流失败", e);
        }
    }

    private OutlookMessageInfo extract(MAPIMessage message) {
        message.setReturnNullOnMissingChunk(true);
        guessEncoding(message);
        return new OutlookMessageInfo(
                value(message::getSubject),
                value(message::getDisplayFrom),
                value(message::getDisplayTo),
                value(message::getDisplayCC),
                value(message::getDisplayBCC),
                value(message::getTextBody),
                value(message::getHtmlBody),
                values(message::getRecipientNamesList),
                values(message::getRecipientEmailAddressList),
                attachments(message.getAttachmentFiles())
        );
    }

    private void guessEncoding(MAPIMessage message) {
        try {
            message.guess7BitEncoding();
        } catch (Exception ignored) {
            // 部分 MSG 缺少编码辅助属性时，HSMF 会回退到 chunk 自身编码。
        }
    }

    private List<OutlookAttachmentInfo> attachments(AttachmentChunks[] chunks) {
        if (chunks == null || chunks.length == 0) {
            return Collections.emptyList();
        }

        List<OutlookAttachmentInfo> attachments = new ArrayList<>();
        for (AttachmentChunks chunk : chunks) {
            attachments.add(new OutlookAttachmentInfo(
                    first(value(chunk.getAttachLongFileName()), value(chunk.getAttachFileName())),
                    value(chunk.getAttachMimeTag()),
                    value(chunk.getAttachContentId()),
                    attachmentSize(chunk),
                    chunk.isEmbeddedMessage()
            ));
        }
        return attachments;
    }

    private long attachmentSize(AttachmentChunks chunk) {
        byte[] data = chunk.getEmbeddedAttachmentObject();
        if (data != null) {
            return data.length;
        }

        ByteChunk attachData = chunk.getAttachData();
        return attachData == null || attachData.getValue() == null ? 0 : attachData.getValue().length;
    }

    private String first(String primary, String fallback) {
        return primary == null || primary.isEmpty() ? fallback : primary;
    }

    private String value(StringChunk chunk) {
        return chunk == null || chunk.getValue() == null ? "" : chunk.getValue();
    }

    private String value(MessageStringSupplier supplier) {
        try {
            String value = supplier.get();
            return value == null ? "" : value;
        } catch (Exception e) {
            return "";
        }
    }

    private List<String> values(MessageStringArraySupplier supplier) {
        try {
            String[] values = supplier.get();
            if (values == null || values.length == 0) {
                return Collections.emptyList();
            }
            return Arrays.asList(values);
        } catch (Exception e) {
            return Collections.emptyList();
        }
    }

    private interface MessageStringSupplier {
        String get() throws Exception;
    }

    private interface MessageStringArraySupplier {
        String[] get() throws Exception;
    }
}

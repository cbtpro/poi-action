package com.chenbitao.word.publisher;

/**
 * Publisher document summary information.
 */
public class PublisherDocumentInfo {

    private final String title;
    private final String subject;
    private final String author;
    private final String keywords;
    private final String comments;
    private final String text;

    public PublisherDocumentInfo(String title,
                                 String subject,
                                 String author,
                                 String keywords,
                                 String comments,
                                 String text) {
        this.title = title;
        this.subject = subject;
        this.author = author;
        this.keywords = keywords;
        this.comments = comments;
        this.text = text;
    }

    public String getTitle() {
        return title;
    }

    public String getSubject() {
        return subject;
    }

    public String getAuthor() {
        return author;
    }

    public String getKeywords() {
        return keywords;
    }

    public String getComments() {
        return comments;
    }

    public String getText() {
        return text;
    }

    public boolean hasText() {
        return text != null && !text.isEmpty();
    }
}

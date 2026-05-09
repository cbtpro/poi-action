package com.chenbitao.word.publisher;

import com.chenbitao.word.exception.PublisherDocumentException;
import org.apache.poi.hpbf.HPBFDocument;
import org.apache.poi.hpbf.extractor.PublisherTextExtractor;
import org.apache.poi.hpsf.SummaryInformation;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;

/**
 * Publisher document reader based on Apache POI HPBF.
 */
public class PublisherDocumentReader {

    /**
     * Read a Publisher document from a file path.
     *
     * @param path Publisher file path
     * @return Publisher document summary
     */
    public PublisherDocumentInfo read(String path) {
        return read(new File(path));
    }

    /**
     * Read a Publisher document from a file path.
     *
     * @param path Publisher file path
     * @return Publisher document summary
     */
    public PublisherDocumentInfo read(Path path) {
        return read(path.toFile());
    }

    /**
     * Read a Publisher document from a file.
     *
     * @param file Publisher file
     * @return Publisher document summary
     */
    public PublisherDocumentInfo read(File file) {
        try (FileInputStream inputStream = new FileInputStream(file)) {
            return read(inputStream);
        } catch (Exception e) {
            throw new PublisherDocumentException("Read Publisher document failed", e);
        }
    }

    /**
     * Read a Publisher document from an input stream.
     *
     * @param inputStream Publisher input stream
     * @return Publisher document summary
     */
    public PublisherDocumentInfo read(InputStream inputStream) {
        try (HPBFDocument document = new HPBFDocument(inputStream);
             PublisherTextExtractor extractor = new PublisherTextExtractor(document)) {
            SummaryInformation summary = document.getSummaryInformation();
            return new PublisherDocumentInfo(
                    summary == null ? "" : value(summary.getTitle()),
                    summary == null ? "" : value(summary.getSubject()),
                    summary == null ? "" : value(summary.getAuthor()),
                    summary == null ? "" : value(summary.getKeywords()),
                    summary == null ? "" : value(summary.getComments()),
                    value(extractor.getText())
            );
        } catch (Exception e) {
            throw new PublisherDocumentException("Read Publisher input stream failed", e);
        }
    }

    private String value(String value) {
        return value == null ? "" : value;
    }
}

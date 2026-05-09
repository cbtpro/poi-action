package com.chenbitao.word.visio;

import com.chenbitao.word.exception.VisioDrawingException;
import org.apache.poi.hdgf.extractor.VisioTextExtractor;
import org.apache.poi.xdgf.geom.Dimension2dDouble;
import org.apache.poi.xdgf.usermodel.XDGFPage;
import org.apache.poi.xdgf.usermodel.XDGFShape;
import org.apache.poi.xdgf.usermodel.XmlVisioDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * Visio drawing reader based on Apache POI HDGF and XDGF.
 */
public class VisioDrawingReader {

    /**
     * Read a Visio drawing from a file path.
     *
     * @param path Visio file path
     * @return drawing summary
     */
    public VisioDrawingInfo read(String path) {
        return read(new File(path));
    }

    /**
     * Read a Visio drawing from a file path.
     *
     * @param path Visio file path
     * @return drawing summary
     */
    public VisioDrawingInfo read(Path path) {
        return read(path.toFile());
    }

    /**
     * Read a Visio drawing from a file.
     *
     * @param file Visio file
     * @return drawing summary
     */
    public VisioDrawingInfo read(File file) {
        String type = extension(file.getName());
        try (FileInputStream inputStream = new FileInputStream(file)) {
            return read(inputStream, type);
        } catch (Exception e) {
            throw new VisioDrawingException("Read Visio drawing failed", e);
        }
    }

    /**
     * Read a Visio drawing from an input stream.
     *
     * @param inputStream Visio input stream
     * @param type file type, currently supports {@code vsd} and {@code vsdx}
     * @return drawing summary
     */
    public VisioDrawingInfo read(InputStream inputStream, String type) {
        if ("vsdx".equalsIgnoreCase(type)) {
            return readVsdx(inputStream);
        }
        if ("vsd".equalsIgnoreCase(type)) {
            return readVsd(inputStream);
        }
        throw new IllegalArgumentException("Unsupported Visio drawing type: " + type);
    }

    private VisioDrawingInfo readVsdx(InputStream inputStream) {
        try (XmlVisioDocument document = new XmlVisioDocument(inputStream)) {
            List<VisioPageInfo> pages = new ArrayList<>();
            List<String> textItems = new ArrayList<>();
            for (XDGFPage page : document.getPages()) {
                List<VisioShapeInfo> shapes = shapes(page.getContent().getTopLevelShapes(), textItems);
                Dimension2dDouble pageSize = page.getPageSize();
                pages.add(new VisioPageInfo(
                        page.getID(),
                        value(page.getName()),
                        page.getPageNumber(),
                        pageSize == null ? 0 : pageSize.getWidth(),
                        pageSize == null ? 0 : pageSize.getHeight(),
                        page.getContent().getConnections().size(),
                        shapes
                ));
            }
            return new VisioDrawingInfo("vsdx", pages, textItems);
        } catch (Exception e) {
            throw new VisioDrawingException("Read VSDX drawing failed", e);
        }
    }

    private VisioDrawingInfo readVsd(InputStream inputStream) {
        try (VisioTextExtractor extractor = new VisioTextExtractor(inputStream)) {
            String[] text = extractor.getAllText();
            return new VisioDrawingInfo(
                    "vsd",
                    Collections.<VisioPageInfo>emptyList(),
                    text == null ? Collections.<String>emptyList() : Arrays.asList(text)
            );
        } catch (Exception e) {
            throw new VisioDrawingException("Read VSD drawing failed", e);
        }
    }

    private List<VisioShapeInfo> shapes(List<XDGFShape> source, List<String> textItems) {
        if (source == null || source.isEmpty()) {
            return Collections.emptyList();
        }
        List<VisioShapeInfo> result = new ArrayList<>();
        for (XDGFShape shape : source) {
            String text = value(shape.getTextAsString());
            if (!text.isEmpty()) {
                textItems.add(text);
            }
            result.add(new VisioShapeInfo(
                    shape.getID(),
                    first(shape.getName(), shape.getShapeType(), shape.getSymbolName()),
                    value(shape.getType()),
                    text,
                    shape.getPinX(),
                    shape.getPinY(),
                    shape.getWidth(),
                    shape.getHeight(),
                    shapes(shape.getShapes(), textItems)
            ));
        }
        return result;
    }

    private String extension(String name) {
        int index = name == null ? -1 : name.lastIndexOf('.');
        return index < 0 ? "" : name.substring(index + 1);
    }

    private String first(String first, String second, String third) {
        if (first != null && !first.isEmpty()) {
            return first;
        }
        if (second != null && !second.isEmpty()) {
            return second;
        }
        return value(third);
    }

    private String value(String value) {
        return value == null ? "" : value;
    }
}

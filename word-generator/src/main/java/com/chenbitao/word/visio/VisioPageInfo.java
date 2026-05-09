package com.chenbitao.word.visio;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Basic Visio page information.
 */
public class VisioPageInfo {

    private final long id;
    private final String name;
    private final long pageNumber;
    private final double width;
    private final double height;
    private final int connectionCount;
    private final List<VisioShapeInfo> shapes;

    public VisioPageInfo(long id,
                         String name,
                         long pageNumber,
                         double width,
                         double height,
                         int connectionCount,
                         List<VisioShapeInfo> shapes) {
        this.id = id;
        this.name = name;
        this.pageNumber = pageNumber;
        this.width = width;
        this.height = height;
        this.connectionCount = connectionCount;
        this.shapes = unmodifiableCopy(shapes);
    }

    public long getId() {
        return id;
    }

    public String getName() {
        return name;
    }

    public long getPageNumber() {
        return pageNumber;
    }

    public double getWidth() {
        return width;
    }

    public double getHeight() {
        return height;
    }

    public int getConnectionCount() {
        return connectionCount;
    }

    public List<VisioShapeInfo> getShapes() {
        return shapes;
    }

    public int getShapeCount() {
        return count(shapes);
    }

    private int count(List<VisioShapeInfo> values) {
        int total = 0;
        for (VisioShapeInfo value : values) {
            total++;
            total += count(value.getChildren());
        }
        return total;
    }

    private static <T> List<T> unmodifiableCopy(List<T> values) {
        if (values == null || values.isEmpty()) {
            return Collections.emptyList();
        }
        return Collections.unmodifiableList(new ArrayList<>(values));
    }
}

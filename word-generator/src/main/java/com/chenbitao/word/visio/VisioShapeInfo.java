package com.chenbitao.word.visio;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Basic Visio shape information.
 */
public class VisioShapeInfo {

    private final long id;
    private final String name;
    private final String type;
    private final String text;
    private final Double pinX;
    private final Double pinY;
    private final Double width;
    private final Double height;
    private final List<VisioShapeInfo> children;

    public VisioShapeInfo(long id,
                          String name,
                          String type,
                          String text,
                          Double pinX,
                          Double pinY,
                          Double width,
                          Double height,
                          List<VisioShapeInfo> children) {
        this.id = id;
        this.name = name;
        this.type = type;
        this.text = text;
        this.pinX = pinX;
        this.pinY = pinY;
        this.width = width;
        this.height = height;
        this.children = unmodifiableCopy(children);
    }

    public long getId() {
        return id;
    }

    public String getName() {
        return name;
    }

    public String getType() {
        return type;
    }

    public String getText() {
        return text;
    }

    public Double getPinX() {
        return pinX;
    }

    public Double getPinY() {
        return pinY;
    }

    public Double getWidth() {
        return width;
    }

    public Double getHeight() {
        return height;
    }

    public List<VisioShapeInfo> getChildren() {
        return children;
    }

    private static <T> List<T> unmodifiableCopy(List<T> values) {
        if (values == null || values.isEmpty()) {
            return Collections.emptyList();
        }
        return Collections.unmodifiableList(new ArrayList<>(values));
    }
}

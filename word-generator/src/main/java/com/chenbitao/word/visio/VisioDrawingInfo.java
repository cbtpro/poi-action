package com.chenbitao.word.visio;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Visio drawing summary information.
 */
public class VisioDrawingInfo {

    private final String format;
    private final List<VisioPageInfo> pages;
    private final List<String> textItems;

    public VisioDrawingInfo(String format, List<VisioPageInfo> pages, List<String> textItems) {
        this.format = format;
        this.pages = unmodifiableCopy(pages);
        this.textItems = unmodifiableCopy(textItems);
    }

    public String getFormat() {
        return format;
    }

    public List<VisioPageInfo> getPages() {
        return pages;
    }

    public int getPageCount() {
        return pages.size();
    }

    public int getShapeCount() {
        int total = 0;
        for (VisioPageInfo page : pages) {
            total += page.getShapeCount();
        }
        return total;
    }

    public List<String> getTextItems() {
        return textItems;
    }

    private static <T> List<T> unmodifiableCopy(List<T> values) {
        if (values == null || values.isEmpty()) {
            return Collections.emptyList();
        }
        return Collections.unmodifiableList(new ArrayList<>(values));
    }
}

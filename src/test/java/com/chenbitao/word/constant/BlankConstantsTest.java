package com.chenbitao.word.constant;

import org.junit.Test;

import static org.junit.Assert.assertEquals;

public class BlankConstantsTest {

    @Test
    public void commonBlankConstantsHaveExpectedValues() {
        assertEquals("", BlankConstants.EMPTY_STRING);
        assertEquals("", BlankConstants.EMPTY);
        assertEquals("-", BlankConstants.DASH);
        assertEquals("无", BlankConstants.NONE);
    }
}

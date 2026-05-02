package com.chenbitao.word.constant;

import org.junit.Test;

import static org.junit.Assert.assertEquals;

/**
 * BlankConstants常量类测试
 * 测试BlankConstants中定义的空白值常量是否具有预期的值
 */
public class BlankConstantsTest {

    /**
     * 测试常用空白常量具有预期的值
     * 验证EMPTY_STRING、EMPTY、DASH和NONE常量的值是否正确
     */
    @Test
    public void commonBlankConstantsHaveExpectedValues() {
        assertEquals("", BlankConstants.EMPTY_STRING);
        assertEquals("", BlankConstants.EMPTY);
        assertEquals("-", BlankConstants.DASH);
        assertEquals("无", BlankConstants.NONE);
    }
}

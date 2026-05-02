package com.chenbitao.word.constant;

/**
 * 空白常量定义类
 * 定义Word文档生成过程中常用的空白值常量，如空字符串、破折号、无等
 */
public final class BlankConstants {

    /** 空字符串常量 */
    public static final String EMPTY_STRING = "";

    /** 空常量（与EMPTY_STRING相同） */
    public static final String EMPTY = "";

    /** 破折号常量，用于表示未填写或默认值 */
    public static final String DASH = "-";

    /** "无"常量，用于表示没有相关信息 */
    public static final String NONE = "无";

    /**
     * 私有构造方法，防止实例化
     */
    private BlankConstants() {
    }
}

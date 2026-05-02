package com.chenbitao.word.core;

/**
 * 抽象Word文档生成器
 * 提供Word文档生成器的基础实现，包含字体设置等通用功能
 */
public abstract class AbstractWordGenerator implements WordGenerator {

    /** 默认字体名称 */
    protected String font = "宋体";

    /**
     * 设置文档字体
     *
     * @param font 字体名称
     */
    public void setFont(String font) {
        this.font = font;
    }
}
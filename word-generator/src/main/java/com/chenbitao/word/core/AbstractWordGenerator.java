package com.chenbitao.word.core;

import lombok.Setter;

/**
 * 抽象Word文档生成器
 * 提供Word文档生成器的基础实现，包含字体设置等通用功能
 */
@Setter
public abstract class AbstractWordGenerator implements WordGenerator {

    /** 默认字体名称
     * -- SETTER --
     *  设置文档字体
     */
    protected String font = "宋体";

}
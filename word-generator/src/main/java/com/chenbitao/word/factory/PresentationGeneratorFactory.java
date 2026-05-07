package com.chenbitao.word.factory;

import com.chenbitao.word.presentation.HslfPresentationGenerator;
import com.chenbitao.word.presentation.PresentationGenerator;

/**
 * 演示文稿生成器工厂。
 *
 * <p>当前支持 PowerPoint 97-2003 的 {@code ppt} 类型。</p>
 */
public class PresentationGeneratorFactory {

    private PresentationGeneratorFactory() {
    }

    /**
     * 获取指定类型的 PowerPoint 生成器。
     *
     * @param type 文件类型，目前支持 {@code ppt}
     * @return 新的 PowerPoint 生成器实例
     */
    public static PresentationGenerator get(String type) {
        if ("ppt".equalsIgnoreCase(type)) {
            return new HslfPresentationGenerator();
        }
        throw new IllegalArgumentException("不支持演示文稿类型: " + type);
    }
}

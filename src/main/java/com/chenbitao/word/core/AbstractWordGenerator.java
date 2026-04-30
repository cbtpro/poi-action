package com.chenbitao.word.core;

public abstract class AbstractWordGenerator implements WordGenerator {

    protected String font = "宋体";

    public void setFont(String font) {
        this.font = font;
    }
}
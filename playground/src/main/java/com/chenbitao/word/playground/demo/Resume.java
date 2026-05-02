package com.chenbitao.word.playground.demo;

import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * 简历数据模型类
 * 包含简历的基本信息，如姓名、联系方式、教育经历、工作经历和技能等
 */
@Data
@Builder
public class Resume {

    /** 姓名 */
    public String name;

    /** 电话号码 */
    public String phone;

    /** 邮箱地址 */
    public String email;

    /** 教育经历列表 */
    public List<String> education;

    /** 工作经历列表 */
    public List<String> experience;

    /** 技能列表 */
    public List<String> skills;
}
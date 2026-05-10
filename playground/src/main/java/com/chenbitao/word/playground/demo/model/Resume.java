package com.chenbitao.word.playground.demo.model;

import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * 简历下载接口的演示数据模型。
 *
 * <p>当前 Controller 仍复用干部模板生成文件，该模型预留给后续简历模板演示使用。</p>
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

package com.cy.easyexcel;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

/**
 * @description:
 * @projectName:poieasyexcel
 * @see:com.cy.easyexcel
 * @author:chengyang
 * @createTime:2020/4/29 11:08 上午
 * @version:1.0
 */

@Data
public class DemoData {
    @ExcelProperty("字符串标题")
    private String string;
    @ExcelProperty("日期标题")
    private Date date;
    @ExcelProperty("数字标题")
    private Double doubleData;
    /**
     * 忽略这个字段
     */
    @ExcelIgnore
    private String ignore;
}


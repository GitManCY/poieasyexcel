package com.cy.poi;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @description:
 * @projectName:poieasyexcel
 * @see:com.cy
 * @author:chengyang
 * @createTime:2020/4/29 9:18 上午
 * @version:1.0
 */
public class ExcelWriteTest {

    private static final String PATH = "/Users/a123/Desktop/my/Idea_workspace/poieasyexcel/src/main/resources/";

    @Test
    public void testWrite2003() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("观众统计表");
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        cell1.setCellValue("今日新增观众");
        Cell cell2 = row.createCell(0);
        cell2.setCellValue("统计时间");

        Row row2 = sheet.createRow(1);
        Cell cell4 = row2.createCell(1);
        cell4.setCellValue(5);
        Cell cell3 = row2.createCell(1);
        cell3.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "观众统计表2003.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

    }

    @Test
    public void testWrite2007() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("观众统计表");
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        cell1.setCellValue("今日新增观众");
        Cell cell2 = row.createCell(1);
        cell2.setCellValue("统计时间");

        Row row2 = sheet.createRow(1);
        Cell cell4 = row2.createCell(0);
        cell4.setCellValue(5);
        Cell cell3 = row2.createCell(1);
        cell3.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "观众统计表2007.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

    }
}

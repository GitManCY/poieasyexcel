package com.cy.poi;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @description:
 * @projectName:poieasyexcel
 * @see:com.cy
 * @author:chengyang
 * @createTime:2020/4/29 9:37 上午
 * @version:1.0
 */
public class BigDataWriteTest {
    private static final String PATH = "/Users/a123/Desktop/my/Idea_workspace/poieasyexcel/src/main/resources/";

    @Test
    public void BigDataWriteTest2003() throws IOException {
        long begin = System.currentTimeMillis();
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        FileOutputStream stream = new FileOutputStream(PATH + "写入大量数据2003.xls");
        workbook.write(stream);
        stream.close();
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }

    @Test
    public void BigDataWriteTest2007() throws IOException {
        long begin = System.currentTimeMillis();
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 100000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        FileOutputStream stream = new FileOutputStream(PATH + "写入大量数据2007.xls");
        workbook.write(stream);
        stream.close();
        workbook.dispose();
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }
}

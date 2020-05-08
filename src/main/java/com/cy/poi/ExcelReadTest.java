package com.cy.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Date;

import static org.apache.poi.ss.usermodel.Cell.*;

/**
 * @description:
 * @projectName:poieasyexcel
 * @see:com.cy.poi
 * @author:chengyang
 * @createTime:2020/4/29 9:59 上午
 * @version:1.0
 */
public class ExcelReadTest {

    private static final String PATH = "/Users/a123/Desktop/my/Idea_workspace/poieasyexcel/src/main/resources/";

    @Test
    public void testRead2003() throws Exception {

        FileInputStream fileInputStream = new FileInputStream(PATH + "观众统计表2003.xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();
    }

    @Test
    public void testReadByUNKnowType() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(PATH + "观众统计表2007.xls");
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        FormulaEvaluator formulaEvaluator = new
                HSSFFormulaEvaluator((HSSFWorkbook) workbook);
        Sheet sheet = workbook.getSheetAt(0);
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null) {
            int cellCount = rowTitle.getPhysicalNumberOfCells();    //拿到第row行的那一行的总个数
            for (int i = 0; i < cellCount; i++) {  //循环个数取出
                Cell cell = rowTitle.getCell(i);
                if (cell != null) {          //如果不等于空取出值
                    int cellType = cell.getCellType();   //这里是知道我们标题是String，考虑不确定的时候怎么取
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "|");
                }
            }
            System.out.println();
        }
        //获取表中内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum); //取出对应的行
            if (rowData != null) {
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("[" + (rowNum + 1 + "-" + (cellNum + 1) + "]"));
                    Cell cell = rowData.getCell(cellNum);
                    //匹配数据类型
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        switch (cellType) {
                            case XSSFCell.CELL_TYPE_STRING:
                                System.out.print("字符串：" + cell.getStringCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.print("布尔：" + cell.getBooleanCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    System.out.println("日期格式：" + new DateTime(cell.getDateCellValue()).toString("yyyy-MM-dd HH:mm:ss"));
                                    break;
                                } else
                                    cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                                System.out.print("整形：" + cell.toString());
                                break;
                            case XSSFCell.CELL_TYPE_BLANK:
                                System.out.print("空");
                                break;
                            case XSSFCell.CELL_TYPE_ERROR:
                                System.out.print("数据类型错误");
                                break;
                            case Cell.CELL_TYPE_FORMULA:
                                String formula = cell.getCellFormula();
                                System.out.println("公式:" + formula);
                                CellValue evaluate = formulaEvaluator.evaluate(cell);
                                String cellValue = evaluate.formatAsString();
                                System.out.println(cellValue);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
        }
        fileInputStream.close();
    }
}
package com.fcg.exceldecment;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@SpringBootTest
class ExceldecmentApplicationTests {

    //读取工作簿
    @Test
    void excelwork() throws IOException {
        //获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook("J:\\Users\\LYC\\Desktop\\工作\\2021年节假日道路旅客运输统计工作表.xlsx");
        //获取工作表
        XSSFSheet sheetAt = workbook.getSheet("Sheet1");
        //获取每一行
        for (Row cells : sheetAt) {
            //获取单元格
            for (Cell cell : cells) {
                //获取内容
                String result = cell.getStringCellValue();
                System.out.println(result);
            }
        }
        workbook.close();
    }

    //创建一个带皮肤的excel
    @Test
    void createexcel (){

    }
    //写入工作簿
    @Test
    void createSheet() throws IOException {
        //创建工作簿
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //创建工作表
        XSSFSheet sheet = xssfWorkbook.createSheet("Sheet1");
        //创建工作行
        XSSFRow row = sheet.createRow(0);
        //创建单元格
        row.createCell(0).setCellValue("aa1");
        row.createCell(1).setCellValue("aa2");
        row.createCell(2).setCellValue("aa3");

        //创建工作行
        XSSFRow row1 = sheet.createRow(1);
        //创建单元格
        row1.createCell(0).setCellValue("bb1");
        row1.createCell(1).setCellValue("bb2");
        row1.createCell(2).setCellValue("bb3");

        FileOutputStream outputStream = new FileOutputStream("J:\\Users\\LYC\\Desktop\\工作\\new hello.xlsx");
        xssfWorkbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        xssfWorkbook.close();
    }


    //合并单元格
    @Test
    void hebing() throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            HSSFSheet sheet = wb.createSheet("new sheet");

            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue("This is a test of merging");

            //合并的单元格 （参数1，参数2 ，参数3，参数4，） 参数1-2位坐标；参数3为起始位置；参数4为合并数量
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 7));
            FileOutputStream fileOut = new FileOutputStream("J:\\Users\\LYC\\Desktop\\工作\\workbook.xls");
            wb.write(fileOut);
        }
    }

    //编写一个工作表
    void readwork() throws IOException {
    }
}
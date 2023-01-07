package com.fcg.exceldecment;

import com.fcg.exceldecment.pojo.Tablesz;
import javafx.scene.layout.Border;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@SpringBootTest
class ExceldecmentApplicationTests {

    //循环给对象赋值
    @Test
    void pojoforech(){
        List<String> list = new ArrayList<String>();
        list.add("zs");
        list.add("ls");
        list.add("ww");
        list.add("zl");
        Tablesz poj = new Tablesz();
        Field[] declaredFields = poj.getClass().getDeclaredFields();
        for (Field declaredField : declaredFields) {
            System.out.println(declaredField.getName());
        }



    }

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
                System.out.println(cell.getStringCellValue());
            }
        }
        workbook.close();
    }

    //创建一个带皮肤的excel
    @Test
    void createexcel () throws IOException {
        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        Color color = new XSSFColor(new java.awt.Color(22, 253, 22));
        short col = 0;
        //这是样式
        CellStyle style = workbook.createCellStyle();
        /*//背景填充
        style.setFillBackgroundColor(col);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);*/
        /*边框填充
        参数THICK MEDIUM为厚的

        参数DOTTED
           SLANTED_DASH_DOT
           DASH_DOT_DOT
           MEDIUM_DASH_DOT_DOT
           DASHED为虚线
        */
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor((short) 0);
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor((short) 0);

        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor((short) 0);

        style.setBorderRight(style.getBorderTopEnum());
        style.setRightBorderColor((short) 0);
        //输出一些参数
        System.out.println(style.getBorderTopEnum());
        //创建一个sheet
        Sheet sheet = workbook.createSheet("sheet1");
        //创建工作行
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("00123123131");
        cell.setCellStyle(style);
        cell.getCellStyle();

        row.createCell(1).setCellValue("01");
        row.createCell(2).setCellValue("02");
        row.createCell(3).setCellValue("03");
        row.createCell(4).setCellValue("04");

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("10");
        row1.createCell(1).setCellValue("11");
        row1.createCell(2).setCellValue("12");
        row1.createCell(3).setCellValue("13");
        row1.createCell(4).setCellValue("14");



        FileOutputStream outputStream = new FileOutputStream("E:\\textwork\\new work.xlsx");
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        workbook.close();
    }

    //边框样式
    @Test
    void createstyle () throws IOException {
        //创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建一个sheet
        Sheet sheet = workbook.createSheet("sheet1");
        //这是样式
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBottomBorderColor((short) 0);
        style.setLeftBorderColor((short) 0);
        style.setTopBorderColor((short) 0);
        style.setRightBorderColor((short) 0);

        //创建工作行
        for (int ii = 0; ii < 5; ii++) {
            Row row = sheet.createRow(ii);
            for (int i = 0; i < 5; i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(style);

            }
        }
        FileOutputStream outputStream = new FileOutputStream("E:\\textwork\\new work.xlsx");
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        workbook.close();
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
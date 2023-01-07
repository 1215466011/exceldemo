package com.fcg.exceldecment;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

public class WorkDemo {
    public static void main(String[] args) throws IOException {
        String fromPath = "E:\\textwork";// excel存放路径
        String toPath = "E:\\new work";// 保存新EXCEL路径

        String excelName = "sum.xlsx";
        Workbook workbook = new XSSFWorkbook();
        double sum = 0;

        File file = new File(fromPath);
        for (File excel : file.listFiles()) {
            String strExcelPath = fromPath + "\\" + excel.getName();
            XSSFWorkbook wb  = new XSSFWorkbook(strExcelPath);
            XSSFSheet wbSheet = wb.getSheet("Sheet1");
            String value = wbSheet.getRow(0).getCell(0).getRawValue();
            System.out.println("value = "+value);
            sum+=Double.parseDouble(value);
        }
        System.out.println(sum);
    }
}

package com.fcg.exceldecment;

import com.sun.scenario.effect.impl.sw.sse.SSEBlend_SRC_OUTPeer;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.w3c.dom.css.RGBColor;

import javax.lang.model.element.VariableElement;
import java.io.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MyWork {
    public static void main(String[] args) throws IOException {
        String fromPath = "E:\\textwork";// excel存放路径
        String toPath = "E:\\new work";// 保存新EXCEL路径
        // 新的excel 文件名
        String excelName = "";
        // 创建新的excel
        XSSFWorkbook wbCreat = new XSSFWorkbook();
        File file = new File(fromPath);
        for (File excel : file.listFiles()) {
        // 打开已有的excel
            excelName = excel.getName();
            String strExcelPath = fromPath + "\\" + excel.getName();
            XSSFWorkbook wb = new XSSFWorkbook(strExcelPath);
            XSSFCellStyle wbStyle = wb.createCellStyle();
            for (int ii = 0; ii < wb.getNumberOfSheets(); ii++) {
                XSSFSheet sheet = wb.getSheet("Sheet"+(ii+1));

                XSSFSheet sheetCreat = wbCreat.createSheet(sheet.getSheetName());
                // 复制源表中的合并单元格
                MergerRegion(sheetCreat, sheet);
                int firstRow = sheet.getFirstRowNum();
                int lastRow = sheet.getLastRowNum();

                for (int i = firstRow; i <= lastRow; i++) {
                // 创建新建excel Sheet的行
                    XSSFRow rowCreat = sheetCreat.createRow(i);
                // 取得源有excel Sheet的行
                    XSSFRow row = sheet.getRow(i);
                // 单元格式样
                    int firstCell = row.getFirstCellNum();
                    int lastCell = row.getLastCellNum();
                    //自适应列宽
                    for (int nub = 0;nub<=row.getPhysicalNumberOfCells();nub++){
                        sheetCreat.autoSizeColumn(nub);
                        sheetCreat.autoSizeColumn(nub, true);
                    }
                    for (int j = firstCell; j < lastCell; j++) {
                        //新格子
                        rowCreat.createCell(j);
                        XSSFCell tocell =rowCreat.getCell(j);
                        //去的源格子
                        XSSFCell fromcell = row.getCell(j);
                        //自适应行高
                        row.setZeroHeight(true);
                        //自适应列宽
                        rowCreat.createCell(j);
                        if (row.getCell(j).getRawValue()==null ||"".equals(row.getCell(j).getRawValue())) {}else{
                            String cetype = fromcell.getCellTypeEnum().toString();
                            fromcell.setCellType( fromcell.getCellTypeEnum());
                            if("STRING".equals(cetype)){
                                rowCreat.getCell(j).setCellValue(fromcell.getStringCellValue());
                            }else{
                                rowCreat.getCell(j).setCellValue(fromcell.getNumericCellValue());
                            }
                        }
                        CellStyle style = wbCreat.createCellStyle();
                        copyCellStyle(fromcell.getCellStyle(),style);
                        tocell.setCellStyle(style);
                    }
                }
            }
            wb.close();
        }

        String path = toPath+"\\"+excelName+"";
        FileOutputStream fileOut = new FileOutputStream(path);
        wbCreat.write(fileOut);
        fileOut.close();

        wbCreat.close();
    }
/*
    public static void copyCellStyle(XSSFWorkbook workbook, XSSFCellStyle fromStyle, XSSFCellStyle toStyle) {
        //背景和前景
        */
/*if(fromStyle instanceof  XSSFCellStyle){
            if(fromStyle.getFillBackgroundColorColor()!=null){
                System.out.println(fromStyle.getFillBackgroundColorColor());
            }
            if(fromStyle.getFillForegroundColorColor()!=null){
                System.out.println(fromStyle.getFillForegroundColorColor());
            }
            toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColorColor());
            toStyle.setFillForegroundColor(fromStyle.getFillForegroundColorColor());
        }else {
            toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
            toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());
        }
        toStyle.setDataFormat(fromStyle.getDataFormat());
        toStyle.setFillPattern(fromStyle.getFillPatternEnum());*//*

//    toStyle.setFont(fromStyle.getFont(null)); // 没有提供get 方法
       */
/* if (fromStyle instanceof XSSFCellStyle) {
            // 处理字体获取：03版 xls
            XSSFCellStyle style = (XSSFCellStyle) fromStyle;
            toStyle.setFont(style.getFont());
        } else if (fromStyle instanceof XSSFCellStyle) {
            // 处理字体获取：07版以及之后 xlsx
            XSSFCellStyle style = (XSSFCellStyle) fromStyle;
            toStyle.setFont(style.getFont());
        }*//*

        toStyle.setHidden(fromStyle.getHidden());
        toStyle.setIndention(fromStyle.getIndention());//首行缩进
        toStyle.setLocked(fromStyle.getLocked());
        toStyle.setRotation(fromStyle.getRotation());//旋转
        toStyle.setWrapText(fromStyle.getWrapText());
    }
*/

    public static void copyCellStyle(XSSFCellStyle fromStyle, CellStyle style) {

        // 水平垂直对齐方式
        style.setAlignment(fromStyle.getAlignmentEnum());
        style.setVerticalAlignment(fromStyle.getVerticalAlignmentEnum());

        //边框和边框颜色
        style.setBorderBottom(fromStyle.getBorderBottomEnum());
        style.setBorderLeft(fromStyle.getBorderLeftEnum());
        style.setBorderRight(fromStyle.getBorderRightEnum());
        style.setBorderTop(fromStyle.getBorderTopEnum());

        style.setTopBorderColor(fromStyle.getTopBorderColor());
        style.setBottomBorderColor(fromStyle.getBottomBorderColor());
        style.setRightBorderColor(fromStyle.getRightBorderColor());
        style.setLeftBorderColor(fromStyle.getLeftBorderColor());

    }
     private static void MergerRegion(XSSFSheet sheetCreat, XSSFSheet sheet) {
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            sheetCreat.addMergedRegion(mergedRegion);
        }
    }
    /**
     * 去除字符串内部空格
     */
    public static String removeInternalBlank(String s) {
// System.out.println("bb:" + s);
        Pattern p = Pattern.compile("\\s*|\t|\r|\n");
        Matcher m = p.matcher(s);
        char str[] = s.toCharArray();
        StringBuffer sb = new StringBuffer();
        for (int i = 0; i < str.length; i++) {
            if (str[i] == ' ') {
                sb.append(' ');
            } else {
                break;
            }
        }
        String after = m.replaceAll("");
        return sb.toString() + after;
    }
}

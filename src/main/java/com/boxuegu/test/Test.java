package com.boxuegu.test;

import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author Xiang想
 * @title: Test
 * @projectName POITest
 * @description: TODO
 * @date 2020/11/29  3:00
 */
public class Test {
    public static void main(String[] args) {
        String homePath = System.getProperty("user.dir");
        String filePath = File.separator + "src" + File.separator + "main" +
                File.separator + "resources" + File.separator;

        String readPath = homePath+filePath+"1.xlsx";
        String writePath = homePath+filePath+"5.xlsx";
        SendExcel(readPath,writePath);

    }

    public static void SendExcel(String readPath,String writePath){
        XSSFWorkbook readWorkBook = null;
        XSSFSheet readSheet = null;
        try {
            readWorkBook = new XSSFWorkbook(readPath);
            readSheet = readWorkBook.getSheetAt(0);
            forEachRead(readSheet,writePath);
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (readWorkBook!=null){
                try {
                    readWorkBook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void forEachRead(XSSFSheet sheet,String writePath){

        // 创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建工作表
        XSSFSheet writeSheet = workbook.createSheet("数据报送");

        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFRow writeRow = writeSheet.createRow(i);
            if (row!=null){
                short lastCellNum = row.getLastCellNum();
                for (int j = 0; j <= lastCellNum; j++) {
                    XSSFCell writeRowCell = writeRow.createCell(j);
                    XSSFCell cell = row.getCell(j);
                    if (cell!=null){
                        String value = cell.getStringCellValue();
                        XSSFCellStyle cellStyle = cell.getCellStyle();
                        writeRowCell.getCellStyle().cloneStyleFrom(cellStyle);
                        writeRowCell.setCellValue(value);
                    }
                }
            }
        }

        // 创建输出流对象
        try (FileOutputStream out = new FileOutputStream(writePath)) {
            workbook.write(out);
            out.flush();
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if (workbook!=null){
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }


    }
}

package com.boxuegu.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

/**
 * @author Xiang想
 * @title: Demo
 * @projectName POITest
 * @description: TODO
 * @date 2020/11/29  1:21
 */
public class Demo {

    public static void main(String[] args) throws IOException {
        String homePath = System.getProperty("user.dir");
        String filePath = File.separator + "src" + File.separator + "main" +
                File.separator + "resources" + File.separator;
        String path =  homePath+filePath+"1.xlsx";
        readExcel(path);
    }

    public static void readExcel(String filePath){
        XSSFWorkbook workbook = null;
        XSSFSheet sheet = null;
        try {
            // 获取工作簿
            workbook = new XSSFWorkbook(filePath);

            // 获取工作表
            sheet = workbook.getSheetAt(0);

            forEachRead(sheet);
//            forRead(sheet);

        } catch (IOException e) {
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

    public static void forEachRead(XSSFSheet sheet){
        // 获取行
        for (Row row : sheet) {
            // 获取单元格
            for (Cell cell : row) {
                // 获取单元格中的内容
                String value = cell.getStringCellValue();
                if (value.equals("")){
                    continue;
                }
                System.out.println(value);
            }
        }
    }

    public static void forRead(XSSFSheet sheet){
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row!=null){
                short lastCellNum = row.getLastCellNum();
                for (int j = 0; j <= lastCellNum; j++) {
                    XSSFCell cell = row.getCell(j);
                    if (cell!=null){
                        String value = cell.getStringCellValue();
                        System.out.println(value);
                    }
                }
            }
        }
    }

}

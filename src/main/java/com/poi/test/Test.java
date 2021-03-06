package com.poi.test;

import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * resources 只保留 1.xlsx 文件 如果存在 5.xlsx 请删除后再执行
 *
 *
 * @author Xiang想
 * @title: Test
 * @projectName POITest
 * @description: TODO
 * @date 2020/11/29  3:00
 */
public class Test {
    public static void main(String[] args) {
        // 获取当前项目位置
        String homePath = System.getProperty("user.dir");
        String filePath = File.separator + "src" + File.separator + "main" +
                File.separator + "resources" + File.separator;

        // 源文件
        String readPath = homePath+filePath+"1.xlsx";
        // 生成文件
        String writePath = homePath+filePath+"5.xlsx";
        SendExcel(readPath,writePath);

    }

    public static void SendExcel(String readPath,String writePath){
        XSSFWorkbook readWorkBook = null;
        XSSFSheet readSheet = null;
        try {
            readWorkBook = new XSSFWorkbook(readPath);
            readSheet = readWorkBook.getSheetAt(0);
            // 循环读取
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
        // 获取最后一行
        int lastRowNum = sheet.getLastRowNum();
        // 循环行
        for (int i = 0; i <= lastRowNum; i++) {
            // row : 源行
            XSSFRow row = sheet.getRow(i);
            // writeRow 生成行
            XSSFRow writeRow = writeSheet.createRow(i);
            if (row!=null){
                short lastCellNum = row.getLastCellNum();
                for (int j = 0; j <= lastCellNum; j++) {
                    // cell 源行
                    XSSFCell cell = row.getCell(j);
                    // writeRowCell 生成后文件行
                    XSSFCell writeRowCell = writeRow.createCell(j);
                    if (cell!=null){
                        // 从源文件行中读取值
                        String value = cell.getStringCellValue();
                        // 从源文件行中获取样式
                        XSSFCellStyle cellStyle = cell.getCellStyle();


                        // 这里的样式，我无法加载到 新的Excel中
                        writeRowCell.getCellStyle().cloneStyleFrom(cellStyle);

                        // 模拟更改文件内容
                        if (i==1&&j==0){
                            writeRowCell.setCellValue("某某保险公司");
                            continue;
                        }

                        // 把值写入 生成文件
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

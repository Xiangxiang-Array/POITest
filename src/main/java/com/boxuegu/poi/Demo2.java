package com.boxuegu.poi;

import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author Xiang想
 * @title: Demo2
 * @projectName POITest
 * @description: TODO
 * @date 2020/11/29  2:01
 */
public class Demo2 {
    public static void main(String[] args) {
        String homePath = System.getProperty("user.dir");
        String filePath = File.separator + "src" + File.separator + "main" +
                File.separator + "resources" + File.separator;
        String path =  homePath+filePath+"4.xlsx";
        write(path);

    }

    public static List<String> getMessage(){
        List<String> list = new ArrayList<>();
        list.add("传智播客");
        list.add("黑马程序员");
        list.add("博学谷");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String date = simpleDateFormat.format(new Date());
        list.add(date);
        return list;
    }


    public static void write(String filePath){

        List<String> message = getMessage();

        // 创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建工作表
        XSSFSheet sheet = workbook.createSheet("工作表");
        // 创建行
        XSSFRow row = sheet.createRow(0);



        // 创建单元格
        for (int i = 0; i < message.size(); i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(message.get(i));
        }

        // 创建输出流对象
        try (FileOutputStream out = new FileOutputStream(filePath)) {

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

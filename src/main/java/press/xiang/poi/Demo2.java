package press.xiang.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 新建Excel文件
 * @author Xiang想
 * @title: Demo2
 * @projectName POITest
 * @description: TODO
 * @date 2020/11/28  23:46
 */
public class Demo2 {
    public static void main(String[] args) {
        String file = "F:/study/IDEA/LeetCode/POITest/src/main/resources/3.xlsx";
        try {
            writerSheel(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 创建Excel文件
     * @param path 路径
     * @throws IOException
     */
    public static void  writerExcel(String path) throws IOException {
        //创建Excel文件(Workbook)
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建工作表(Sheet)
//        HSSFSheet sheet = workbook.createSheet();
        //创建工作表(Sheet)
//        sheet = workbook.createSheet("Test");
        FileOutputStream out = new FileOutputStream(path);
        //保存Excel文件
        workbook.write(out);
        //关闭文件流
        out.close();
        System.out.println("OK!");
        File file = new File(path);
        System.out.println(file.delete()?"删除成功":"删除失败");
    }


    public static void writerSheel(String path) throws IOException {

        FileOutputStream out = new FileOutputStream(path);


        //创建Excel文件(Workbook)
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建工作表(Sheet)
        HSSFSheet sheet = workbook.createSheet("Test");
        // 创建行,从0开始
        HSSFRow row = sheet.createRow(3);
        // 创建行的单元格,也是从0开始
        HSSFCell cell = row.createCell(0);
        // 设置单元格内容
        cell.setCellValue("李志伟");
        // 设置单元格内容,重载
        row.createCell(1).setCellValue(false);
        // 设置单元格内容,重载

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMdd HH:mm:ss");
        String format = simpleDateFormat.format(new Date());

        row.createCell(2).setCellValue(format);
        // 设置单元格内容,重载
        row.createCell(3).setCellValue(12.345);

        //保存Excel文件
        workbook.write(out);
        //关闭文件流
        out.close();
        System.out.println("OK!");

    }
}

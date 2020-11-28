package press.xiang.poi.test;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

/**
 * @author Xiang想
 * @title: Test
 * @projectName POITest
 * @description: TODO
 * @date 2020/11/29  0:11
 */
public class Test {
    public static void main(String[] args) {
        String in = "F:/study/IDEA/LeetCode/POITest/src/main/resources/2.xlsx";
        String out = "F:/study/IDEA/LeetCode/POITest/src/main/resources/4.xlsx";
        try {
            writeExcel(in,out);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    public static void writeExcel(String  inpath,String outpath) throws Exception{

        FileInputStream is =  new FileInputStream(inpath);
        FileOutputStream out = new FileOutputStream(outpath);
        HSSFWorkbook excel=new HSSFWorkbook(is);
        //创建Excel文件(Workbook)
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建工作表(Sheet)
        HSSFSheet sheet = workbook.createSheet("Test");



        int rowNum = 0;
        int cellNum = 0;


        //获取第一个sheet
        HSSFSheet sheet0=excel.getSheetAt(0);
        for (Iterator rowIterator = sheet0.iterator(); rowIterator.hasNext();)
        {
            XSSFRow row=(XSSFRow) rowIterator.next();
            // 创建行,从0开始
            HSSFRow inRow = sheet.createRow(rowNum);
            for (Iterator iterator=row.cellIterator();iterator.hasNext();)
            {
                String text = "";
                // 创建行的单元格,也是从0开始
                HSSFCell inCell = inRow.createCell(cellNum);
                XSSFCell cell=(XSSFCell) iterator.next();
                //根据单元的的类型 读取相应的结果
                if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
                    text = cell.getStringCellValue()+"\t";
                }
                else if(cell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){
                    text = cell.getNumericCellValue()+"\t";
                }
                else if(cell.getCellType()==HSSFCell.CELL_TYPE_FORMULA) {
                    text = cell.getCellFormula()+"\t";
                }
                inRow.createCell(cellNum).setCellValue(text);
                cellNum++;
            }
            rowNum++;
        }
        workbook.write(out);
        //关闭文件流
        out.close();
        is.close();
        System.out.println("OK!");

    }
}

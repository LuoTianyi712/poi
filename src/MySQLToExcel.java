import bean.StudentBean;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import util.DBHelper;

import java.io.*;
import java.sql.ResultSet;
import java.util.Scanner;

public class MySQLToExcel {
//    public static void main(String [] args){
//
//        String [] title = {"id", "name", "sex", "birth"};
//
//        HSSFWorkbook workbook = new HSSFWorkbook();
//        HSSFSheet sheet = workbook.createSheet("student");
//
//        HSSFRow sheetRow = sheet.createRow(0);
//        HSSFCell sheetCell = null;
//
//        for (int row = 0; row < 5; row++) {
//            sheetRow = sheet.createRow(row);
//            for (int column = 0; column < title.length; column++) {
//                sheetRow.createCell(column).setCellValue(title[column]);
//            }
//        }
//
//        try {
//            workbook.write(new FileOutputStream(new File("test.xls")));
//            workbook.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//    }

    public static void main(String [] args)
    {
        String [] title = {"id", "name", "sex", "birth"};

        FileOutputStream studentExcel;

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("student");

        HSSFRow sheetRow = sheet.createRow(0);
        HSSFCell sheetCell = null;

        sheetRow.createCell(0).setCellValue("id");
        sheetRow.createCell(1).setCellValue("name");
        sheetRow.createCell(2).setCellValue("sex");
        sheetRow.createCell(3).setCellValue("birth");

//        Scanner

        for (int row = 1; row < 2; row++)
        {
            sheetRow = sheet.createRow(row);

            for (int column = 0; column < title.length; column++)
            {
                sheetRow.createCell(column).setCellValue(title[column]);
            }
        }

        try {
            studentExcel = new FileOutputStream("test.xls");
            workbook.write(studentExcel);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    DBHelper db = new DBHelper();

    String [] title = {"id", "name", "sex", "birth"};

    FileOutputStream studentExcel;

    HSSFWorkbook workbook = new HSSFWorkbook();//new出来的对象，是赋值给了workbook这个变量，存储在内存里面
                                                //我可以做什么？我可以对这个变量做什么，我可以调用里面的函数，可以对其进行执行

    HSSFSheet sheet = workbook.createSheet("student");//创建表

    HSSFRow sheetRow = sheet.createRow(0);
    HSSFCell sheetCell = null;

    public void insertDataToExcel(StudentBean studentBean){

        sheetRow.createCell(0).setCellValue("id");
        sheetRow.createCell(1).setCellValue("name");
        sheetRow.createCell(2).setCellValue("sex");
        sheetRow.createCell(3).setCellValue("birth");

        //数据库添加数据在此执行
        db.connectToMySql();

        try {

            for (int row = 1; row < db.countRow(); row++)//这里则需要改成行数
            {
                //在这里创建对象的意义是？对象创建出来了，赋值给变量？
               HSSFRow sheetRow = sheet.createRow(db.countRow());

                for (int column = 0; column < db.countColumn(); column++)//这里需要把条件改成列数，怎么获取列数
                {
                    //有问题，数据库数据读取到了，没有存放在excel中，
                    // 即将编写一个list，将数据存入list中，再读取list
//                    sheetRow.createCell(column).setCellValue(columnNum);

                    sheetRow.createCell(db.countColumn()).setCellValue(studentBean.getStuId());
                    sheetRow.createCell(db.countColumn()).setCellValue(studentBean.getStuName());
                    sheetRow.createCell(db.countColumn()).setCellValue(studentBean.getStuSex());
                    sheetRow.createCell(db.countColumn()).setCellValue(studentBean.getStuBirth());
                }
            }

//           int i = 1;
//            HSSFRow hssfRow = sheet.createRow(i);
//
//            for (int j = 0; j < db.countColumn() ; j++)
//            {
//                HSSFCell hssfCell = hssfRow.createCell(j);
//                hssfCell.setCellValue(studentBean.getStuId());
//                hssfCell.setCellValue(studentBean.getStuName());
//                hssfCell.setCellValue(studentBean.getStuSex());
//                hssfCell.setCellValue(studentBean.getStuBirth());
//            }
//            i++;

        } catch (Exception e) {
            e.printStackTrace();
        }

        //将数据插入文件
        try {
            studentExcel = new FileOutputStream("test.xls");
            workbook.write(studentExcel);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
//        db.disconnectToMysql();
    }
}

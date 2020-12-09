package cn.neusoft.poi.tools;

import org.apache.poi.hssf.usermodel.*;
import cn.neusoft.poi.util.DBHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;

public class MySQLToExcel {

    private final DBHelper db = new DBHelper();
    private final HSSFWorkbook workBook= new HSSFWorkbook();
    private HSSFSheet sheet;
    private ResultSet resultSet;
    private int columnNum;


    /***
     * 创建sheet并且自动添加数据库列名
     * @param sheetName
     */
    public void createExcelSheet(String sheetName)
    {
        db.connectToMySql();

        try {
            String sql = "select * from Student";
            resultSet = db.search(sql);
            columnNum = db.countColumn(sql);
            sheet = workBook.createSheet(sheetName);
            String[] names = new String[columnNum];

            for (int i = 0; i<columnNum; i++) {
                names[i] = resultSet.getMetaData().getColumnName(++i);
//                System.out.println(names[i]);
            }

            HSSFRow firstRow = sheet.createRow(0);
            HSSFCell[] firstCell = new HSSFCell[columnNum];

            for (int j= 0 ;j<columnNum; j++)
            {
                firstCell[j] = firstRow.createCell(j);
                firstCell[j].setCellValue(new HSSFRichTextString(names[j]));
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /***
     * 添加数据至sheet
     * 自增变量i，用于循环增加row行数，向row中添加数据
     */
    public void addAllDataToSheet()
    {
        try
        {
            int i =1;
            while (resultSet.next())
            {
                HSSFRow excelRow = sheet.createRow(i);
                for (int j = 0 ; j < columnNum ; j++)
                {
                    HSSFCell cell = excelRow.createCell(j);
                    cell.setCellValue(resultSet.getString(++j));
                }
                i++;
            }
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
    }

    /***
     * 输出xls文件
     * @param filePath
     */
    public void writeDataToFile(String filePath)
    {
        FileOutputStream studentExcel;
        try {
            studentExcel = new FileOutputStream(filePath);
            workBook.write(studentExcel);

            if (resultSet != null){
                resultSet.close();
            }
            System.out.println("\nxls文件输出成功");
        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }
}

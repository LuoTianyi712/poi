import org.apache.poi.hssf.usermodel.*;
import util.DBHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;

public class MySQLToExcel {

    private final DBHelper db = new DBHelper();
    // 创建Excel文档
    private final HSSFWorkbook workBook= new HSSFWorkbook();
    private ResultSet resultSet;
    private int columnNum;
    private HSSFSheet sheet;

    public void createExcelSheet(String sheetName)
    {
        db.connectToMySql();

        try {
            String sql = "select * from Student";
            resultSet = db.search(sql);
            columnNum = db.countColumn(sql);
            sheet = workBook.createSheet(sheetName);
            String[] names = new String[columnNum];

            //循环将列名存放到names数组中
            for (int i = 0; i<columnNum; i++) {
                names[i] = resultSet.getMetaData().getColumnName(i+1);
            }

            //创建第一行，第一列
            HSSFRow firstRow = sheet.createRow(0);
            HSSFCell[] firstCell = new HSSFCell[columnNum];
            //循环将数据库表头输入第一列
            for (int j= 0 ;j<columnNum; j++)
            {
                firstCell[j] = firstRow.createCell(j);
                firstCell[j].setCellValue(new HSSFRichTextString(names[j]));
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void addAllDataToSheet()
    {
        try {
            int i =1;//自增变量，用于循环增加row行数，向row中添加数据
            while (resultSet.next())
            {
                HSSFRow excelRow = sheet.createRow(i);
                for (int j = 0 ; j < columnNum ; j++)
                {
                    HSSFCell cell = excelRow.createCell(j);
                    cell.setCellValue(resultSet.getString(j+1));
                }
                i++;
            }
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
    }

    /***
     * 文档输出
     * @param filePath
     */
    public void writeDataToFile(String filePath)
    {
        //输出xls
        FileOutputStream studentExcel;
        try {
            studentExcel = new FileOutputStream(filePath);
            workBook.write(studentExcel);

            if (resultSet != null){
                resultSet.close();
            }
            System.out.println("resultSet关闭，文件输出成功");
        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }
}

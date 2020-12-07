import bean.StudentBean;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import util.DBHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
/*------------------------------------------------------------------------------*/
public class Main {
    public static void main(String [] args)
    {
        //实例化属性对象
        DBHelper db = new DBHelper();//这个是一个对象，也就是只能被反复使用，而不是说有无数个可以赋值，因为我只写了一个
                                        //疑问，如果实例化对象放在while或者for循环里，是不是会反复创建对象?
//        StudentBean stuBean = new StudentBean();
        MySQLToExcel mySQLToExcel = new MySQLToExcel();
        XWPFDocument document = new XWPFDocument();


        try {
//            //注册JDBC驱动
//            Class.forName(util.DBHelper.jdbcDriver);
//
//            //打开链接
//            System.studentWord.println("连接数据库");
//            connection = DriverManager.getConnection(
//                    util.DBHelper.dbUrl,
//                    util.DBHelper.user,
//                    util.DBHelper.password);
//
//            // 执行查询
//            System.studentWord.println(" 实例化Statement对象...");
//            statement = connection.createStatement();
//            ResultSet resultSet = statement.executeQuery(sql);
            String sql;
            sql = "select * from Student";
            ResultSet resultSet = db.search(sql,null);
            ResultSetMetaData resultSetMetaData ;
/*----------------------数据库插入EXCEL定义变量↓------------------------------------*/
            int countColumnNum = resultSet.getMetaData().getColumnCount();
            int i =1;
            // 创建Excel文档
            HSSFWorkbook workBook = new HSSFWorkbook() ;
            // sheet 对应一个工作页
            HSSFSheet sheet = workBook.createSheet("student") ;
            HSSFRow firstRow = sheet.createRow(0); //下标为0的行开始
            HSSFCell[] firstCell = new HSSFCell[countColumnNum];

            String[] names = new String[countColumnNum];
            names[0] ="id";
            names[1] ="name";
            names[2] ="sex";
            names[3] ="birth";

            for(int j= 0 ;j<countColumnNum; j++)
            {
                firstCell[j] = firstRow.createCell(j);
                firstCell[j].setCellValue(new HSSFRichTextString(names[j]));
            }
/*----------------------数据库插入EXCEL定义变量↑------------------------------------*/

            //列宽自动分割
            XWPFTable comTable = document.createTable();
            CTTblWidth comTableWidth = comTable.getCTTbl().addNewTblPr().addNewTblW();
            comTableWidth.setType(STTblWidth.DXA);
            comTableWidth.setW(BigInteger.valueOf(9072));

            XWPFTableCell [] firstWordCell = new XWPFTableCell[countColumnNum];
            XWPFTableRow firstWordRow;
            firstWordRow = comTable.getRow(0);
            firstWordRow.getCell(0).setText(names[0]);
            for (int j = 1 ; j < countColumnNum; j++)
            {

//                firstWordCell[j] = firstWordRow.createCell();
//                firstWordCell[j].setText(names[j]);

                firstWordRow.addNewTableCell().setText(names[j]);



            }

            while (resultSet.next()) {
                StudentBean stuBean = new StudentBean();//会反复赋值吗？会一行一行添加数据吗？
                stuBean.setStuId(resultSet.getInt("stu_id"));
                stuBean.setStuName(resultSet.getString("stu_name"));
                stuBean.setStuSex(resultSet.getString("stu_sex"));
                stuBean.setStuBirth(resultSet.getString("stu_birth"));

                System.out.println("学生ID" + "\t\t" + stuBean.getStuId());
                System.out.println("学生姓名" + "\t\t" + stuBean.getStuName());
                System.out.println("学生性别" + "\t\t" + stuBean.getStuSex());
                System.out.println("学生生日" + "\t\t" + stuBean.getStuBirth());
                System.out.println();
/*----------------------数据库插入EXCEL循环打印↓------------------------------------*/

//                mySQLToExcel.insertDataToExcel(stuBean);//插入数据到excel
                HSSFRow excelRow = sheet.createRow(i);
                for (int j = 0 ; j < countColumnNum ; j++)
                {
                    HSSFCell cell = excelRow.createCell(j);
                    cell.setCellValue(
                            new HSSFRichTextString(resultSet.getString(j+1)));
                }
                i++;


            }
            FileOutputStream studentExcel;
            studentExcel = new FileOutputStream("test.xls");
            workBook.write(studentExcel);
/*----------------------数据库插入EXCEL循环打印↑------------------------------------*/

            FileOutputStream studentWord;
            studentWord = new FileOutputStream("poi_word.docx");
            document.write(studentWord);

            resultSet.close();

//            db.countColumn();

        } catch (Exception se) {
            //处理JDBC错误
            se.printStackTrace();
        }
        //处理Class.forName
//        finally {
//            //释放资源,关闭端口
//            try {
//                if (db.prepareStatement != null)
//                    db.prepareStatement.close();
//            } catch (SQLException se2) {
//                //不执行
//            }
//            try {
//                if (db.connection != null)
//                    db.connection.close();
//            } catch (SQLException se) {
//                se.printStackTrace();
//            }
//        }
        System.out.println("END");
    }
}

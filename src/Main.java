import bean.StudentBean;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import util.DBHelper;

import java.io.FileOutputStream;
import java.math.BigInteger;
import java.sql.ResultSet;

/*------------------------------------------------------------------------------*/
public class Main {
    public static void main(String [] args)
    {
        //实例化属性对象
        DBHelper db = new DBHelper();
        StudentBean stuBean = new StudentBean();
        MySQLToExcel mySQLToExcel = new MySQLToExcel();
        MySQLToWord mySQLToWord = new MySQLToWord();

        try {
            db.connectToMySql();

            String sql;
            sql = "select * from Student";

            ResultSet resultSet = db.search(sql);
//            int countColumnNum = resultSet.getMetaData().getColumnCount();
//            int countColumnNum;
//            countColumnNum = db.countColumn(sql);
//            String[] names = new String[countColumnNum];
//            //此处应该用从数据库导入列名的方式来进行修改，而不是直接定义数组
////            names[0] ="id";
////            names[1] ="name";
////            names[2] ="sex";
////            names[3] ="birth";
//
//            for (int i = 0; i<countColumnNum; i++)
//            {
//                names[i] = resultSet.getMetaData().getColumnName(i+1);
//                System.out.println(resultSet.getMetaData().getColumnName(i+1));
//            }
//            XWPFDocument document = new XWPFDocument();
//            //列宽自动分割
//            XWPFTable comTable = document.createTable();
//
//            CTTblWidth comTableWidth = comTable.getCTTbl().addNewTblPr().addNewTblW();
//            comTableWidth.setType(STTblWidth.DXA);
//            comTableWidth.setW(BigInteger.valueOf(9072));
//
////            XWPFTableCell [] firstWordCell = new XWPFTableCell[countColumnNum];
//
//            XWPFTableRow wordRow = comTable.getRow(0);
//            wordRow.getCell(0).setText(names[0]);//直接将0行0列放置于word表格中
//            //再通过循环方式将数据放入剩余表格中
//            for (int j = 1 ; j < countColumnNum; j++)
//            {
////                firstWordCell[j] = wordRow.createCell();
////                firstWordCell[j].setText(names[j]);
//                wordRow.addNewTableCell().setText(names[j]);
//            }

            while (resultSet.next())
            {
//                StudentBean stuBean = new StudentBean();
                // 会反复赋值吗？会一行一行添加数据吗？
                //不需要循环创建对象，因为会占堆内存，只需要创建一个对象，然后循环赋值给他
                //赋值完成后插入一行，并创建新行

                stuBean.setStuId(resultSet.getInt("stu_id"));
                stuBean.setStuName(resultSet.getString("stu_name"));
                stuBean.setStuSex(resultSet.getString("stu_sex"));
                stuBean.setStuBirth(resultSet.getString("stu_birth"));

                System.out.println("学生ID" + "\t\t" + stuBean.getStuId());
                System.out.println("学生姓名" + "\t\t" + stuBean.getStuName());
                System.out.println("学生性别" + "\t\t" + stuBean.getStuSex());
                System.out.println("学生生日" + "\t\t" + stuBean.getStuBirth());
                System.out.println();
///*↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓数据库插入word循环打印 while创建新行↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓*/
////                XWPFTableRow wordRow = comTable.createRow();
//                wordRow = comTable.createRow();
//                for (int j = 0; j < countColumnNum; j++)
//                {
//                    XWPFTableCell tableCell = wordRow.getCell(j);//空，无法找到行，需要创建行
//                    tableCell.setText(resultSet.getString(j+1));
//                }
///*↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑数据库插入word循环打印 while创建新行↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑*/
            }

            mySQLToExcel.createExcelSheet("student");
            mySQLToExcel.addAllDataToSheet();
            mySQLToExcel.writeDataToFile("poi_student.xls");

            mySQLToWord.createWordForm();
            mySQLToWord.addDateToForm();
            mySQLToWord.writeFormToDocx();

//            //输出word文档
//            FileOutputStream studentWord;
//            studentWord = new FileOutputStream("poi_student.docx");
//            document.write(studentWord);

            resultSet.close();
            db.disconnectToMysql();

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

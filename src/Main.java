import bean.StudentBean;

import util.DBHelper;

import java.sql.ResultSet;

/*------------------------------------------------------------------------------*/
public class Main {
    public static void main(String [] args)
    {
        DBHelper db = new DBHelper();
        StudentBean stuBean = new StudentBean();
        MySQLToExcel mySQLToExcel = new MySQLToExcel();
        MySQLToWord mySQLToWord = new MySQLToWord();

        try {
            db.connectToMySql();
            String sql = "select * from Student";
            ResultSet resultSet = db.search(sql);

            while (resultSet.next())
            {
                stuBean.setStuId(resultSet.getInt("stu_id"));
                stuBean.setStuName(resultSet.getString("stu_name"));
                stuBean.setStuSex(resultSet.getString("stu_sex"));
                stuBean.setStuBirth(resultSet.getString("stu_birth"));

                System.out.println("学生ID" + "\t\t" + stuBean.getStuId());
                System.out.println("学生姓名" + "\t\t" + stuBean.getStuName());
                System.out.println("学生性别" + "\t\t" + stuBean.getStuSex());
                System.out.println("学生生日" + "\t\t" + stuBean.getStuBirth());
                System.out.println();
            }

            mySQLToExcel.createExcelSheet("student");
            mySQLToExcel.addAllDataToSheet();
            mySQLToExcel.writeDataToFile("poi_student.xls");

            mySQLToWord.createWordForm();
            mySQLToWord.addAllDateToWordForm();
            mySQLToWord.writeFormToDocx("poi_student.docx");

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

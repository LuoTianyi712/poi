import bean.StudentBean;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import util.DBHelper;

import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

public class Main {
    public static void main(String [] args)
    {
        //实例化属性对象
        DBHelper db = new DBHelper();//这个是一个对象，也就是只能被反复使用，而不是说有无数个可以赋值，因为我只写了一个
                                        //疑问，如果实例化对象放在while或者for循环里，是不是会反复创建对象?
//        StudentBean stuBean = new StudentBean();
        MySQLToExcel mySQLToExcel = new MySQLToExcel();

        try {
//            //注册JDBC驱动
//            Class.forName(util.DBHelper.jdbcDriver);
//
//            //打开链接
//            System.out.println("连接数据库");
//            connection = DriverManager.getConnection(
//                    util.DBHelper.dbUrl,
//                    util.DBHelper.user,
//                    util.DBHelper.password);
//
//            // 执行查询
//            System.out.println(" 实例化Statement对象...");
//            statement = connection.createStatement();
//            ResultSet resultSet = statement.executeQuery(sql);

            String sql;
            sql = "select * from Student";
            ResultSet resultSet = db.search(sql,null);
            ResultSetMetaData resultSetMetaData ;


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

                mySQLToExcel.insertDataToExcel(stuBean);//插入数据到excel

            }
            resultSet.close();

            db.countColumn();

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

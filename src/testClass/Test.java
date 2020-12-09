package testClass;

import cn.neusoft.poi.bean.StudentBean;
import cn.neusoft.poi.util.DBHelper;

public class Test {
//    public static void main(String [] args){
//        for (int i = 1 ; i<10 ; i++)
//        {
//            System.out.println("ABCDE");
//            for (int j = 1 ; j < 5 ; j++)
//            {
//                System.out.println("FGHIJ");
//            }
//        }
//    }
    public static void main(String[] args) {
        StudentBean studentBean = new StudentBean();
        DBHelper dbHelper = new DBHelper();
        System.out.println(studentBean);//堆内存地址

        int a = 1;
        dbHelper.connectToMySql();
        System.out.println(System.identityHashCode(a));//唯一标识符
        System.out.println(dbHelper.countRow("select * from Student"));
        dbHelper.disconnectToMysql();
    }
}

package util;

import java.sql.*;

public class DBHelper {
    // MySQL 8.0 以下版本 - JDBC 驱动名及数据库 URL
    static final String jdbcDriver = "com.mysql.jdbc.Driver";
    static final String dbUrl = "jdbc:mysql://localhost:3306/TestDB?useSSL=false";

    // MySQL 8.0 以上版本 - JDBC 驱动名及数据库 URL
    //static final String JDBC_DRIVER = "com.mysql.cj.jdbc.Driver";
    //static final String DB_URL = "jdbc:mysql://localhost:3306/TestDB?useSSL=false&allowPublicKeyRetrieval=true&serverTimezone=UTC";


    // 数据库的用户名与密码，需要根据自己的设置
    static final String user = "root";
    static final String password = "root";

    public Connection connection = null;
    public PreparedStatement prepareStatement = null;
    public ResultSet resultSet = null;

    public void connectToMySql() {
        try {
            Class.forName(jdbcDriver);
            connection = DriverManager.getConnection(dbUrl, user, password);
        } catch (ClassNotFoundException e) {
            // TODO Auto-generated catch block
            System.err.println("装载 JDBC/ODBC 驱动程序失败。" );
            e.printStackTrace();
        } catch (SQLException e) {
            // TODO Auto-generated catch block
            System.err.println("无法连接数据库" );
            e.printStackTrace();
        }
    }

    //查询
    public ResultSet search(String sql, String [] str){
        connectToMySql();
        try {
            prepareStatement = connection.prepareStatement(sql);
            if(str != null){
                for(int i=0;i<str.length-1;i++){
                    prepareStatement.setString(i+1, str[i]);
                }
                prepareStatement.setInt(str.length, Integer.parseInt(str[str.length-1]));
            }
            resultSet = prepareStatement.executeQuery();
        } catch (SQLException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
//        disconnectToMysql();
        return resultSet;
    }

    //计算行数
    public int countRow(){

//        connectToMySql();
        int rowCount = 0;
        try {
            prepareStatement = connection.prepareStatement("select * from student");
            resultSet = prepareStatement.executeQuery();
            resultSet.last();//检索当前行号。第一行是1，第二行是2，依此类推。
            rowCount = resultSet.getRow();

        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
//        Integer.parseInt(String.valueOf(resultSet));
//        System.out.println(rowCount);
//        disconnectToMysql();
        return rowCount;
    }

    //计算列数
    public int countColumn(){
//        connectToMySql();
        int columnCount = 0;
        try {
//            prepareStatement = connection.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_UPDATABLE);
            resultSet = prepareStatement.executeQuery("select * from student");
            ResultSetMetaData rsmd = resultSet.getMetaData() ;
            columnCount = rsmd.getColumnCount();
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
//        System.out.println(columnCount);
//        disconnectToMysql();
        return columnCount;
    }

    //关闭连接
    public void disconnectToMysql()
    {
        try
        {
            connection.close();
            prepareStatement.close();
            resultSet.close();
        } catch (SQLException e)
        {
            e.printStackTrace();
        }
    }
}

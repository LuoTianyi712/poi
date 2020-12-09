package util;

import java.sql.*;

public class DBHelper {
    // MySQL 8.0 以下版本 - JDBC 驱动名及数据库 URL
    private static final String jdbcDriver = "com.mysql.jdbc.Driver";
    private static final String dbUrl = "jdbc:mysql://localhost:3306/TestDB?useSSL=false";

    // MySQL 8.0 以上版本 - JDBC 驱动名及数据库 URL
    //static final String JDBC_DRIVER = "com.mysql.cj.jdbc.Driver";
    //static final String DB_URL = "jdbc:mysql://localhost:3306/TestDB?useSSL=false&allowPublicKeyRetrieval=true&serverTimezone=UTC";

    // 数据库的用户名与密码，需要根据自己的设置
    private static final String user = "root";
    private static final String password = "root";

    private Connection connection = null;
    ResultSet resultSet = null;

    public void connectToMySql() {
        try {
            Class.forName(jdbcDriver);
            connection = DriverManager.getConnection(dbUrl, user, password);
        } catch (ClassNotFoundException e) {
            System.err.println("无法找到JDBC驱动");
            e.printStackTrace();
        } catch (SQLException e) {
            System.err.println("无法连接数据库");
            e.printStackTrace();
        }
    }

    //查询
//    public ResultSet search(String sql, String [] str){
    public ResultSet search(String sql){
        PreparedStatement prepareStatement;
        try {
            prepareStatement = connection.prepareStatement(sql);
//            if(str != null){
//                for(int i=0;i<str.length-1;i++){
//                    prepareStatement.setString(i+1, str[i]);
//                }
//                prepareStatement.setInt(str.length, Integer.parseInt(str[str.length-1]));
//            }
            resultSet = prepareStatement.executeQuery();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return resultSet;
    }

    //计算行数
    public int countRow(String sql){

        PreparedStatement prepareStatement;
        ResultSet resultSet;
        int rowCount = 0;
        try {
            prepareStatement = connection.prepareStatement(sql);
            resultSet = prepareStatement.executeQuery();
            resultSet.last();//检索当前行号。第一行是1，第二行是2，依此类推。
            // 如果不写直接getRow则只会获取到0行数据
            rowCount = resultSet.getRow();

            //关闭资源
            if(prepareStatement != null)
            {
                prepareStatement.close();
            }
            if (resultSet != null)
            {
                resultSet.close();
            }

        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
        return rowCount;
    }

    //计算列数
    public int countColumn(String sql){

        //创建变量
        PreparedStatement prepareStatement;
        ResultSet resultSet = null;
        int columnCount = 0;

        try {
            prepareStatement = connection.prepareStatement(sql);
            resultSet = prepareStatement.executeQuery();
            columnCount = resultSet.getMetaData().getColumnCount();
            //关闭资源
            if(prepareStatement != null)
            {
                prepareStatement.close();
            }
            if (resultSet != null)
            {
                resultSet.close();
            }
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }

        return columnCount;
    }

    //关闭连接
    public void disconnectToMysql()
    {
        try
        {
            if (connection != null) {
                connection.close();
            }
            if (resultSet != null){
                resultSet.close();
            }

        } catch (SQLException e)
        {
            e.printStackTrace();
        }
        System.out.println("连接关闭");
    }
}

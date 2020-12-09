package cn.neusoft.poi.util;

import java.sql.*;

public class DBHelper {
    // MySQL 8.0 以下版本 - JDBC 驱动名及数据库 URL
    private static final String jdbcDriver = "com.mysql.jdbc.Driver";
    private static final String mysqlUrl = "jdbc:mysql://localhost:3306/TestDB?useSSL=false";

    // MySQL 8.0 以上版本 - JDBC 驱动名及数据库 URL
    //static final String JDBC_DRIVER = "com.mysql.cj.jdbc.Driver";
    //static final String DB_URL = "jdbc:mysql://localhost:3306/TestDB?useSSL=false&allowPublicKeyRetrieval=true&serverTimezone=UTC";

    // MySQL用户名与密码
    private static final String user = "root";
    private static final String password = "root";

    private Connection connection = null;
    private ResultSet resultSet = null;
    private PreparedStatement prepareStatement = null;

    /***
     * 连接数据库
     */
    public void connectToMySql() {
        try {
            Class.forName(jdbcDriver);
            connection = DriverManager.getConnection(mysqlUrl, user, password);
        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }
    }
    /***
     * MySQL查询表数据
     * @param sql
     * @return
     */
    public ResultSet search(String sql){
        try {
            prepareStatement = connection.prepareStatement(sql);
            resultSet = prepareStatement.executeQuery();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return resultSet;
    }

    /***
     * 测试用例，重载search
     * @param sql
     * @param str
     * @return
     */
    public ResultSet search(String sql, String[] str){
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
            e.printStackTrace();
        }
        return resultSet;
    }

    /***
     * 计算MySQL表行数
     * last()检索当前行号。第一行是1，第二行是2，依此类推。
     * 如果不写resultSet.last()，直接getRow，则只会获取到0行数据
     * @param sql
     * @return
     */
    public int countRow(String sql){
        int rowCount = 0;
        try {
            prepareStatement = connection.prepareStatement(sql);
            resultSet = prepareStatement.executeQuery();
            resultSet.last();

            rowCount = resultSet.getRow();

        } catch (SQLException se) {
            se.printStackTrace();
        }
        return rowCount;
    }

    /***
     * 计算MySQL表列数
     * @param sql
     * @return
     */
    public int countColumn(String sql){

        int columnCount = 0;

        try {
            prepareStatement = connection.prepareStatement(sql);
            resultSet = prepareStatement.executeQuery();
            columnCount = resultSet.getMetaData().getColumnCount();
        } catch (SQLException se) {
            se.printStackTrace();
        }

        return columnCount;
    }

    /***
     * 关闭MySQL连接，
     * 释放资源
     */
    public void disconnectToMysql()
    {
        try
        {
            System.out.println("\n关闭所有连接...");
            if (connection != null) {
                connection.close();
            }
            if (prepareStatement != null) {
                prepareStatement.close();
            }
            if (resultSet != null){
                resultSet.close();
            }
        } catch (SQLException se)
        {
            se.printStackTrace();
        }
        System.out.println("所有连接已关闭");
    }
}

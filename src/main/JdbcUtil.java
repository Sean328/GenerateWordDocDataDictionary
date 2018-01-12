package main;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

/**
 * jdbc工具类
 * @author APPle
 *
 */
public class JdbcUtil {
    private static String url = "JDBC:oracle:thin:@120.24.101.***:1522:orcl";
    private static String user = "idsp_ifp";
    private static String password = "123456";
    private static String driverClass = "oracle.jdbc.driver.OracleDriver";

    /**
     * 抽取获取连接对象的方法
     */
    public Connection getConnection(){
        try {
            Class.forName(driverClass);
            Connection conn = DriverManager.getConnection(url, user, password);
            return conn;
        } catch (SQLException | ClassNotFoundException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }


    /**
     * 释放资源的方法
     */
    public static void close(Connection conn,Statement stmt){
        if(stmt!=null){
            try {
                stmt.close();
            } catch (SQLException e) {
                e.printStackTrace();
                throw new RuntimeException(e);
            }
        }
        if(conn!=null){
            try {
                conn.close();
            } catch (SQLException e) {
                e.printStackTrace();
                throw new RuntimeException(e);
            }
        }
    }

    public static void close(Connection conn,Statement stmt,ResultSet rs){
        if(rs!=null)
            try {
                rs.close();
            } catch (SQLException e1) {
                e1.printStackTrace();
                throw new RuntimeException(e1);
            }
        if(stmt!=null){
            try {
                stmt.close();
            } catch (SQLException e) {
                e.printStackTrace();
                throw new RuntimeException(e);
            }
        }
        if(conn!=null){
            try {
                conn.close();
            } catch (SQLException e) {
                e.printStackTrace();
                throw new RuntimeException(e);
            }
        }
    }
}
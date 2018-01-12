package main;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Test {
    
    private static JdbcUtil jdbc=new JdbcUtil();
    private static Connection conn=jdbc.getConnection();
    //要生成表格的名字，可以模糊查询
    private static String name="";
    
   public static void main(String[] args) {
       WordBean wordBean = new WordBean();    
       wordBean.setVisible(false); // 是否前台打开word 程序，或者后台运行    
       wordBean.createNewDocument();// 创建一个新文档    
       wordBean.setLocation();    
       List<Map<String,String>> resultList=new ArrayList<Map<String,String>>();
       resultList = selectMnue(name);
       for(int i=0;i<resultList.size();i++){
           List<List<String>> detailList=selectDetail(resultList.get(i).get("TABLE_NAME"));
           System.out.println(i+"    "+detailList.size()+"   "+resultList.get(i).get("TABLE_NAME"));
           wordBean.insertTable((i+1)+"."+resultList.get(i).get("TABLE_NAME")+"("+resultList.get(i).get("COMMENTS")+")", detailList.size(), 5, detailList);
//           if(i==10){
//             break;  
//           }
       }
       System.out.println("生成完成");
       wordBean.saveFileAs("d:\\字典树.doc");    
       wordBean.closeDocument();    
       wordBean.closeWord(); 
   }
   
   public static List<Map<String,String>> selectMnue(String name){
       List<Map<String,String>> result=new ArrayList<Map<String,String>>();
       PreparedStatement stmt=null;
       ResultSet res = null; 
       String sql = "select * from user_tab_comments WHERE TABLE_NAME LIKE concat(concat('%',?),'%')";
       try {
        stmt = conn.prepareStatement(sql);
        stmt.setString(1, name);
        res = stmt.executeQuery();
        while(res.next()){
            Map<String,String> map = new HashMap<String,String>();
            map.put("TABLE_NAME", res.getString("TABLE_NAME"));
            map.put("TABLE_TYPE", res.getString("TABLE_TYPE"));
            map.put("COMMENTS", res.getString("COMMENTS"));
            result.add(map);
        }
    } catch (SQLException e) {
        e.printStackTrace();
    }
       return result;
   }
   
   public static List<List<String>> selectDetail(String tableName){
       List<List<String>> result=new ArrayList<List<String>>();
       PreparedStatement stmt=null;
       ResultSet res = null; 
       String sql = "SELECT A.COLUMN_NAME AS name,A.DATA_TYPE AS typ,A.DATA_DEFAULT AS def,A.NULLABLE AS nul,B.comments AS com FROM sys.user_tab_columns A, sys.user_col_comments B WHERE A.table_name = B.table_name AND A.COLUMN_NAME = B.COLUMN_NAME AND A.TABLE_NAME = ? ORDER BY A.TABLE_NAME";
       try {
        stmt = conn.prepareStatement(sql);
        stmt.setString(1, tableName);
        res = stmt.executeQuery();
        List<String> titleList = new ArrayList<String>();
        titleList.add("字段名");
        titleList.add("字段类型");
        titleList.add("默认值");
        titleList.add("能否为空");
        titleList.add("备注");
        result.add(titleList);
        while(res.next()){
            List<String> map = new ArrayList<String>();
            map.add(res.getString("name"));
            map.add(res.getString("typ"));
            map.add(res.getString("def"));
            map.add(res.getString("nul"));
            if(null==res.getString("com")){
                map.add("无");
            }else{
                map.add(res.getString("com"));
            }
            result.add(map);
        }
    } catch (SQLException e) {
        e.printStackTrace();
    }
       return result;
   }
}

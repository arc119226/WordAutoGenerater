package com.arcare.document.docx.wrap;

import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class ResultSetToJson {
	/**
	 * 輸入JDBC ResultSet 回傳JSON
	 * @param rs
	 * @return
	 */
     public static final JsonObject resultSetToJsonObject(ResultSet rs) {
        JsonObject element = null;
        JsonArray ja = new JsonArray();
        JsonObject jo = new JsonObject();
        ResultSetMetaData resultSetMetaData = null;
        String columnName, columnValue = null;
        boolean hasData=false;
        try {
            resultSetMetaData = rs.getMetaData();
            while (rs.next()) {
            	hasData=true;
                element = new JsonObject();
                for (int i = 0; i < resultSetMetaData.getColumnCount(); i++) {
                    columnName = resultSetMetaData.getColumnName(i + 1);
                    columnValue = rs.getString(columnName);
                    if(columnValue==null) {
                    	columnValue="";
                    }
                    element.addProperty(columnName, columnValue);
                }
                ja.add(element);
            }
            jo.add("result", ja);
        } catch (SQLException e) {
            Log.error(e);
        }
        if(hasData){
        	return jo;
        }else{
        	return null;
        }
    }
}

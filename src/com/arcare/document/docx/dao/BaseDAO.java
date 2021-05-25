package com.arcare.document.docx.dao;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Optional;

import javax.sql.DataSource;

import com.arcare.document.docx.wrap.Log;
import com.arcare.document.docx.wrap.ResultSetToJson;
import com.google.gson.JsonObject;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public abstract class BaseDAO {

	protected DataSource dataSource;
	
	public BaseDAO(DataSource dataSource) {
		this.dataSource=dataSource;
	}
	
	/**
	 * 從檔案載入SQL
	 * @param sqlfile
	 * @return
	 */
	protected String getSql(String sqlfile){
		String sql="";
		try {
			sql= new String(Files.readAllBytes(Paths.get(sqlfile)),StandardCharsets.UTF_8);
		} catch (IOException e) {
			Log.error(e);
		}
		return sql;
	}
	
	/**
	 * 
	 * 依據sql執行query
	 * @return
	 */
	protected Optional<JsonObject> query(String sql){
		Connection con=null;
		ResultSet rs=null;
		try {
			con=this.dataSource.getConnection();
			Log.log(sql);
			PreparedStatement pst=con.prepareStatement(sql);
			rs=pst.executeQuery();
			JsonObject obj=ResultSetToJson.resultSetToJsonObject(rs);
			if(obj==null) {
				return Optional.empty();
			}else {
				return Optional.of(obj);
			}
		} catch (Exception e) {
			Log.error(e);
			return Optional.empty();
		}finally{
			try {
				rs.close();
				con.close();
			} catch (SQLException e) {
				Log.error(e);
			}
		}
	}
}

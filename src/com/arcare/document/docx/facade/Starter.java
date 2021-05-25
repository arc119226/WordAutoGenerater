package com.arcare.document.docx.facade;

import java.beans.PropertyVetoException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Map;

import javax.sql.DataSource;

import org.eclipse.jetty.server.Server;
import org.eclipse.jetty.servlet.ServletContextHandler;
import org.eclipse.jetty.servlet.ServletHolder;

import com.arcare.document.docx.controller.ReportApiServlet;
import com.arcare.document.docx.dao.CommonDAOV1;
import com.arcare.document.docx.service.WordAutoGeneraterService;
import com.arcare.document.docx.service.WordAutoGeneraterServiceV1;
import com.arcare.document.docx.wrap.CommonUtil;
import com.arcare.document.docx.wrap.Log;
import com.mchange.v2.c3p0.ComboPooledDataSource;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
@Deprecated
public class Starter {
	/**
	 * init datasource
	 * @param config
	 * @return
	 * @throws PropertyVetoException
	 */
	private static DataSource prepareDataSource(Map<String,String> config) throws PropertyVetoException {
		//close c3p0 error log
//		Properties p = new Properties(System.getProperties());
//		p.put("com.mchange.v2.log.MLog", "com.mchange.v2.log.FallbackMLog");
//		p.put("com.mchange.v2.log.FallbackMLog.DEFAULT_CUTOFF_LEVEL", "OFF");
//		System.setProperties(p);
		
		ComboPooledDataSource dataSource= new ComboPooledDataSource();
		dataSource.setDriverClass(config.get("agent.db.drive"));
		dataSource.setJdbcUrl(config.get("agent.db.jdbc"));
		dataSource.setUser(config.get("agent.db.user"));
		dataSource.setPassword(config.get("agent.db.password"));
		dataSource.setMinPoolSize(10);
		dataSource.setAcquireIncrement(10);
		dataSource.setMaxPoolSize(10);
		dataSource.setMaxStatements(10);
		dataSource.setAutoCommitOnClose(false);
		dataSource.setUnreturnedConnectionTimeout(100);
		Connection con=null;
		ResultSet rs=null;
		try {
			con=dataSource.getConnection();
			PreparedStatement pst=con.prepareStatement("select 1");
			rs=pst.executeQuery();
			while(rs.next()) {
				Log.log("test connection :"+("1".equals(rs.getString(1))?"OK":"ERROR"));
			}
		} catch (SQLException e) {
			Log.error(e);
		}finally {
			try {
				if(con!=null) {
					con.close();
				}
				if(rs!=null) {
					rs.close();
				}
			} catch (SQLException e) {
				Log.error(e);
			}

		}
		return dataSource;
	}

	/**
	 * test case
	 * 
	 * @param args
	 * @throws PropertyVetoException 
	 * @throws UnsupportedEncodingException
	 * @throws IOException
	 */
	public static void main(String[] args) throws PropertyVetoException {

		Map<String,String> config=new HashMap<>();
		config=CommonUtil.initDataBind("./config/config.properties",false);

		if(config.isEmpty()) {
			Log.log("./config/config.properties can't found.");
			return;
		}

		DataSource datasource=prepareDataSource(config);
		CommonDAOV1 commonDao=new CommonDAOV1(datasource);

		WordAutoGeneraterService wordAutoGeneraterServiceV1=new WordAutoGeneraterServiceV1(commonDao,config);
		
		ReportApiServlet reportApiServletV1=new ReportApiServlet(wordAutoGeneraterServiceV1);

        //curl -i -d '{"ApplyNumber":"EO-2018-0001", "ReporCategory":"EO101","Revision":"A"}' -H "Content-Type: application/json" -X POST  http://127.0.0.1:18080/v1/Reporter_API

		ServletContextHandler contextHandler = new ServletContextHandler(); 
        contextHandler.setContextPath("/"); 
        contextHandler.addServlet(new ServletHolder(reportApiServletV1), "/v1/Reporter_API"); 

        try {
            Server server = new Server(Integer.valueOf(config.get("server.port"))); 
            server.setHandler(contextHandler); 
			server.start();
		} catch (Exception e) {
			Log.error(e);
		} 
	}
	
	public static void stop(String[] args) {
		Log.log("shotDown...");
		System.exit(0);
	}
}

package com.arcare.document.docx.wrap;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Comparator;
import java.util.Map;
import java.util.Properties;
import java.util.TreeMap;
import java.util.concurrent.ConcurrentHashMap;

import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.util.Zip4jConstants;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class CommonUtil {
	/**
	 * init initDataBind
	 * @param arg0
	 */
	public static Map<String,String> initDataBind(String configPath,Boolean exchange){
		FileInputStream fis=null;
		try {
			Map<String,String> map=new TreeMap<String, String>();
			Properties prop = new Properties();
			fis=new FileInputStream(configPath);
			prop.load(new InputStreamReader(fis, Charset.forName("UTF-8")));
			prop.stringPropertyNames()
				.forEach( key-> {
					if(exchange) {
						map.put(prop.getProperty(key).toUpperCase(),key);
					}else {
						map.put(key, prop.getProperty(key));
					}
				});
			return map;
		} catch (IOException io) {
			Log.error(io);
			return new ConcurrentHashMap<String, String>();
		}finally{
			if(fis!=null){
				try {
					fis.close();
				} catch (IOException e) {
					Log.error(e);
				}
			}
		}
	}
	/**
	 * unzip file
	 * @param source
	 * @param outputDir
	 * @return
	 */
	@Deprecated
	public static boolean unzip(String source,String outputDir){
		try{
			Log.log("unzip "+source+" -> "+outputDir);
			ZipFile zipFile = new ZipFile(source);
	        zipFile.extractAll(outputDir);
			return true;
		}catch(Exception e){
			Log.error(e);
		}
		return false;
	}
	/**
	 * zip file
	 * @param source
	 * @param targetFilePath
	 * @return
	 */
	@Deprecated
	public static boolean zip(String source,String targetFilePath){
		try{
			Log.log("zip "+source+" -> "+targetFilePath);
			ZipFile zipFile = new ZipFile(targetFilePath);
			ZipParameters parameters = new ZipParameters();
			parameters.setCompressionMethod(Zip4jConstants.COMP_DEFLATE);
			parameters.setCompressionLevel(Zip4jConstants.DEFLATE_LEVEL_NORMAL);		
			for(File f:new File(source).listFiles()){
				if(f.isDirectory()){
					zipFile.addFolder(f, parameters);
				}else{
					zipFile.addFile(f, parameters);
				}
			}
			return true;
		}catch(Exception e){
			Log.error(e);
		}
		return false;
	}
	/**
	 * remove old temp folder
	 * @param rootOutput
	 */
	public static void cleanOldTemp(String rootOutput){
		if(new File(rootOutput).exists()){
			Log.log("cleanOldTemp "+rootOutput);
			Path pathToBeDeleted = Paths.get(rootOutput);
			try {
				Files.walk(pathToBeDeleted)
				  .sorted(Comparator.reverseOrder())
				  .map(Path::toFile)
				  .forEach(File::delete);
			} catch (IOException e) {
				Log.error(e);
			}
		}
	}
}

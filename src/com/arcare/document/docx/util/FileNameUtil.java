package com.arcare.document.docx.util;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class FileNameUtil {
	/**
	 * 
	 * @param config
	 * @return <root_dor>/<currentDate>_result.docx
	 */
	public static String generateResultFileName(Map<String,String> config) {
		File outputDir = new File(config.get("agent.path.rootdir"));
		outputDir.mkdirs();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		String currentDate = sdf.format(new Date());
		String section=config.get("agent.path.rootdir") + File.separator + currentDate + "_result.docx";
		return section;
	}
	/**
	 * 
	 * @param config
	 * @return <root_dor>/<currentDate>_section.docx
	 */
	public static String generateSectionFileName(Map<String,String> config) {
		File outputDir = new File(config.get("agent.path.rootdir"));
		outputDir.mkdirs();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		String currentDate = sdf.format(new Date());
		String section=config.get("agent.path.rootdir") + File.separator + currentDate + "_section.docx";
		return section;
	}
	/**
	 * 
	 * @param config
	 * @return <root_dor>/<currentDate>_concat.docx
	 */
	public static String generateConcatFileName(Map<String,String> config) {
		File outputDir = new File(config.get("agent.path.rootdir"));
		outputDir.mkdirs();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		String currentDate = sdf.format(new Date());
		String section=config.get("agent.path.rootdir") + File.separator + currentDate + "_concat.docx";
		return section;
	}
}

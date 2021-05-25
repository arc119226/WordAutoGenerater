package com.arcare.document.docx.wrap;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class Log {
	public static void log(String msg) {
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String time = df.format(new Date());
		System.out.println(String.format("[INFO][%s] %s", time,msg));
		try {
			File f = new File("info.txt");
			if (!f.exists()) {
				f.createNewFile();
			}
			Files.write(Paths.get("info.txt"),
					Arrays.asList(String.format("[INFO][%s] %s", time,msg)),
					Charset.forName("UTF-8"),
					StandardOpenOption.APPEND);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public static String error(Exception e) {
		StringWriter sw = new StringWriter();
		PrintWriter pw = new PrintWriter(sw);
		e.printStackTrace(pw);
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String time = df.format(new Date());
		System.out.println(String.format("[ERROR][%s] %s", time,sw));
		try {
			File f = new File("error.txt");
			if (!f.exists()) {
				f.createNewFile();
			}
			Files.write(Paths.get("error.txt"),
					Arrays.asList(String.format("[ERROR][%s] %s", time,sw)),
					Charset.forName("UTF-8"),
					StandardOpenOption.APPEND);
			return sw.toString();
		} catch (IOException e1) {
			e.printStackTrace();
		}
		return e.getMessage();
	}
}

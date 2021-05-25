package com.arcare.document.docx.wrap;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class ExcelUtil {
	/**
	 * 讀取excel
	 * @param excelFilePath
	 * @return
	 */
	public static Workbook readExcel(String excelFilePath) {
		try {
			Log.log("readExcel:"+excelFilePath);
			Path path = Paths.get(excelFilePath);
			byte[] byteData = Files.readAllBytes(path);
			Workbook excel = new XSSFWorkbook(new ByteArrayInputStream(byteData));
			return excel;
		} catch (IOException e) {
			Log.error(e);
		}
		return null;
	}
	/**
	 * excel 存檔
	 * @param workbook
	 * @param excelFilePath
	 */
	public static void savExcel(Workbook workbook,String excelFilePath) {
		 FileOutputStream outputStream=null;
		try {
            //save file
			outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
			outputStream.flush();
			workbook.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }finally {
        	if(outputStream!=null) {
        		 try {
        			 outputStream.close();
				} catch (IOException e) {
					Log.error(e);
				}
        	}
        }
	}
	
	/**
	 * 計算有幾個sheet
	 * @param workbook
	 * @return
	 */
	public static int getSheetCout(Workbook workbook) {
		return workbook.getNumberOfSheets();
	}
	/**
	 * 切換sheet 到 index (index從0開始)
	 * @param workbook
	 * @param index
	 */
	public static void switchActiveSheet(Workbook workbook,int index) {
		workbook.setActiveSheet(index);
	}
	/**
	 * 切換sheet 依據sheet名稱
	 * @param workbook
	 * @param index
	 */
	public static void switchActiveSheetByName(Workbook workbook,String sheetName) {
		Iterator<Sheet> it=workbook.sheetIterator();
		int i=0;
		while(it.hasNext()) {
			Sheet sheet=it.next();
			if(sheet.getSheetName().equals(sheetName)) {
				ExcelUtil.switchActiveSheet(workbook, i);
				return;
			}
			i++;
		}
	}
	/**
	 * test case
	 * @param args
	 */
	public static void main(String args[]) {
		Workbook workbook=ExcelUtil.readExcel("D:\\gitcode\\ArcareExcelToPdf\\resource\\ColorTest.xlsx");
		ExcelUtil.switchActiveSheet(workbook, 0);
		ExcelUtil.switchActiveSheetByName(workbook, "工作表2");
		ExcelUtil.savExcel(workbook, "D:\\gitcode\\ArcareExcelToPdf\\resource\\ColorTest.xlsx");
	}
}

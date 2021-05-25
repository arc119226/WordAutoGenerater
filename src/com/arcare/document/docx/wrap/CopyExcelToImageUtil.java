package com.arcare.document.docx.wrap;


import java.awt.Graphics;
import java.awt.Image;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Workbook;
/**
 * 以模擬桌面操作方式複製excel表格為圖片
 * @author FUHSIANG_LIU
 *
 */
public class CopyExcelToImageUtil {
	/**
	 * 程式運行速度
	 */
	private static final int speed=1000;
	/**
	 * test case 模擬該方法每次只有一個執行續可以訪問
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		Toolkit.getDefaultToolkit().setLockingKeyState(KeyEvent.VK_NUM_LOCK, false);//開放鍵齊用
		for(int i=1;i<=10;i++) {
			Thread t=new Thread() {
				public void run() {
					System.out.println("start");
					String resultPath = CopyExcelToImageUtil.generateImage("D:\\gitcode\\ArcareExcelToPdf\\resource\\ColorTest.xlsx","./",0);
					System.out.println("end");
					System.out.println(resultPath+" "+new File(resultPath).exists());
				};
			};
			t.start();
		}

	}
	/**
	 * 
	 * @param excelFilePath
	 * @param outputPath
	 * @param sheetName
	 * @return
	 */
	public static synchronized String generateImage(String excelFilePath,String outputPath,String sheetName) {
		try {
			if(!new File(excelFilePath).isFile()) {
				throw new Exception("ERROR:"+excelFilePath+" file not found.");
			}
			
			Workbook workbook=ExcelUtil.readExcel(excelFilePath);
			ExcelUtil.switchActiveSheetByName(workbook, sheetName);
			ExcelUtil.savExcel(workbook, excelFilePath);
			
			String time = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
			String imagePath=outputPath+ time + ".jpg";
			CopyExcelToImageUtil.createImageFileFromClip(imagePath);
			CopyExcelToImageUtil.mouseClickAndEnter((int)Toolkit.getDefaultToolkit().getScreenSize().getWidth() - 15, 12);
			if(new File(imagePath).exists()) {
				return imagePath;
			}else {
				return null;
			}
		} catch (Exception e) {
			Log.error(e);
			return null;
		}
	}
	/**
	 * 
	 * @param excelFilePath
	 * @param outputPath
	 * @param sheetIndex
	 * @return
	 */
	public static synchronized String generateImage(String excelFilePath,String outputPath,int sheetIndex) {
		try {
			if(!new File(excelFilePath).isFile()) {
				throw new Exception("ERROR:"+excelFilePath+" file not found.");
			}
			
			Workbook workbook=ExcelUtil.readExcel(excelFilePath);
			ExcelUtil.switchActiveSheet(workbook, sheetIndex);
			ExcelUtil.savExcel(workbook, excelFilePath);
			
			CopyExcelToImageUtil.openExcelFileInDesktop(excelFilePath);
			CopyExcelToImageUtil.jumpToSheet(sheetIndex);
			CopyExcelToImageUtil.selectAllDataInExcelThenCopy();
			
			String time = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
			String imagePath=outputPath+ time + ".jpg";
			CopyExcelToImageUtil.createImageFileFromClip(imagePath);
			CopyExcelToImageUtil.mouseClickAndEnter((int)Toolkit.getDefaultToolkit().getScreenSize().getWidth() - 15, 12);
			if(new File(imagePath).exists()) {
				return imagePath;
			}else {
				return null;
			}
		} catch (Exception e) {
			Log.error(e);
			return null;
		}
	}

	/**
	 *  點滑鼠於座標x,y 然後按下enter
	 * @param x
	 * @param y
	 * @throws Exception
	 */
	public static void mouseClickAndEnter(int x, int y) throws Exception {
		Robot robot = new Robot();

		robot.delay(speed/2);
		robot.mouseMove(x, y);
		
		robot.delay(speed/2);
		robot.mousePress(InputEvent.BUTTON1_MASK);

		robot.delay(speed/2);
		robot.mouseRelease(InputEvent.BUTTON1_MASK);
		
		robot.delay(speed/2);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		
		robot.delay(speed/2);
//		robot.delay(100);
//		robot.keyPress(KeyEvent.VK_ENTER);
//		robot.keyRelease(KeyEvent.VK_ENTER);
	}

	/**
	 * 開啟excel
	 * @param file
	 * @throws Exception
	 */
	private static void openExcelFileInDesktop(String file) throws Exception {
		Runtime.getRuntime().exec("cmd /k soffice.exe -o " + file);
	}

	private static void jumpToSheet(int i) throws Exception {
		Robot robot = new Robot();
		robot.delay(speed*4);
		
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_SPACE);
		robot.keyRelease(KeyEvent.VK_ALT);
		robot.keyRelease(KeyEvent.VK_SPACE);

		robot.delay(speed/9);
		robot.keyPress(KeyEvent.VK_UP);
		robot.keyRelease(KeyEvent.VK_UP);
		robot.keyPress(KeyEvent.VK_UP);
		robot.keyRelease(KeyEvent.VK_UP);

		robot.delay(speed/9);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);

		robot.delay(speed/9);
		robot.keyPress(KeyEvent.VK_UP);
		robot.keyRelease(KeyEvent.VK_UP);
		robot.keyPress(KeyEvent.VK_UP);
		robot.keyRelease(KeyEvent.VK_UP);
		
		robot.delay(speed/9);
		
		//切換sheet
//		for(int jumpTime=0;jumpTime<=255;jumpTime++) {
//			robot.delay(speed/90);
//			robot.keyPress(KeyEvent.VK_CONTROL);
//			robot.keyPress(KeyEvent.VK_PAGE_UP);
//			robot.keyRelease(KeyEvent.VK_CONTROL);
//			robot.keyRelease(KeyEvent.VK_PAGE_UP);
//		}
//		for(int c=1;c<=i;c++) {
//			robot.delay(speed/90);
//			robot.keyPress(KeyEvent.VK_CONTROL);
//			robot.keyPress(KeyEvent.VK_PAGE_DOWN);
//			robot.keyRelease(KeyEvent.VK_CONTROL);
//			robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
//		}
	}
	/**
	 *  全選資料然後複製到剪貼簿
	 * @throws Exception
	 */
	private static void selectAllDataInExcelThenCopy() throws Exception {
		Robot robot = new Robot();

		robot.delay(speed/2);
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_HOME);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_HOME);
		
		robot.delay(speed/2);
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_HOME);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_HOME);

		robot.delay(speed/2);
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_SHIFT);
		robot.keyPress(KeyEvent.VK_END);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_SHIFT);
		robot.keyRelease(KeyEvent.VK_END);

		robot.delay(speed/2);
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_SHIFT);
		robot.keyPress(KeyEvent.VK_END);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_SHIFT);
		robot.keyRelease(KeyEvent.VK_END);
		
		robot.delay(speed/2);
//		robot.delay(500);
//		robot.keyPress(KeyEvent.VK_ALT);
//		robot.keyPress(KeyEvent.VK_O);
//		robot.keyRelease(KeyEvent.VK_ALT);
//		robot.keyRelease(KeyEvent.VK_O);
//		robot.delay(100);
//		robot.keyPress(KeyEvent.VK_C);
//		robot.keyRelease(KeyEvent.VK_C);
//		robot.keyPress(KeyEvent.VK_A);
//		robot.keyRelease(KeyEvent.VK_A);

//		robot.delay(500);
//		robot.keyPress(KeyEvent.VK_ALT);
//		robot.keyPress(KeyEvent.VK_O);
//		robot.keyRelease(KeyEvent.VK_ALT);
//		robot.keyRelease(KeyEvent.VK_O);
//		robot.delay(100);
//		robot.keyPress(KeyEvent.VK_R);
//		robot.keyRelease(KeyEvent.VK_R);
//		robot.keyPress(KeyEvent.VK_A);
//		robot.keyRelease(KeyEvent.VK_A);

		robot.delay(speed);
		
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_C);
		
		robot.delay(speed);
	}
	/**
	 * 從剪貼簿取得圖片
	 * @param dir
	 * @throws Exception
	 */
	private static void createImageFileFromClip(String dir) throws Exception {
		Transferable t = Toolkit.getDefaultToolkit().getSystemClipboard().getContents(null);
		if (null != t && t.isDataFlavorSupported(DataFlavor.imageFlavor)) {
			Image image = (Image) t.getTransferData(DataFlavor.imageFlavor);
			savePImage(image, dir);
		}
	}
	/**
	 * 存檔
	 * @param iamge
	 * @param dir
	 * @return
	 * @throws Exception
	 */
	private static String savePImage(Image iamge, String dir) throws Exception {
		int w = iamge.getWidth(null);
		int h = iamge.getHeight(null);
		BufferedImage bi = new BufferedImage(w, h, BufferedImage.TYPE_3BYTE_BGR);
		Graphics g = bi.getGraphics();
		g.drawImage(iamge, 0, 0, null);
		ImageIO.write(bi, "jpg", new File(dir));
		return dir;
	}
}

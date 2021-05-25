package com.arcare.document.docx.core.v2;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Queue;
import java.util.TreeMap;
import java.util.stream.Collectors;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlObject;

import com.arcare.document.docx.vo.DefineVO;
import com.arcare.document.docx.wrap.Log;
import com.arcare.document.docx.wrap.WordDocUtil;

/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class WordMultiColumnProviderV2 {
	/**
	 * 
	 * @param docx
	 * @param tableBookmark
	 * @param dataMap
	 * @return
	 * @throws IOException
	 */
	 private static Map<Integer,String> prepareTableMapping(XWPFDocument docx,String tableBookmark, DefineVO preFixObj){ 
		//find table cell with bookmarks
		//one time n step
		final Map<Integer,String> tableIndexDataMapiing=new TreeMap<>();
		docx.getTables().stream()
			.filter(table->table.getCTTbl().toString().matches(String.format("[.\\S\\s]*<w:bookmarkStart w:id=\"\\d+\" w:name=\"%s\"\\s?\\/>[.\\S\\s]*", tableBookmark)))
		.forEach(table->{
			List<String> keys=preFixObj.getDataFields().stream()
					.map(it->it.getVarName())
					.collect(Collectors.toList());

			List<XWPFTableRow> tableRows = table.getRows();
			for(int i=0;i<tableRows.size();i++){
				XWPFTableRow row=tableRows.get(i);
				for(String key:keys){
					if(row.getCtRow().toString().matches(String.format("[.\\S\\s]*<w:bookmarkStart w:id=\"\\d+\" w:name=\"%s\"\\s?\\/>[.\\S\\s]*", key))){
						tableIndexDataMapiing.put(i,key);
		    		}
				}
			}
		});
		return tableIndexDataMapiing;      
	} 
	
	/**
	 * 處理圖片TABLE
	 * @param docx
	 * @param templateFilePath
	 * @param imageDataSource
	 * @return
	 */
	public static XWPFDocument processMultiColumn(XWPFDocument docx,List<DefineVO> multiColumnV2,String imgDir) {
		Log.log("processMultiColumn start");
		if(multiColumnV2==null) {
			Log.log("processMultiColumn end");
			return docx;
		}
		multiColumnV2.forEach(prefixObj->{
			final String prefix=prefixObj.getVarName();
			final String tableBookMark=prefixObj.getVarName()+"_TABLE";

			Map<Integer,String> keyIndexMap=prepareTableMapping(docx,tableBookMark,prefixObj);

			List<XWPFTable> tables = docx.getTables();
			Optional<XWPFTable> optTable=tables.stream()
					.filter(table->table.getCTTbl().toString().matches(String.format("[.\\S\\s]*<w:bookmarkStart w:id=\"\\d+\" w:name=\"%s\"\\s?\\/>[.\\S\\s]*", tableBookMark)))
					.findFirst();
			if(optTable.isPresent()) {
				
				//copy font style
				Map<String,Object> styleMap=new HashMap<>();
				Optional<XWPFTableRow> opRow = optTable.get().getRows().stream().filter(r->r.getTableCells().size()>1).findFirst();
				if(opRow.isPresent()) {
					Optional<XWPFTableCell> opCell=opRow.get().getTableCells().stream().findFirst();
					if(opCell.isPresent()) {
						XWPFTableCell cell=opCell.get();
						Optional<XWPFParagraph> opPara=cell.getParagraphs().stream().findFirst();
						if(opPara.isPresent()) {
							XWPFParagraph para=opPara.get();
							Optional<XWPFRun> opRun=para.getRuns().stream().findFirst();
							if(opRun.isPresent()) {
								XWPFRun run=opRun.get();
								//v2 flag
								if(run.getCTR()!=null && run.getCTR().getRPr()!=null) {
									styleMap.put("RPr", run.getCTR().getRPr().copy());
								}
							}
						}
					}
				}
				
				XWPFTable table=optTable.get();
				List<XWPFTableRow> tableRows = table.getRows();
				int defaultHeight = tableRows.get(0).getHeight();
				int increaseStep=keyIndexMap.size();
            	//一次取得N列
            	int cellSize=0;
            	int rowIndex=0;
            	for(;rowIndex < tableRows.size(); rowIndex += increaseStep){
    	    		for(int stepKeyIndex = rowIndex,step=0;step < increaseStep; stepKeyIndex++,step++){
    	    			
    	    			String log=String.format("rowIndex %s,increaseStep %s,step %s,stepKeyIndex %s", rowIndex,increaseStep,step,stepKeyIndex);
    	    			System.out.println(log);
    	    			
    	    			String rowDataKey=keyIndexMap.get(stepKeyIndex%increaseStep);
    	    			System.out.println("rowDataKey:"+rowDataKey);

    	    			XWPFTableRow row=tableRows.get(stepKeyIndex);
    	    			row.setHeight(defaultHeight);
    	    			List<XWPFTableCell> celllist=row.getTableCells();
    	    			cellSize=celllist.size();
    	    			
    	    			celllist.forEach(it->{
    	    				Queue<String> dataQueue=prefixObj.getDataFields().stream()
    	    						.filter(o->o.getVarName().equals(rowDataKey))
    	    						.findFirst().get().getDatas();
    	    				String obj=dataQueue.poll();
    	    				if(obj!=null){

    	    					if(prefixObj.getDataFields().stream()
        	    						.filter(o->o.getVarName().equals(rowDataKey))
        	    						.findFirst().get().getFieldType().contains("text")) {
    	    						if(styleMap.get("RPr")!=null) {
    	    							XWPFParagraph cellp=it.addParagraph();
    	    							XWPFRun cellpr=cellp.createRun();
    	    							cellpr.getCTR().addNewRPr().set(XmlObject.class.cast(styleMap.get("RPr")));
    	    							cellpr.setText(obj);
    	    						}else {
    	    							it.setText(obj);
    	    						}
    	    					}else if(prefixObj.getDataFields().stream()
        	    						.filter(o->o.getVarName().equals(rowDataKey))
        	    						.findFirst().get().getFieldType().contains("photo")) {
    	    							
    	    							Queue<String> filenames=prefixObj.getDataFields().stream()
    		    	    						.filter(o->o.getVarName().equals(prefix+"PhotofileName"))
    		    	    						.findFirst().get().getDatas();
    	    							
	    							if(filenames.peek()!=null) {
	    								String fileaname=filenames.poll();
	    								File image = new File(imgDir+File.separator+fileaname.trim());

        			        			double width = Double.valueOf(prefixObj.getDataFields().stream()
                	    						.filter(o->o.getVarName().equals(rowDataKey))
                	    						.findFirst().get().getWidth());
        			        			width=WordDocUtil.cmToP(width);
        			        			double height = Double.valueOf(prefixObj.getDataFields().stream()
                	    						.filter(o->o.getVarName().equals(rowDataKey))
                	    						.findFirst().get().getHeight());
        			        			height=WordDocUtil.cmToP(height);
        			        			int imgFormat = WordDocUtil.getImageFormat(image.getName());
        			        			
        			        			FileInputStream inputStream=null;
        			        			try {
        			        				inputStream=new FileInputStream(image);
        			        				it.addParagraph().createRun().addPicture(inputStream,imgFormat,image.getName(),Units.toEMU(width),Units.toEMU(height));
        			        			}catch(Exception e1){
        			        				Log.error(e1);
        			        			}finally {
        			        				if(inputStream!=null) {
        			        					try {
													inputStream.close();
												} catch (IOException e) {
													Log.error(e);
												}
        			        				}
        			        			}
	    							}
    	    					}
	    						if(it.getParagraphs().size()>0){
	    							it.removeParagraph(0);
	    						}
    	    				}
    	    			});
    	    		}
            	}
            	System.out.println(prefixObj.getDataFields().stream()
						.filter(o->o.getVarName().equals(keyIndexMap.get(0)))
						.findFirst().get().getDatas().size());
            	System.out.println(keyIndexMap.get(0));
    	    	double generateRow=Math.ceil((double)prefixObj.getDataFields().stream()
						.filter(o->o.getVarName().equals(keyIndexMap.get(0)))
						.findFirst().get().getDatas().size()/cellSize);
    	    	System.out.println("generateRow:"+generateRow);
    	    	for(int t=0;t<keyIndexMap.size()*generateRow;t++) {
    	    		table.createRow().getTableCells();
    	    	}
				
    	    	//insert data
    	    	for(;rowIndex < tableRows.size(); rowIndex += increaseStep){
    	    		for(int stepKeyIndex = rowIndex,step=0;step < increaseStep; stepKeyIndex++,step++){
    	    			
    	    			System.out.println("add point ->"+rowIndex+" ,"+stepKeyIndex);
    	    			
    	    			String rowDataKey=keyIndexMap.get(stepKeyIndex%increaseStep);
    	    			XWPFTableRow row=tableRows.get(stepKeyIndex);//
    	    			List<XWPFTableCell> celllist=row.getTableCells();
    	    			cellSize=celllist.size();
    	    			celllist.forEach(it->{
    	    				Queue<String> dataQueue=prefixObj.getDataFields().stream()
    	    						.filter(o->o.getVarName().equals(rowDataKey))
    	    						.findFirst().get().getDatas();
    	    				String obj=dataQueue.poll();
    	    				if(obj!=null){
    	    					if(prefixObj.getDataFields().stream()
        	    						.filter(o->o.getVarName().equals(rowDataKey))
        	    						.findFirst().get().getFieldType().contains("text")) {
    	    						if(styleMap.get("RPr")!=null) {	
    	    							XWPFParagraph cellp=it.addParagraph();
    	    							XWPFRun cellpr=cellp.createRun();
    	    							cellpr.getCTR().addNewRPr().set(XmlObject.class.cast(styleMap.get("RPr")));
    	    							cellpr.setText(obj);
    	    						}else {
    	    							it.setText(obj);
    	    						}
    	    					}else if(prefixObj.getDataFields().stream()
        	    						.filter(o->o.getVarName().equals(rowDataKey))
        	    						.findFirst().get().getFieldType().contains("photo")) {
    	    						//PhotofileName
    	    						Queue<String> filenames=prefixObj.getDataFields().stream()
		    	    						.filter(o->o.getVarName().equals(prefix+"PhotofileName"))
		    	    						.findFirst().get().getDatas();
	    	    					if(filenames.peek()!=null) {
		    							String fileaname=filenames.poll();
		    							File image = new File(imgDir+File.separator+fileaname.trim());
		    								
        			        			double width = Double.valueOf(prefixObj.getDataFields().stream()
                	    						.filter(o->o.getVarName().equals(rowDataKey))
                	    						.findFirst().get().getWidth());
        			        			width=WordDocUtil.cmToP(width);
        			        			double height = Double.valueOf(prefixObj.getDataFields().stream()
                	    						.filter(o->o.getVarName().equals(rowDataKey))
                	    						.findFirst().get().getHeight());
        			        			height=WordDocUtil.cmToP(height);
        			        			
        			        			int imgFormat = WordDocUtil.getImageFormat(image.getName());

	    								FileInputStream inputStream=null;
        			        			try {
        			        				inputStream=new FileInputStream(image);
        			        				it.addParagraph().createRun().addPicture(inputStream,imgFormat,image.getName(),Units.toEMU(width),Units.toEMU(height));
        			        			}catch(Exception e1){
        			        				Log.error(e1);
        			        			}finally {
        			        				if(inputStream!=null) {
        			        					try {
													inputStream.close();
												} catch (IOException e) {
													Log.error(e);
												}
        			        				}
        			        			}
    	    						}
    	    					}
	    						if(it.getParagraphs().size()>0){
	    							it.removeParagraph(0);
	    						}
    	    				}
    	    			});
    	    		}
    	    	}
			}else {
				//no table
			}
		});
		Log.log("processMultiColumn end");
		return docx;
	}

}

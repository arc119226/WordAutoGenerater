package com.arcare.document.docx.core.v2;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Queue;
import java.util.TreeMap;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

import com.arcare.document.docx.vo.DefineVO;
import com.arcare.document.docx.wrap.Log;
import com.arcare.document.docx.wrap.WordDocUtil;

/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class WordMultiHeaderBodyHProviderV2 {

	/**
	 * 
	 * @param docx
	 * @param multiHeaderVH2
	 * @param multiBodyVH2
	 * @return
	 */
	public static XWPFDocument processAllHeaderBodyTableH(XWPFDocument docx,List<DefineVO> multiHeaderVH2,List<DefineVO> multiBodyVH2) {
		Log.log("processAllHeaderBodyTableH start");
		if(multiHeaderVH2==null) {
			Log.log("processAllHeaderBodyTableH end");
			return docx;
		}
		if(multiBodyVH2==null) {
			Log.log("processAllHeaderBodyTableH end");
			return docx;
		}
		prepareDataHeaderQueueMap(multiHeaderVH2,multiBodyVH2).forEach((key,headerQueue)->{	
			
			Optional<DefineVO> optHeaderPrefix=multiHeaderVH2.stream().filter(h->h.getVarName().startsWith(key.substring(0, 3))).findFirst();
			DefineVO headerPrefix=null;
			if(optHeaderPrefix.isPresent()) {
				headerPrefix=optHeaderPrefix.get();
			}
			
			Optional<DefineVO> optBodyPrefix = headerPrefix.getBody().stream().findFirst();
			DefineVO bodyPrefix=null;
			if(optBodyPrefix.isPresent()) {
				bodyPrefix=optBodyPrefix.get();
			}
			
			Map<String,Queue<String>> dataBodyQueueMap = prepareDataBodyQueueMap(bodyPrefix.getVarName(),headerPrefix.getBody().stream().findFirst().get().getDataFields());
			
			Optional<XWPFTable> optTable = WordDocUtil.findBodyTableByBookMark(docx,bodyPrefix.getVarName()+"_TABLE");
			if(optTable.isPresent()) {
				List<XmlCursor> cursorList = new ArrayList<>();
				int i=0;
				while(headerQueue.peek() != null){
					String headerTitle = headerQueue.poll();
					if(cursorList.size() == 0){
						XmlCursor tmpTblCursor = optTable.get().getCTTbl().newCursor();
						cursorList.add(tmpTblCursor);
					}
					cursorList.get(cursorList.size()-1).toEndToken();
					while(cursorList.get(cursorList.size()-1).toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START);
	
					XWPFParagraph cp=docx.insertNewParagraph(cursorList.get(cursorList.size()-1));
					XmlCursor cpCursor = cp.getCTP().newCursor();
					cursorList.add(cpCursor);
					XWPFTable cloneTable = WordDocUtil.copyTable(optTable.get(),docx,cursorList.get(cursorList.size()-1));
					insertDataToTable(i,cloneTable,bodyPrefix.getVarName(),headerTitle,dataBodyQueueMap,bodyPrefix.getRelaFieldName());
					i++;
				}
				//remove template Table
				int position = docx.getPosOfTable(optTable.get());
				docx.removeBodyElement(position);
			}
		});
		Log.log("processAllHeaderBodyTableH end");
		return docx;
	}
	
	/**
	 * 
	 * @param index
	 * @param currentTable
	 * @param prefix
	 * @param headerTitle
	 * @param dataBodyQueueMap
	 * @return
	 */
	private static XWPFTable insertDataToTable(int index,XWPFTable currentTable,String prefix,String headerTitle,Map<String,Queue<String>> dataBodyQueueMap,String relaFieldName){
		Map<Integer,String> indexMapping=new HashMap<>();
		XWPFTableRow row=currentTable.getRows().get(1);
		List<XWPFTableCell> cells=row.getTableCells();
		for(int i=0;i<cells.size();i++){
			XWPFTableCell cell=cells.get(i);
			cell.getCTTc();
			indexMapping.put(i, cell.getText().trim());
		}
		
		Queue<String> refQueue=dataBodyQueueMap.get(prefix+relaFieldName);
		
		if(!refQueue.peek().equals(headerTitle)) {
			return currentTable;
		}
		
		Map<Integer,String> _indexColorMaping=new HashMap<>();
		Map<Integer,Object> _indexStyleMap=new HashMap<>();
		Map<Integer,Object> _indexAlignment=new HashMap<>();
		Iterator<Integer> itPrepareColor=indexMapping.keySet().iterator();
		while(itPrepareColor.hasNext()) {
			Integer _index=itPrepareColor.next();
			if(_indexColorMaping.get(_index)==null) {
				Optional<XWPFParagraph> p=currentTable.getRows().get(0).getTableCells().get(_index).getParagraphs().stream().findFirst();
				if(p.isPresent()) {
					XWPFParagraph _p=p.get();
					Optional<XWPFRun> or=_p.getRuns().stream().findFirst();
					if(or.isPresent()) {
						String color=or.get().getColor();
						_indexColorMaping.put(_index, color);
					}
				}
			}
			if(_indexAlignment.get(_index)==null) {
				Optional<XWPFParagraph> p=currentTable.getRows().get(0).getTableCells().get(_index).getParagraphs().stream().findFirst();
				if(p.isPresent()) {
					XWPFParagraph _p=p.get();
					_indexAlignment.put(_index, _p.getAlignment());
				}
			}
			if(_indexStyleMap.get(_index)==null) {
				Optional<XWPFParagraph> p=currentTable.getRows().get(0).getTableCells().get(_index).getParagraphs().stream().findFirst();
				if(p.isPresent()) {
					XWPFParagraph _p=p.get();
					Optional<XWPFRun> or=_p.getRuns().stream().findFirst();
					if(or.isPresent()) {
						_indexStyleMap.put(_index, or.get().getCTR().getRPr().copy());
					}
				}
			}
		}
		
		if(index==0) {
			//remain first row
			for(int i=0;i<currentTable.getRows().size();i++) {
				currentTable.removeRow(1);
			}
		}else {
			currentTable.removeRow(0);
			for(int i=0;i<currentTable.getRows().size();i++) {
				currentTable.removeRow(0);
			}
		}
		
		boolean isFirstRow=false;
		boolean isFirstCell=false;
//		start insert data
		while(refQueue.peek() != null && 
			  refQueue.peek().equals(headerTitle)){//is reference
			String item=refQueue.poll();
			Iterator<Integer> it=indexMapping.keySet().iterator();
			XWPFTableRow _r=currentTable.createRow();
			
			
			while(it.hasNext()){
				Integer _index=it.next();
				String key=indexMapping.get(_index);
				if(!key.contains(prefix)) {
					key=key.replace(key.subSequence(0, 3), prefix);
				}
				XWPFTableCell cell=_r.getTableCells().get(_index);
				
				if(!key.equals(prefix+relaFieldName)) {
					String data=dataBodyQueueMap.get(key).poll();
					XWPFParagraph cellp=cell.addParagraph();
					if(_indexAlignment.get(_index)!=null) {
						cellp.setAlignment(ParagraphAlignment.class.cast(_indexAlignment.get(_index)));
					}
					XWPFRun cellpr=cellp.createRun();
					if(_indexStyleMap.get(_index)!=null) {
						cellpr.getCTR().addNewRPr().set(XmlObject.class.cast(_indexStyleMap.get(_index)));
					}
					if(_indexColorMaping.get(_index)!=null) {
						cellpr.setColor(_indexColorMaping.get(_index));
					}
					cellpr.setText(data);
					cell.removeParagraph(0);
				}else {
					if(!isFirstCell) {

						XWPFParagraph cellp=cell.addParagraph();
						if(_indexAlignment.get(_index)!=null) {
							cellp.setAlignment(ParagraphAlignment.class.cast(_indexAlignment.get(_index)));
						}
						XWPFRun cellpr=cellp.createRun();
						if(_indexStyleMap.get(_index)!=null) {
							cellpr.getCTR().addNewRPr().set(XmlObject.class.cast(_indexStyleMap.get(_index)));
						}
						if(_indexColorMaping.get(_index)!=null) {
							cellpr.setColor(_indexColorMaping.get(_index));
						}
						cellpr.setText(item);

						isFirstCell=true;
					}
				}
			}
			//remove row 0
			if(index>0) {
				if(!isFirstRow) {
					currentTable.removeRow(0);
					isFirstRow=true;
				}
			}
		}
		
//		do merge cell
		if(index>0) {
			WordDocUtil.mergeCellsInColumn(currentTable,0,0,currentTable.getRows().size()-1);
		}else {
			WordDocUtil.mergeCellsInColumn(currentTable,0,1,currentTable.getRows().size()-1);
		}
		return currentTable;
	}
	
	/**
	 * prepare header queue
	 * @param headerV
	 * @return
	 */
	private static Map<String,Queue<String>> prepareDataHeaderQueueMap(List<DefineVO> multiHeaderVV2,List<DefineVO> multiBodyVV2){
		Map<String,Queue<String>> dataHeaderQueueMap = new TreeMap<>();
		multiHeaderVV2.forEach(prefixObj->{
			String preFix=prefixObj.getVarName();

			Optional<DefineVO> optBody=multiBodyVV2.stream()
				.filter(bodyPrefix->bodyPrefix.getRelaviewname().equals(prefixObj.getViewName()))
				.findFirst();
			if(optBody.isPresent()) {
				DefineVO body=optBody.get();
				String relName=preFix+body.getRelaFieldName();
				dataHeaderQueueMap.put(relName,prefixObj.getDataFields().stream()
						.filter(p->p.getVarName().endsWith(body.getRelaFieldName())).findFirst().get().getDatas());
			}
		});
		return dataHeaderQueueMap;
	}
	
	/**
	 * prepare body queue
	 * @param bodyBookMarkPrefix
	 * @param bodyV
	 * @return
	 */
	private static Map<String,Queue<String>> prepareDataBodyQueueMap(String bodyBookMarkPrefix,List<DefineVO> dataFields){
		Map<String,Queue<String>> dataBodyQueueMap=new TreeMap<>();
		dataFields.forEach(field->{
			dataBodyQueueMap.put(field.getVarName(), field.getDatas());
		});
		return dataBodyQueueMap;
	}
}

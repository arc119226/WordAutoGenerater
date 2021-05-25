package com.arcare.document.docx.core.v2;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
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
public class WordMultiHeaderBodyVProviderV2 {
	/**
	 * 
	 * @param docx
	 * @param headerV
	 * @param bodyV
	 * @return
	 */
	public static XWPFDocument processAllHeaderBodyTableV(XWPFDocument docx,List<DefineVO> headerV,List<DefineVO> bodyV) {
		Log.log("processAllHeaderBodyTableV start");
		if(headerV==null) {
			Log.log("processAllHeaderBodyTableV end");
			return docx;
		}
		if(bodyV==null) {
			Log.log("processAllHeaderBodyTableV end");
			return docx;
		}
		Map<String,Queue<String>> newheaderQueueMap=prepareDataHeaderQueueMap(headerV,bodyV);
		List<String> headerList=new ArrayList<>();
		headerV.forEach(header->{
			final String headerPrefix = header.getVarName();//前三碼
			Log.log("MULTI HEADER前三碼:"+headerPrefix);
			
			//header 段落範本 list
			List<Map<String,Optional<XWPFParagraph>>> templateParaList = new ArrayList<>();
			//header 表格範本 list XXX_TABLE
			List<Map<String,Optional<XWPFTable>>> templateTableList = new ArrayList<>();
			
			newheaderQueueMap.forEach((k,v)->{
				if(k.startsWith(headerPrefix)) {
					Log.log("\tMULTI HEADER書籤:"+k);
					//查段落範本
					Optional<XWPFParagraph> templateParagraph = WordDocUtil.findHeaderParagraphByBookMark(docx,k);
					if(templateParagraph.isPresent()) {
						Log.log("\t找到HEADER段落");
						Map<String,Optional<XWPFParagraph>> obj=new TreeMap<>();
						obj.put(k, templateParagraph);
						templateParaList.add(obj);
					}
				}
			});

			//查HEADER表格範本
			Optional<XWPFTable> templateTable = WordDocUtil.findBodyTableByBookMark(docx, headerPrefix+"_TABLE");
			if(templateTable.isPresent()) {
				Log.log("\t找到HEADER表格");
				Map<String,Optional<XWPFTable>> obj=new TreeMap<>();
				obj.put(headerPrefix+"_TABLE", templateTable);
				templateTableList.add(obj);
			}
			
			Optional<DefineVO> optBodyPrefix = header.getBody().stream().findFirst();
			optBodyPrefix.ifPresent(bodyPrefix->{
				Log.log("\t找到body vo");
				//prepare body map data
				Map<String,Queue<String>> dataBodyQueueMap = prepareDataBodyQueueMap(bodyPrefix.getVarName(),bodyPrefix.getDataFields());
				dataBodyQueueMap.forEach((bodyk,bodyv)->{
					Log.log("\t\tbodyk:"+bodyk);
					Log.log("\t\t\tbodyv:"+bodyv);
				});
				Log.log("\tbody前三碼:"+bodyPrefix.getVarName());
				Optional<XWPFTable> optBodyTable = WordDocUtil.findBodyTableByBookMark(docx,bodyPrefix.getVarName()+"_TABLE");
				if(optBodyTable.isPresent()) {
					Log.log("\t找到body表格");
					List<XmlCursor> cursorList = new ArrayList<>();//xml指標list
					//處理header段落
					templateParaList.forEach(map->{
						map.forEach((key,optPara)->{
							Queue<String> headerQueue = newheaderQueueMap.get(key);
							while(headerQueue.peek() != null) {
								StringBuffer headerTitle = new StringBuffer();
								StringBuffer relationTitle = new StringBuffer();
								
								if(!key.contains(optBodyPrefix.get().getRelaFieldName())) {
									headerTitle.append(headerQueue.poll());
									relationTitle.append(newheaderQueueMap.get(header.getVarName()+optBodyPrefix.get().getRelaFieldName()).poll());
									Log.log(headerTitle+" "+relationTitle);
								}else {
									headerTitle.append(headerQueue.poll());
									relationTitle.append(headerTitle.toString());
									Log.log(headerTitle+" "+relationTitle);
								}
								
								headerList.add(headerTitle.toString());
								//取得template
								if(cursorList.size() == 0){
									XmlCursor tmpTblCursor = optBodyTable.get().getCTTbl().newCursor();
									tmpTblCursor.toEndToken();
									cursorList.add(tmpTblCursor);
								}
								XWPFParagraph cp = WordDocUtil.copyParagraphToCurserAndUpdateText(optPara.get(),docx,cursorList.get(cursorList.size()-1),headerTitle.toString());
								XmlCursor cpCursor = cp.getCTP().newCursor();
								cpCursor.toEndToken();
								cursorList.add(cpCursor);
								//處理header表格
								templateTableList.forEach(headerTableMap->{
									headerTableMap.forEach((headerTableK,headerTableV)->{
										if(headerTableK.equals(headerPrefix+"_TABLE")) {
											Log.log("處理"+headerPrefix+"_TABLE");
											XWPFTable cloneHeaderTable = WordDocUtil.copyTable(headerTableV.get(),docx,cursorList.get(cursorList.size()-1));
											//replace clone header
											newheaderQueueMap.forEach((tableBookMark,headertabledataQueue)->{
												if(key.equals(tableBookMark)) {
													if(key.contains(optBodyPrefix.get().getRelaFieldName())) {
														//relation
														Queue<String> queue=new LinkedList<>();
														queue.offer(relationTitle.toString());
														WordDocUtil.replaceTable(cloneHeaderTable, tableBookMark, queue);
													}else {
														//header
														Queue<String> queue=new LinkedList<>();
														queue.offer(headerTitle.toString());
														WordDocUtil.replaceTable(cloneHeaderTable, tableBookMark, queue);
													}
												}else {
													WordDocUtil.replaceTable(cloneHeaderTable, tableBookMark, headertabledataQueue);
												}
											});
										}
									});
								});
								//處理bodytable
								XWPFTable cloneTable = WordDocUtil.copyTable(optBodyTable.get(),docx,cursorList.get(cursorList.size()-1));
								//process newTable..
								insertDataToTable(cloneTable,optBodyPrefix.get().getVarName(),relationTitle.toString(),dataBodyQueueMap,optBodyPrefix.get().getRelaFieldName(),key);
							}//end while
							//remove header,body template Table and template Paragraph
							int position = docx.getPosOfTable(optBodyTable.get());
							docx.removeBodyElement(position);

							templateTableList.forEach(headerTableMap->{
								headerTableMap.forEach((headerTableK,headerTableV)->{
									if(headerTableK.equals(headerPrefix+"_TABLE")) {
										int _position = docx.getPosOfTable(headerTableV.get());
										docx.removeBodyElement(_position);
									}
								});
							});

							position=docx.getPosOfParagraph(optPara.get());
							docx.removeBodyElement(position);
						});
					});
				}

			});
		});
		
		//table之間空一行
		for(int i=0;i<headerList.size();i++) {
			String header=headerList.get(i);
			if(i>0) {
				docx.getParagraphs().stream()
					.filter(p->p.getParagraphText().trim().equals(header.trim()))
					.findFirst()
					.ifPresent(p->{
						String str=p.getParagraphText();
						XmlObject ctr=p.getRuns().stream().findFirst().get().getCTR().copy();
						p.getRuns().forEach(r->{
							r.setText("", 0);
						});
						XWPFRun r0=p.createRun();
//						r0.addBreak();
						XWPFRun r=p.createRun();
						r.setText(str);
						r.getCTR().set(ctr);
					});
			}
		}
		
		Log.log("processAllHeaderBodyTableV end");
		return docx;

	}
	
	/**
	 * 處理body table 依據 header
	 * @param currentTable
	 * @param prefix
	 * @param headerTitle
	 * @param dataBodyQueueMap
	 * @param relaFieldName
	 * @param headerName
	 * @return
	 */
	private static XWPFTable insertDataToTable(
			XWPFTable currentTable,//當前處理的table
			String prefix,//body Prefix
			String headerTitle,//header 斷落的文字
			Map<String,
			Queue<String>> dataBodyQueueMap,
			String relaFieldName,
			String headerName){

		Map<Integer,String> indexMapping=new HashMap<>();
		int defaultRow = currentTable.getRows().size();

		XWPFTableRow row=currentTable.getRows().get(currentTable.getRows().size()-1);

		List<XWPFTableCell> cells=row.getTableCells();
		for(int i=0;i<cells.size();i++){
			XWPFTableCell cell=cells.get(i);
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
				Optional<XWPFParagraph> p=currentTable.getRows().stream()
					.filter(r->r.getTableCells().size()==indexMapping.keySet().size())
					.findFirst().get().getTableCells().get(_index).getParagraphs().stream().findFirst();
				p.ifPresent(_p->{
					_p.getRuns().stream().findFirst().ifPresent(r->{
						String color=r.getColor();
						_indexColorMaping.put(_index, color);
					});
				});
			}
			
			if(_indexAlignment.get(_index)==null) {
				Optional<XWPFParagraph> p=currentTable.getRows().stream()
					.filter(r->r.getTableCells().size()==indexMapping.keySet().size())
					.findFirst().get().getTableCells().get(_index).getParagraphs().stream().findFirst();
				p.ifPresent(_p->{
					_indexAlignment.put(_index, _p.getAlignment());
				});
			}
			
			if(_indexStyleMap.get(_index) == null) {
				Optional<XWPFParagraph> p=currentTable.getRows().stream()
					.filter(r->r.getTableCells().size() == indexMapping.keySet().size())
					.findFirst().get().getTableCells().get(_index).getParagraphs().stream().findFirst();
				p.ifPresent(_p->{
					_p.getRuns().stream().findFirst().ifPresent(r->{
						_indexStyleMap.put(_index, r.getCTR().getRPr().copy());
					});
				});
			}
		}
		
		//remove template row
		int needDeleteRow=currentTable.getRows().size()-1;

		while(refQueue.peek() != null && refQueue.peek().equals(headerTitle)){//is reference
			refQueue.poll();
			Iterator<Integer> it=indexMapping.keySet().iterator();
			XWPFTableRow _r=currentTable.createRow();
			int currentSize=_r.getTableCells().size();
			if(indexMapping.keySet().size()>currentSize) {
				for(int i=0;i<indexMapping.keySet().size()-currentSize;i++) {
					_r.createCell();
				}
			}
			while(it.hasNext()){
				Integer _index=it.next();
				String key=indexMapping.get(_index);
				XWPFTableCell cell=_r.getTableCells().get(_index);
				XWPFParagraph cellp=cell.addParagraph();
				XWPFRun cellpr=cellp.createRun();
				if(_indexAlignment.get(_index)!=null) {
					cellp.setAlignment(ParagraphAlignment.class.cast(_indexAlignment.get(_index)));
				}
				cellpr.setText(dataBodyQueueMap.get(key).poll());
				cell.removeParagraph(0);
			}
		}
		
		if(defaultRow == 1) {
			//如果範本只有一行 複製原有樣式
			WordDocUtil.copyTableStyleFull(currentTable, indexMapping.keySet().size(), false);
		}else if(defaultRow > 1){
			//如果範本有多行 找第一個符合的行複製樣式
			WordDocUtil.copyTableStyleFromFirstTarget(currentTable,indexMapping.keySet().size());
		}
		currentTable.removeRow(needDeleteRow);
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
			final String preFix=prefixObj.getVarName();
			multiBodyVV2.stream()
			.filter(bodyPrefix->bodyPrefix.getRelaviewname().equals(prefixObj.getViewName()))
			.findFirst()
			.ifPresent(body->{
				String relName=preFix+body.getRelaFieldName();
				
				prefixObj.getDataFields().stream()
				.filter(p->p.getVarName().startsWith(preFix))
				.forEach(it->{//取得header
					dataHeaderQueueMap.put(it.getVarName(),it.getDatas());
				});

				prefixObj.getDataFields().stream()
				.filter(p->p.getVarName().endsWith(body.getRelaFieldName()))
				.findFirst()
				.ifPresent(it->{
					dataHeaderQueueMap.put(relName, it.getDatas());
				});
			});
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

package com.arcare.document.docx.core.v2;

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Queue;
import java.util.TreeMap;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.arcare.document.docx.vo.DefineVO;
import com.arcare.document.docx.wrap.Log;
import com.arcare.document.docx.wrap.WordDocUtil;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class WordReplaceProviderV2 {
	
	
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
	
	public static void insertMultiRowToTABLE(XWPFDocument docx,List<DefineVO> multiRowV2) {
		Log.log("insertMultiRowToTABLE start");
		if(multiRowV2==null) {
			Log.log("insertMultiRowToTABLE end");
			return;
		}
		multiRowV2.forEach(prefixObj->{
			String tableBookmark = prefixObj.getVarName()+"_TABLE";
			Map<String,Queue<String>> bodyMap=prepareDataBodyQueueMap( prefixObj.getVarName(),prefixObj.getDataFields());

			Optional<XWPFTable> currentTable = WordDocUtil.findBodyTableByBookMark(docx, tableBookmark);
			currentTable.ifPresent(table->{
				int defaultRow=table.getRows().size();
				Map<Integer,String> indexMapping=new HashMap<>();
				List<XWPFTableRow> rows=table.getRows();
				int startRowIndex=0;
				for(int i=0;i<rows.size();i++) {
					boolean containBookmark=rows.get(i).getTableCells().stream().findFirst().get().getCTTc().toString().contains(prefixObj.getVarName());
					if(containBookmark) {
						startRowIndex=i;
					}
				}
				List<XWPFTableCell> cells=rows.get(startRowIndex).getTableCells();
				for(int i=0;i<cells.size();i++){
					XWPFTableCell cell=cells.get(i);
					indexMapping.put(i, cell.getText().trim());
				}

				int times = rows.size()-startRowIndex;

				for(int i=0;i<times;i++) {
					table.removeRow(startRowIndex+1);
				}

				int cellsize=indexMapping.size();
				int dataRows=bodyMap.get(indexMapping.get(0)).size();
				
				for(int i=0;i<dataRows;i++) {
					XWPFTableRow _r=table.createRow();
					if(_r.getTableCells().size()<cellsize) {
						int incrase=cellsize-_r.getTableCells().size();
						for(int j=0;j<incrase;j++) {
							_r.createCell();
						}
					}

					Iterator<Integer> it=indexMapping.keySet().iterator();
					while(it.hasNext()){
						Integer index=it.next();
						String key=indexMapping.get(index);
						XWPFTableCell cell=_r.getTableCells().get(index);;
						if(bodyMap.get(key).peek()!=null) {
							cell.setText(bodyMap.get(key).poll());
						}
					}
				}
				
				if(times>0&&startRowIndex==1) {
					int insert=times-table.getRows().size();
					for(int i=0;i<=insert+1;i++) {
						XWPFTableRow _r=table.createRow();
						if(_r.getTableCells().size()<cellsize) {
							int incrase=cellsize-_r.getTableCells().size();
							for(int j=0;j<incrase;j++) {
								_r.createCell();
							}
						}
					}
				}
				if(defaultRow == 1) {
					WordDocUtil.copyTableStyleFull(table, indexMapping.keySet().size(), false);
				}else if(defaultRow > 1){
					WordDocUtil.copyTableStyleFromIndexRow(startRowIndex,table);
				}
				table.removeRow(startRowIndex);
			});
		});
		Log.log("insertMultiRowToTABLE end");
	}
	
}

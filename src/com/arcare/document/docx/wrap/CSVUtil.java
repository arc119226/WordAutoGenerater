package com.arcare.document.docx.wrap;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import com.arcare.document.docx.vo.DefineVO;

/**
 * parse define csv to vo
 * @author FUHSIANG_LIU
 *
 */
public class CSVUtil {

	/**
	 * 第一層標頭
	 */
	private static final String DataSource = "DataSource";
	/**
	 * 第二層標頭
	 */
	private static final String DataField = "DataField";

	/**
	 * 依據定義檔 產生 定義數據結構
	 * 	key = recordType
	 *	value = datasource vo list
	 * @return
	 * @throws IOException
	 */
	public static Map<String,List<DefineVO>> generateDefineMap(String filePath) throws Exception {
		Log.log("generateDefineMap");
		List<String> csvList= Files.readAllLines(Paths.get(filePath),StandardCharsets.UTF_16);
		
		List<DefineVO> dataList=new ArrayList<>();
		csvList.stream()
		.forEach(it->{
//			String[] line=it.replaceAll(",", " , ").split(",");
			String[] line=it.replaceAll("\t", " \t ").split("\t");
			DefineVO vo=new DefineVO();
			if(line.length>=10) {
				vo.setConfigtype(line[0].trim());
				vo.setVarName(line[1].trim());//.replaceAll("\\-", ""));
				vo.setRecordtype(line[2].trim());
				vo.setColumns(line[3].trim());
				vo.setRelaviewname(line[4].trim());
				vo.setRelaFieldName(line[5].trim());
				vo.setViewName(line[6].trim());
				vo.setFieldName(line[7].trim());//.replaceAll("\\-", ""));
				vo.setFieldType(line[8].trim());
				vo.setFieldLen(line[9].trim());
			}
			if(line.length>=13) {
				vo.setWidth(line[10].trim());
				vo.setHeight(line[11].trim());
				vo.setGap(line[12].trim());
			}
			if(line.length>=17) {
				vo.setOnHeader(line[13].trim());
				vo.setOnFooter(line[14].trim());
				vo.setMergeRow(line[15].trim());
				vo.setMergeColumn(line[16].trim());
			}
			dataList.add(vo);
		});

		//first layer DataSource
		 List<DefineVO> allDatasourceList=dataList.stream()
				 .filter(vo->DataSource.equals(vo.getConfigtype()))
				 .collect(Collectors.toList());
		 
		 allDatasourceList.forEach(d->{
			 if(d.getRelaviewname()!=null) {
				 
				 List<DefineVO> header = allDatasourceList.stream()
						 .filter(f->f.getViewName().equalsIgnoreCase(d.getRelaviewname()))
						 .collect(Collectors.toList());
				 
				 if(!header.isEmpty()) {
					 //relation header and body
					 d.setColumns(header.get(0).getColumns());
					 d.setHeader(header.get(0));
					 if(header.get(0).getBody()==null) {
						 header.get(0).setBody(new ArrayList<>());
						 header.get(0).getBody().add(d);
					 }
				 }
			 }
		 });
		 
		 //second layer DataFieldList
		 List<DefineVO> allDataFieldList=dataList.stream()
				 .filter(vo->DataField.equalsIgnoreCase(vo.getConfigtype()))
				 .collect(Collectors.toList());
		 
		 allDatasourceList.forEach(ds->{
			 List<DefineVO> dataFieldList = allDataFieldList.stream()
			 	.filter(it->it.getVarName().startsWith(ds.getVarName()))
			 	.collect(Collectors.toList());
			 ds.setDataFields(dataFieldList);
		 });
		 
		 //put to map
		 Map<String,List<DefineVO>> recordtypeMap=new HashMap<>();
		 allDatasourceList.forEach(ds->{
			 if(null == recordtypeMap.get(ds.getKey())){
				 List<DefineVO> list=new ArrayList<>();
				 list.add(ds);
				 recordtypeMap.put(ds.getKey(),list);
			 }else {
				 recordtypeMap.get(ds.getKey()).add(ds);
			 }
		 });
		 
		 return recordtypeMap;
	}
}

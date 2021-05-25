package com.arcare.document.docx.vo;

import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.Queue;

/**
 * 
 * @author FUHSIANG_LIU
 * 定義檔對應之數據結構分層物件
 * 第一層 prefix
 *	第二層 對應變數
 *		dataList層 真實資料
 */
public class DefineVO {
	/**
	 * 來源類別
	 */
	private String configtype;
	/**
	 * 變數名
	 */
	private String varName;
	/**
	 *紀錄類型
	 */
	private String recordtype;
	/**
	 * 欄位方向
	 */
	private String columns;
	/**
	 * 關聯表格
	 */
	private String relaviewname;
	/**
	 * 關聯欄位
	 */
	private String relaFieldName;
	/**
	 * 檢視表名
	 */
	private String viewName;
	/**
	 * 欄位名稱
	 */
	private String fieldName;
	/**
	 * 欄位型態
	 */
	private String fieldType;
	/**
	 * 欄位長度
	 */
	private String fieldLen;
	/**
	 * 圖片寬度
	 */
	private String width;
	/**
	 * 圖片高度
	 */
	private String height;
	/**
	 * 間距
	 */
	private String gap;
	/**
	 * 可在頁首
	 */
	private String onHeader;
	/**
	 * 可在頁尾
	 */
	private String onFooter;
	/**
	 * 同值列合併
	 */
	private String mergeRow;
	/**
	 * 同值欄合併
	 * 
	 */
	private String mergeColumn;
	
	private DefineVO header;
	
	private List<DefineVO> body;
	
	private LinkedList<String> datas=new LinkedList<>();
	
	/**
	 * generate download image sql
	 * @param corpDBName
	 * @param caseNo
	 * @return
	 */
	public String getDownloadImageSql(String corpDBName,String caseNo) {
		return String.format("select '%s.jpg' as PHOTOFILENAME, %s as PHOTODATA from %s..%s where CASENO='%s'", 
					this.varName,this.fieldName,corpDBName,this.viewName,caseNo);
	}
	
	/**
	 * 載入設定檔 產勝對應的KEY
	 * @return
	 */
	public String getKey() {
		if(header != null || body!=null ) {
			return this.getRecordtype()+(this.getColumns()==null?"":this.getColumns());
		}else {
			return this.getRecordtype();
		}
	}
	
	private List<DefineVO> dataFields=new ArrayList<>();
	
	public String getConfigtype() {
		return configtype;
	}

	public void setConfigtype(String configtype) {
		this.configtype = configtype;
	}

	public String getColumns() {
		return columns;
	}

	public void setColumns(String columns) {
		this.columns = columns;
	}

	public String getRelaviewname() {
		return relaviewname;
	}

	public void setRelaviewname(String relaviewname) {
		this.relaviewname = relaviewname;
	}

	public String getRelaFieldName() {
		return relaFieldName;
	}

	public void setRelaFieldName(String relaFieldName) {
		this.relaFieldName = relaFieldName;
	}

	public String getViewName() {
		return viewName;
	}

	public void setViewName(String viewName) {
		this.viewName = viewName;
	}

	public String getFieldName() {
		return fieldName;
	}

	public void setFieldName(String fieldName) {
		this.fieldName = fieldName;
	}

	public String getFieldType() {
		return fieldType;
	}

	public void setFieldType(String fieldType) {
		this.fieldType = fieldType;
	}

	public String getFieldLen() {
		return fieldLen;
	}

	public void setFieldLen(String fieldLen) {
		this.fieldLen = fieldLen;
	}

	public String getWidth() {
		return width;
	}

	public void setWidth(String width) {
		this.width = width;
	}

	public String getHeight() {
		return height;
	}

	public void setHeight(String height) {
		this.height = height;
	}

	public String getGap() {
		return gap;
	}

	public void setGap(String gap) {
		this.gap = gap;
	}

	public String getRecordtype() {
		return recordtype;
	}

	public void setRecordtype(String recordtype) {
		this.recordtype = recordtype;
	}

	public String getVarName() {
		return varName;
	}
	
	public void setVarName(String varName) {
		this.varName = varName;
	}
	/**
	 * 取得下層物件
	 * @return
	 */
	public List<DefineVO> getDataFields() {
		return dataFields;
	}
	/**
	 * 設定下層物件
	 * @param dataFields
	 */
	public void setDataFields(List<DefineVO> dataFields) {
		this.dataFields = dataFields;
	}


	@Override
	public String toString() {
		return "DefineVO \t\n[configtype=" + configtype + ", varName=" + varName + ", recordtype=" + recordtype
				+ ", columns=" + columns + ", relaviewname=" + relaviewname + ", relaFieldName=" + relaFieldName
				+ ", viewName=" + viewName + ", fieldName=" + fieldName + ", fieldType=" + fieldType + ", fieldLen="
				+ fieldLen + ", width=" + width + ", height=" + height + ", gap=" + gap + ", onHeader=" + onHeader
				+ ", onFooter=" + onFooter + ", mergeRow=" + mergeRow + ", mergeColumn=" + mergeColumn + ", header=\t\n"
				+ header + ", body= \t\n" + body + ", datas=" + datas + ", dataFields=\t\n" + dataFields + "]";
	}

	/**
	 * 當物件位於prefix層 其結構為body 可取得對應之header物件
	 * @return
	 */
	public DefineVO getHeader() {
		return header;
	}
	/**
	 * 當物件位於prefix 層 其結構為body 設定prefix header物件
	 * @param header
	 */
	public void setHeader(DefineVO header) {
		this.header = header;
	}
	/**
	 * 當物件位於prefix層時 結構為header 可取得 prefix body 物件
	 * @return
	 */
	public List<DefineVO> getBody() {
		return body;
	}
	/**
	 * prefix層 header設定之body關聯
	 * @param body
	 */
	public void setBody(List<DefineVO> body) {
		this.body = body;
	}

	/**
	 * data 加入 隊列
	 * @param data
	 */
	public void setData(String data) {
		if(data!=null) {
			datas.offer(data.trim());
		}
	}
	/**
	 * 以Queue結構處理資料
	 * @return
	 */
	public Queue<String> getDatas(){
		return datas;
	}
	/**
	 * 以List結構處理資料
	 * @return
	 */
	public List<String> getDatasList(){
		return datas;
	}

	public String getOnHeader() {
		return onHeader;
	}

	public void setOnHeader(String onHeader) {
		this.onHeader = onHeader;
	}

	public String getOnFooter() {
		return onFooter;
	}

	public void setOnFooter(String onFooter) {
		this.onFooter = onFooter;
	}

	public String getMergeRow() {
		return mergeRow;
	}

	public void setMergeRow(String mergeRow) {
		this.mergeRow = mergeRow;
	}

	public String getMergeColumn() {
		return mergeColumn;
	}

	public void setMergeColumn(String mergeColumn) {
		this.mergeColumn = mergeColumn;
	}

	public void setDatas(LinkedList<String> datas) {
		this.datas = datas;
	}
	
}

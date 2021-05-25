package com.arcare.document.docx.vo;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class CaseVO {

	private String caseNo;
	/**
	 * 設定檔路徑
	 */
	private String configFilePath;
	/**
	 * 範本檔路徑
	 */
	private String templateFilePath;
	/**
	 * render之後的檔案路徑
	 */
	private String afterRenderFilePath;
	/**
	 * 1 底稿
	 * 2 章節
	 * 5 excel
	 * 6 word
	 * 7 圖片
	 */
	private Integer type;
	/**
	 * 章節位於底稿上的書籤名稱
	 */
	private String bookMark;

	public String getCaseNo() {
		return caseNo;
	}
	public void setCaseNo(String caseNo) {
		this.caseNo = caseNo;
	}
	public String getConfigFilePath() {
		return configFilePath;
	}
	public void setConfigFilePath(String configFilePath) {
		this.configFilePath = configFilePath;
	}
	public String getTemplateFilePath() {
		return templateFilePath;
	}
	public void setTemplateFilePath(String templateFilePath) {
		this.templateFilePath = templateFilePath;
	}
	public Integer getType() {
		return type;
	}
	public void setType(Integer type) {
		this.type = type;
	}
	public String getBookMark() {
		return bookMark;
	}
	public void setBookMark(String bookMark) {
		this.bookMark = bookMark;
	}
	public String getAfterRenderFilePath() {
		return afterRenderFilePath;
	}
	public void setAfterRenderFilePath(String afterRenderFilePath) {
		this.afterRenderFilePath = afterRenderFilePath;
	}
	@Override
	public String toString() {
		return "CaseVO [caseNo=" + caseNo + ", configFilePath=" + configFilePath + ", templateFilePath="
				+ templateFilePath + ", type=" + type + ", bookMark=" + bookMark + "]";
	}
}

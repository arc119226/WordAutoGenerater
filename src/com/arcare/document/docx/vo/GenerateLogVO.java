package com.arcare.document.docx.vo;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class GenerateLogVO {

	private String soureType;

	private String soureKey;

	private String fileItem;

	private Boolean photoResult;

	public String getSoureType() {
		return soureType;
	}

	public void setSoureType(String soureType) {
		this.soureType = soureType;
	}

	public String getSoureKey() {
		return soureKey;
	}

	public void setSoureKey(String soureKey) {
		this.soureKey = soureKey;
	}

	public String getFileItem() {
		return fileItem;
	}

	public void setFileItem(String fileItem) {
		this.fileItem = fileItem;
	}

	public Boolean getPhotoResult() {
		return photoResult;
	}

	public void setPhotoResult(Boolean photoResult) {
		this.photoResult = photoResult;
	}
}

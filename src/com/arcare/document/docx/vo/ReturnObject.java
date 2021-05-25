package com.arcare.document.docx.vo;

import java.util.ArrayList;
import java.util.List;

/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class ReturnObject {

	private Boolean Result;
	
	private String ErrorMessage="";
	
	private List<GenerateLogVO> GenerateLog=new ArrayList<>();

	public Boolean getResult() {
		return Result;
	}

	public void setResult(Boolean result) {
		Result = result;
	}

	public String getErrorMessage() {
		return ErrorMessage;
	}

	public void setErrorMessage(String errorMessage) {
		ErrorMessage = errorMessage;
	}

	public List<GenerateLogVO> getGenerateLog() {
		return GenerateLog;
	}

	public void setGenerateLog(List<GenerateLogVO> generateLog) {
		GenerateLog = generateLog;
	}
	


}

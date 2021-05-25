package com.arcare.document.docx.exception;

import com.arcare.document.docx.vo.ReturnObject;

/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class WordGenerateException extends Exception {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	private ReturnObject result = new ReturnObject();

	public ReturnObject getResult() {
		return result;
	}

	public void setResult(ReturnObject result) {
		this.result = result;
	}

}

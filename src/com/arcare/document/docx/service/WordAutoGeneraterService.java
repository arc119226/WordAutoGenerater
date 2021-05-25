package com.arcare.document.docx.service;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public interface WordAutoGeneraterService {
	
	/**
	 * v1 版本接口
	 * @param applyNumber
	 * @param reporCategory
	 * @param revision
	 * @return
	 */
	public String generateReport(String applyNumber, String reporCategory, String revision);
	/**
	 * 
	 * @param caseNo
	 * @param revision
	 * @param corpDBName
	 * @return
	 */
	public String generateReportV2(String caseNo, String revision,String corpDBName,String FileName);
}

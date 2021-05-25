package com.arcare.document.docx.controller;

import java.io.IOException;
import java.io.PrintWriter;
import java.util.Arrays;
import java.util.stream.Collectors;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.jsoup.helper.StringUtil;

import com.arcare.document.docx.service.WordAutoGeneraterService;
import com.arcare.document.docx.vo.ReturnObject;
import com.arcare.document.docx.wrap.Log;
import com.google.gson.Gson;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class ReportApiServlet extends HttpServlet {

	private static final long serialVersionUID = 1L;
	/**
	 * bind service
	 */
	private WordAutoGeneraterService wordAutoGeneraterService;
	
	public ReportApiServlet(WordAutoGeneraterService wordAutoGeneraterService) {
		this.wordAutoGeneraterService=wordAutoGeneraterService;
	}

	@Override
	protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
		Log.log("POST start");
		String s=req.getReader().lines().collect(Collectors.joining(System.lineSeparator()));
		JsonElement json=new JsonParser().parse(s);
		String result=new Gson().toJson(Arrays.asList("SUCCESS"));
		ReturnObject resultObj=new ReturnObject();
		if(json.isJsonObject()) {

			JsonObject jsonObj=json.getAsJsonObject();

			//v1
			String applyNumber="";
			if(jsonObj.get("ApplyNumber")!=null) {
				applyNumber=jsonObj.get("ApplyNumber").getAsString();
			}
			String reporCategory="";
			if(jsonObj.get("ReporCategory")!=null) {
				reporCategory = jsonObj.get("ReporCategory").getAsString();
			}

			//v2
			String caseNo="";
			if(jsonObj.get("CaseNo")!=null) {
				caseNo = jsonObj.get("CaseNo").getAsString();
			}
			String revision="";
			if(jsonObj.get("Revision")!=null) {
				revision = jsonObj.get("Revision").getAsString();
			}
			String corpDBName="";
			if(jsonObj.get("CorpDBName")!=null) {
				corpDBName = jsonObj.get("CorpDBName").getAsString();
			}
			String fileName="";
			if(jsonObj.get("FileName")!=null) {
				fileName = jsonObj.get("FileName").getAsString();
			}
			
			//v1
			if(!StringUtil.isBlank(applyNumber) &&
			   !StringUtil.isBlank(reporCategory) &&
			   !StringUtil.isBlank(revision)) {

				Log.log("ApplyNumber:"+applyNumber);
				Log.log("ReporCategory:"+reporCategory);
				Log.log("Revision:"+revision);

				String r=this.wordAutoGeneraterService.generateReport(applyNumber,reporCategory,revision);
				
				if(r!=null) {
					if(!r.startsWith("ERROR:")) {
						Log.log("SUCCESS");
						result=new Gson().toJson(Arrays.asList("SUCCESS"));
					}else {
						Log.log(r);
						result=new Gson().toJson(Arrays.asList(r));
					}
				}else {
					String resultMsg="ERROR:null";
					Log.log(resultMsg);
					result=new Gson().toJson(Arrays.asList(resultMsg));
				}
			//v2
			}else if(!StringUtil.isBlank(caseNo) || 
					 !StringUtil.isBlank(revision)||
					 !StringUtil.isBlank(corpDBName)||
					 !StringUtil.isBlank(fileName)){
				
				Log.log("CaseNo:"+caseNo);
				Log.log("Revision:"+revision);
				Log.log("CorpDBName:"+corpDBName);
				Log.log("FileName:"+fileName);
				String r=this.wordAutoGeneraterService.generateReportV2(caseNo,revision,corpDBName,fileName);
				if(r!=null) {
					if(!r.startsWith("ERROR:")) {
						Log.log("SUCCESS");
						resultObj.setResult(true);
						result=new Gson().toJson(resultObj);
					}else {
						Log.log(r);
						resultObj.setResult(false);
						resultObj.setErrorMessage(r);
						result=new Gson().toJson(resultObj);
					}
				}else {
					String resultMsg="ERROR:null";
					Log.log(resultMsg);
					resultObj.setResult(false);
					resultObj.setErrorMessage(resultMsg);
					result=new Gson().toJson(resultObj);
				}
			}else {
				String resultMsg=String.format("ERROR: params empty ApplyNumber:%s , ReporCategory:%s , Revision:%s, CaseNo %s, CorpDBName:%s ",applyNumber,reporCategory,revision,caseNo,corpDBName);
				Log.log(resultMsg);
				resultObj.setResult(false);
				resultObj.setErrorMessage(resultMsg);
				result=new Gson().toJson(resultObj);
			}
		}else {
			String resultMsg="ERROR:not json object";
			Log.log(resultMsg);
			resultObj.setResult(false);
			resultObj.setErrorMessage(resultMsg);
			result=new Gson().toJson(resultObj);
		}
		PrintWriter out = resp.getWriter();
		resp.setContentType("application/json");
		resp.setCharacterEncoding("UTF-8");
		out.print(result);
		out.flush();
		Log.log("POST end");
	}
}
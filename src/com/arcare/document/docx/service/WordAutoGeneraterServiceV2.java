package com.arcare.document.docx.service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Collectors;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.arcare.document.docx.core.v2.WordMultiColumnProviderV2;
import com.arcare.document.docx.core.v2.WordMultiHeaderBodyHProviderV2;
import com.arcare.document.docx.core.v2.WordMultiHeaderBodyVProviderV2;
import com.arcare.document.docx.core.v2.WordReplaceProviderV2;
import com.arcare.document.docx.dao.BaseDAO;
import com.arcare.document.docx.dao.CommonDAOV2;
import com.arcare.document.docx.util.FileNameUtil;
import com.arcare.document.docx.vo.CaseVO;
import com.arcare.document.docx.vo.DefineVO;
import com.arcare.document.docx.wrap.CSVUtil;
import com.arcare.document.docx.wrap.CommonUtil;
import com.arcare.document.docx.wrap.Log;
import com.arcare.document.docx.wrap.WordDocUtil;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class WordAutoGeneraterServiceV2 implements WordAutoGeneraterService{

	/**
	 * bind dao
	 */
	private CommonDAOV2 commonDAOV2;
	/**
	 * bind config
	 */
	private Map<String,String> config;
	/**
	 * 是否需要刪除檔案
	 * @return
	 */
	private boolean isNeedClean() {
		return Boolean.valueOf(this.config.get("delete.tempfile.v2"));
	}

	private WordAutoGeneraterService wordAutoGeneraterServiceV1;

	public WordAutoGeneraterServiceV2(BaseDAO commonDAO,Map<String,String> config,WordAutoGeneraterService wordAutoGeneraterServiceV1) {
		this.commonDAOV2=(CommonDAOV2) commonDAO;
		this.config=config;
		this.wordAutoGeneraterServiceV1=wordAutoGeneraterServiceV1;
	}
	/**
	 * 向下相容
	 */
	@Override
	public String generateReport(String applyNumber, String reporCategory, String revision) {
		return this.wordAutoGeneraterServiceV1.generateReport(applyNumber, reporCategory, revision);
	}
	
	
	/**
	 * query photo object in stream
	 * @return
	 */
	private Set<DefineVO> getPhotoSet(Map<String, List<DefineVO>> data) {
		Set<DefineVO> allPhotos=new HashSet<>();
		data.forEach((k1,v1)->{
			List<DefineVO> photos=v1.stream()
				.filter(obj->obj.getDataFields().stream().filter(deepObj->"photo".equals(deepObj.getFieldType())).count()>0)
				.collect(Collectors.toList());
			if(photos.size()>0) {
				allPhotos.addAll(photos);
			}
		});
		return allPhotos;
	}
	
	/**
	 * download Image from database
	 * @param applyNumber
	 * @param photos
	 * @param rootdir
	 * @return
	 */
	private String downloadImage(String caseNo,String corpDBName, Set<DefineVO> photos,String rootdir) {
		return commonDAOV2.pullImages(caseNo, corpDBName, photos,rootdir);
	}
	
	/**
	 * v3
	 */
	@Override
	public String generateReportV2(String caseNo, String revision,String corpDBName,String FileName) {
		final StringBuffer resultMsg = new StringBuffer();
		try {
			String fileName=FileName;
			Log.log("generateReport :"+fileName);
			final String rootDir = this.config.get("agent.path.rootdir.v2");//"./output";
			List<CaseVO> configAndTemplateList=this.commonDAOV2.pullConfig(caseNo, rootDir, corpDBName);
			
			//process type 2
			//key = 書籤 , value = 檔案
			Map<String,String> sectionMapping=new TreeMap<>();
			configAndTemplateList.stream()
				.filter(caseVo->caseVo.getType()==2)
				.forEach(caseVo->{
					try {
						Log.log(caseVo.getBookMark()+" "+caseVo.getConfigFilePath());
						if(caseVo.getConfigFilePath() == null) {
							//不用做data bind
							XWPFDocument docx = WordDocUtil.readDocx(caseVo.getTemplateFilePath());
							String resultFile = FileNameUtil.generateSectionFileName(config);
							WordDocUtil.savDocx(docx, resultFile);
							sectionMapping.put(caseVo.getBookMark(), resultFile);
						}else {
							//要做data bind
							Map<String, List<DefineVO>> data = CSVUtil.generateDefineMap(caseVo.getConfigFilePath());
							Set<DefineVO> photoSet = this.getPhotoSet(data);
							String imageDir = this.downloadImage(caseNo, corpDBName, photoSet, rootDir);
							Map<String, List<DefineVO>> result = this.commonDAOV2.queryBindDataMap(caseNo,data,corpDBName);						
							String output = this.generateSection(
									caseVo.getTemplateFilePath(), // input file
									result.get("MultiColumn"), // image case
									result.get("MultiHeaderV"), // headerV
									result.get("MultiBodyV"), // bodyV
									result.get("MultiRow"), // multiRow table
									result.get("SingleRow"), // single row value
									result.get("MultiHeaderH"), //headerH
									result.get("MultiBodyH"), //bodyH
									imageDir, //img dir
									result.get("Header"),
									result.get("Footer"));
							sectionMapping.put(caseVo.getBookMark(), output);
							//清除圖檔
							if(this.isNeedClean()) {
								Log.log("delete imageDir:"+imageDir);
								CommonUtil.cleanOldTemp(imageDir);
							}
						}
					} catch (Exception e) {
						Log.error(e);
						resultMsg.append("ERROR:"+e.getMessage()+"\n");
					}
			});
	
			StringBuffer base=new StringBuffer();
			//process type 1
			configAndTemplateList.stream()
			.filter(vo->vo.getType()==1)
			.forEach(vo->{
				try {
					Map<String, List<DefineVO>> data;
					data = CSVUtil.generateDefineMap(vo.getConfigFilePath());
					Set<DefineVO> photoSet=this.getPhotoSet(data);
					String imageDir=this.downloadImage(caseNo, corpDBName, photoSet, rootDir);
					Map<String, List<DefineVO>> result = this.commonDAOV2.queryBindDataMap(caseNo, data,corpDBName);
					String output=this.generateSection(
							vo.getTemplateFilePath(), // input file
							result.get("MultiColumn"), // image case
							result.get("MultiHeaderV"), // headerV
							result.get("MultiBodyV"), // bodyV
							result.get("MultiRow"), // multiRow table
							result.get("SingleRow"), // single row value
							result.get("MultiHeaderH"), //headerH
							result.get("MultiBodyH"), //bodyH
							imageDir, //img dir
							result.get("Header"),
							result.get("Footer"));
					base.append(output);
					//清除圖檔
					if(this.isNeedClean()) {
						Log.log("delete imageDir:"+imageDir);
						CommonUtil.cleanOldTemp(imageDir);
					}
				} catch (Exception e) {
					resultMsg.append("ERROR:"+e.getMessage()+"\n");
				}
			});
			
			Log.log("base:"+base.toString());
			//讀取底稿
			XWPFDocument baseDocument = WordDocUtil.readDocx(base.toString());
			//合併章節
			WordDocUtil.mergeSectionToOneDoc(baseDocument, sectionMapping, resultMsg);
			//底稿concat 檔名
			String abstractFile=FileNameUtil.generateConcatFileName(config);
			//存成草稿
			WordDocUtil.savDocx(baseDocument, abstractFile);
			//concat後草稿
			XWPFDocument concat = WordDocUtil.readDocx(abstractFile);
			//移除書籤
			List<String> bookmarkNeedDelete=new ArrayList<>();
			sectionMapping.forEach((bookmark,filepath)->{
				bookmarkNeedDelete.add(bookmark);
			});
			WordDocUtil.removeStringByBookMarks(concat, bookmarkNeedDelete);
			//重建目錄
			WordDocUtil.refreshIndex(concat, "REPORT_INDEX");
			//另存新檔
			//最終文件
			String resultFile=FileNameUtil.generateResultFileName(this.config);
			WordDocUtil.savDocx(concat, resultFile);
			//上船結果檔案
			String updateResultMsg=this.commonDAOV2.updateResultToDbV2(caseNo, revision, fileName, resultFile);
			if(updateResultMsg.startsWith("ERROR:")) {
				resultMsg.append(updateResultMsg+"\n");
			}
			
			this.removeTemplateFiles(sectionMapping, base, configAndTemplateList, abstractFile);

		}catch(Exception e) {
			Log.error(e);
			resultMsg.append("ERROR:"+e.getMessage()+"\n");
		}
		if("".equals(resultMsg.toString())) {
			return "SUCCESS";
		}else {
			return resultMsg.toString();
		}
	}
	/**
	 * 刪除temp檔
	 * @param sectionMapping
	 * @param base
	 * @param configAndTemplateList
	 * @param abstractFile
	 */
	private void removeTemplateFiles(
			Map<String,String> sectionMapping,
			StringBuffer base,
			List<CaseVO> configAndTemplateList,
			String abstractFile) {
		//移除範本檔案
		if(this.isNeedClean()) {
			sectionMapping.forEach((bookmark,filepath)->{
				Log.log("delete section file:"+filepath);
				new File(filepath).delete();
			});
			Log.log("delete base file"+base.toString());
			new File(base.toString()).delete();
			Log.log("delete concat file:"+new File(abstractFile));
			new File(abstractFile).delete();
			configAndTemplateList.stream().findFirst().ifPresent(configVO->{
				if(configVO.getTemplateFilePath()!=null) {
					Log.log("delete template folder:"+new File(configVO.getTemplateFilePath()).getParent());
					CommonUtil.cleanOldTemp(new File(configVO.getTemplateFilePath()).getParent());
				}
			});
		}
	}
	/**
	 * generae report detail
	 * @param sourceFilePath
	 * @param multiColumn
	 * @param multiHeaderV
	 * @param multiBodyV
	 * @param multiRow
	 * @param singleRow
	 * @param multiHeaderH
	 * @param multiBodyH
	 * @param imgDir
	 * @param headers
	 * @param footers
	 * @return
	 * @throws IOException
	 */
	private String generateSection(String sourceFilePath, 
			List<DefineVO> multiColumn,
			List<DefineVO> multiHeaderV, List<DefineVO> multiBodyV, 
			List<DefineVO> multiRow,
			List<DefineVO> singleRow,
			List<DefineVO> multiHeaderH,List<DefineVO> multiBodyH,
			String imgDir,
			List<DefineVO> headers,List<DefineVO> footers) {
		try {
			XWPFDocument docx = WordDocUtil.readDocx(sourceFilePath);
//			process muticolumn
			WordMultiColumnProviderV2.processMultiColumn(docx, multiColumn, imgDir);
//			//process table v
			WordMultiHeaderBodyVProviderV2.processAllHeaderBodyTableV(docx, multiHeaderV, multiBodyV);
//			process table h
			WordMultiHeaderBodyHProviderV2.processAllHeaderBodyTableH(docx, multiHeaderH, multiBodyH);
//			process multirow table
			WordReplaceProviderV2.insertMultiRowToTABLE(docx, multiRow);
//			process header
			WordDocUtil.replaceHeader(docx, headers);
//			process footer
			WordDocUtil.replaceFooter(docx, footers);
//			process body single value and image
			WordDocUtil.replaceBody(docx, singleRow, imgDir);
//			save file
			String resultFile = FileNameUtil.generateSectionFileName(this.config);
			WordDocUtil.savDocx(docx, resultFile);
			return resultFile;
		} catch (Exception e) {
			Log.error(e);
			return "ERROR:"+e.getMessage();
		}
	}
}

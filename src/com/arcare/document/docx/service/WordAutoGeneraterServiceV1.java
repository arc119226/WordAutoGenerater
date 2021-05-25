package com.arcare.document.docx.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.arcare.document.docx.core.v2.WordMultiColumnProviderV2;
import com.arcare.document.docx.core.v2.WordMultiHeaderBodyHProviderV2;
import com.arcare.document.docx.core.v2.WordMultiHeaderBodyVProviderV2;
import com.arcare.document.docx.core.v2.WordReplaceProviderV2;
import com.arcare.document.docx.dao.BaseDAO;
import com.arcare.document.docx.dao.CommonDAOV1;
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
public class WordAutoGeneraterServiceV1 implements WordAutoGeneraterService{

	/**
	 * bind dao
	 */
	private CommonDAOV1 commonDAO;
	/**
	 * bind config
	 */
	private Map<String,String> config;
	
	public WordAutoGeneraterServiceV1(BaseDAO commonDAO,Map<String,String> config) {
		this.commonDAO=(CommonDAOV1) commonDAO;
		this.config=config;
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
	private String generate(String sourceFilePath, 
			List<DefineVO> multiColumn,
			List<DefineVO> multiHeaderV, List<DefineVO> multiBodyV, 
			List<DefineVO> multiRow,
			List<DefineVO> singleRow,
			List<DefineVO> multiHeaderH,List<DefineVO> multiBodyH,
			String imgDir,List<DefineVO> headers,List<DefineVO> footers) throws IOException {

		FileOutputStream out = null;
		XWPFDocument docx=null;
		try {
			docx = WordDocUtil.readDocx(sourceFilePath);
			
			//process muticolumn
			if(multiColumn!=null) {
				WordMultiColumnProviderV2.processMultiColumn(docx, multiColumn, imgDir);
			}
			//process table v
			if(multiHeaderV!=null && multiBodyV!=null) {
				WordMultiHeaderBodyVProviderV2.processAllHeaderBodyTableV(docx, multiHeaderV, multiBodyV);
			}
			//process table h
			if(multiHeaderH!=null && multiBodyH!=null) {
				WordMultiHeaderBodyHProviderV2.processAllHeaderBodyTableH(docx, multiHeaderH, multiBodyH);
			}
			//process multirow table
			if(multiRow!=null) {
				WordReplaceProviderV2.insertMultiRowToTABLE(docx, multiRow);
			}
			//process header
			if(headers!=null) {
				WordDocUtil.replaceHeader(docx, headers);
			}
			//process footer
			if(footers!=null) {
				WordDocUtil.replaceFooter(docx, footers);
			}
			//process body single value
			if(singleRow!=null) {
				WordDocUtil.replaceBody(docx, singleRow);
			}
			//refresh index
			WordDocUtil.refreshIndex(docx, "REPORT_INDEX");

			//save file
			File outputDir = new File(this.config.get("agent.path.rootdir"));
			outputDir.mkdirs();
			SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
			String currentDate = sdf.format(new Date());
			String resultFile=this.config.get("agent.path.rootdir") + File.separator + currentDate + ".docx";
			out = new FileOutputStream(resultFile);
			docx.write(out);
			return resultFile;
		} catch (Exception e) {
			Log.error(e);
			return "ERROR:"+e.getMessage();
		} finally {
			try {
				if(docx!=null) {
					docx.close();
				}
				if (out != null) {
					out.close();
				}
			} catch (IOException e) {
				Log.error(e);
			}
		}
	}
	
	/**
	 * query config in database and check exists
	 * @param reporCategory
	 * @param revision
	 * @param rootdir
	 * @return
	 * @throws Exception
	 */
	private Map<String,String> generateConfigTemplate(String reporCategory,String revision,String rootdir) throws Exception{
		Map<String,String> confgTempalte=commonDAO.pullConfig(reporCategory,revision, rootdir);
		if(confgTempalte==null || confgTempalte.isEmpty()) {
			throw new Exception("ERROR:can't found define file");
		}
		if(confgTempalte.get("config")==null) {
			throw new Exception("ERROR:cant't found config.csv");
		}
		if(confgTempalte.get("template")==null) {
			throw new Exception("ERROR:cant't found template.docx");
		}
		return confgTempalte;
	}
	/**
	 * query photo object in stream
	 * @return
	 */
	private Set<DefineVO> getPhotoSet(Map<String, List<DefineVO>> data) {
		Set<DefineVO> allPhotos=new HashSet<>();
		data.forEach((k1,v1)->{
			List<DefineVO> photos=v1.stream()
				.filter(obj->
					obj.getDataFields().stream().filter(deepObj->"photo".equals(deepObj.getFieldType())).count()>0)
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
	private String downloadImage(String applyNumber, Set<DefineVO> photos,String rootdir) {
		return commonDAO.pullImages(applyNumber, photos, rootdir);
	}
	
	@Override
	public String generateReport(String applyNumber, String reporCategory, String revision) {
		try {
			final String rootDir=this.config.get("agent.path.rootdir");//"./output";

			Map<String,String> confgTempalte = this.generateConfigTemplate(reporCategory,revision,rootDir);
			
			Map<String,List<DefineVO>> data = CSVUtil.generateDefineMap(confgTempalte.get("config"));
			
			Set<DefineVO> allPhotos = this.getPhotoSet(data);
			
			String imgDir = this.downloadImage(applyNumber, allPhotos, rootDir);
			
			Map<String,List<DefineVO>> dataMap = commonDAO.queryBindDataMap(applyNumber, data);

			String output=this.generate(
					confgTempalte.get("template"), // input file
					dataMap.get("MultiColumn"), // image case
					dataMap.get("MultiHeaderV"), // headerV
					dataMap.get("MultiBodyV"), // bodyV
					dataMap.get("MultiRow"), // multiRow table
					dataMap.get("SingleRow"), // single row value
					dataMap.get("MultiHeaderH"), //headerH
					dataMap.get("MultiBodyH"), //bodyH
					imgDir, //img dir
					dataMap.get("Header"),dataMap.get("Footer")
			);
			
			if(!output.startsWith("ERROR")) {
				//申請書編號+"_"+版本.docx
				//upload output file
				String fileName=applyNumber+"_"+revision+".docx";
				Log.log("upload "+fileName);
				//最後上傳至DB
				String result=this.commonDAO.updateResultToDb(applyNumber,revision,fileName,output);
				Log.log("upload result:"+result);
			}
			
			if(Boolean.valueOf(config.get("delete.tempfile"))) {
				confgTempalte.forEach((key,filePath)->{
					Log.log(key+" = "+filePath);
					CommonUtil.cleanOldTemp(filePath);
				});
				Optional<String> keyOpt=confgTempalte.keySet().stream().findFirst();
				if(keyOpt.isPresent()) {
					String path = confgTempalte.get(keyOpt.get());
					File parent=new File(path).getParentFile();
					parent.delete();
				}
				CommonUtil.cleanOldTemp(imgDir);
//				new File(output).delete();
			}
			return output;
		} catch (Exception e) {
			Log.error(e);
			return e.getMessage();
		}
	}
	@Override
	public String generateReportV2(String caseNo, String revision, String corpDBName,String FileName) {
		// TODO Auto-generated method stub
		return null;
	}


}

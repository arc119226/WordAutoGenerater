package com.arcare.document.docx.dao;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Collectors;

import javax.sql.DataSource;

import com.arcare.document.docx.vo.DefineVO;
import com.arcare.document.docx.wrap.HashUtil;
import com.arcare.document.docx.wrap.Log;
import com.google.gson.JsonObject;

/**
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class CommonDAOV1 extends BaseDAO{


	public CommonDAOV1(DataSource dataSource) {
		super(dataSource);
	}


	/**
	 * 依據SQL 下載設定檔及範本檔
	 * @param sql
	 * @param targetDDir
	 * @return
	 */
	public Map<String,String> pullConfigAndTemplateData(String sql,String targetDDir){
		Map<String,String> result=new TreeMap<>();
		Connection con=null;
		ResultSet rs=null;
		try {
			con=this.dataSource.getConnection();
			Log.log(sql);
			PreparedStatement pst=con.prepareStatement(sql);
			rs=pst.executeQuery();
			while(rs.next()){
				String fileName = rs.getString(3);
				byte[] fileBytes = rs.getBytes(4);
				OutputStream targetFile=null;
				try {
					if(fileName!=null) {
						fileName=fileName.trim();
					}
					targetFile=new FileOutputStream(targetDDir+File.separator+"config.csv");
					targetFile.write(fileBytes);
					targetFile.close();
					result.put("config", targetDDir+File.separator+"config.csv");
				}catch(Exception e) {
					Log.error(e);
				}finally {
					if(targetFile!=null) {
						targetFile.close();
					}
				}
                
				String _fileName = rs.getString(1);
				byte[] _fileBytes = rs.getBytes(2);
				
				OutputStream _targetFile=null;
				try {
					if(_fileName!=null) {
						_fileName=_fileName.trim();
					}
					_targetFile=new FileOutputStream(targetDDir+File.separator+"template.docx");
					_targetFile.write(_fileBytes);
					_targetFile.close();
					result.put("template", targetDDir+File.separator+"template.docx");
				}catch(Exception e) {
					Log.error(e);
				}finally {
					if(_targetFile!=null) {
						_targetFile.close();
					}
				}
			}
		} catch (Exception e) {
			Log.error(e);
		}finally{
			try {
				rs.close();
				con.close();
			} catch (SQLException e) {
				Log.error(e);
			}
		}
		return result;
	}
	/**
	 * 依據SQL下載圖檔
	 * @param sql
	 * @param targetDir
	 */
	public void pullImageData(String sql,String targetDir){
		Connection con=null;
		ResultSet rs=null;
		try {
			con=this.dataSource.getConnection();
			Log.log(sql);
			PreparedStatement pst=con.prepareStatement(sql);
			rs=pst.executeQuery();
			while(rs.next()){
				String fileName = rs.getString(1);
				byte[] fileBytes = rs.getBytes(2);
				OutputStream targetFile=null;
				try {
					if(fileName!=null) {
						fileName=fileName.trim();
					}
					targetFile=new FileOutputStream(targetDir+File.separator+fileName);
					targetFile.write(fileBytes);
					targetFile.close();
				}catch(Exception e) {
					Log.error(e);
				}finally {
					if(targetFile!=null) {
						targetFile.close();
					}
				}
			}
		} catch (Exception e) {
			Log.error(e);
		}finally{
			try {
				if(rs!=null) {
					rs.close();
				}
				if(con!=null) {
					con.close();
				}
			} catch (Exception e) {
				Log.error(e);
			}
		}
	}

	/**
	 * 依據定義查詢資料 塞入VO
	 * @param applyNumber
	 * @param data
	 * @return
	 */
	public Map<String,List<DefineVO>> queryBindDataMap(String applyNumber, Map<String,List<DefineVO>> data){
		/**
		 * type
		 *   varName data
		 */
		final Map<String,Map<String,String>> result=new TreeMap<>();
		
		//prepare first layer result set
		data.forEach((k1,v1)->{
			result.put(k1, new TreeMap<>());
				v1.forEach(type->{
					if(type.getViewName()==null || "".equals(type.getViewName().trim())) {
						//ignore
					}else {
						Optional<JsonObject> json=this.query("select * from "+type.getViewName()+" where APPLYNUMBER='"+applyNumber+"' ");
						
						if(json.isPresent() && json.get()!=null && json.get().get("result") !=null) {
							json.get().get("result").getAsJsonArray().forEach(it->{
								if(it.isJsonObject()) {
									it.getAsJsonObject().entrySet().forEach(entity->{
									
										List<DefineVO> dvos=type.getDataFields().stream()
											.filter(d->entity.getKey().equalsIgnoreCase(d.getFieldName()))
											.collect(Collectors.toList());

										if(dvos.size()>0) {
											for(int i=0;i<dvos.size();i++) {
												
												if(result.get(k1).get(dvos.get(i).getVarName())==null && entity.getValue().getAsString().length()>=10) {
													if(dvos.get(i).getFieldType().equals("date")) {//process date
														dvos.get(i).setData(entity.getValue().getAsString().substring(0,10).trim()+" ");
//														result.get(k1).put(dvos.get(i).getVarName(),entity.getValue().getAsString().substring(0,10).trim()+" ");
														result.get(k1).put(dvos.get(i).getVarName()," ");
													}else {
														dvos.get(i).setData(entity.getValue().getAsString().trim()+" ");
//														result.get(k1).put(dvos.get(i).getVarName(),entity.getValue().getAsString().trim()+" ");
														result.get(k1).put(dvos.get(i).getVarName()," ");
													}
												}else {
													if(dvos.get(i).getFieldType().equals("date") && entity.getValue().getAsString().length()>=10) {//process date
//														result.get(k1)
//														.put(dvos.get(i).getVarName(),
//															 result.get(k1).get(dvos.get(i).getVarName())+HashUtil.getSplitString()+entity.getValue().getAsString().substring(0, 10).trim()+" "
//															);
														result.get(k1).put(dvos.get(i).getVarName()," "+HashUtil.getSplitString()+" ");
														dvos.get(i).setData(entity.getValue().getAsString().substring(0,10).trim()+" ");
													}else {
//														result.get(k1)
//															.put(dvos.get(i).getVarName(),
//																 result.get(k1).get(dvos.get(i).getVarName())+HashUtil.getSplitString()+entity.getValue().getAsString().trim()+" "
//																);
														result.get(k1).put(dvos.get(i).getVarName()," "+HashUtil.getSplitString()+" ");
														dvos.get(i).setData(entity.getValue().getAsString().trim()+" ");
													}
												}

											}
										}
									});
								}
							});		
						}
					}
				});

		});
		
		//1 level
		data.forEach((processType,types)->{
			Log.log(processType);
			//2 level
			types.forEach(vo->{
				Log.log("\t"+vo.getVarName()+","+vo.getDataFields().size());
				//3 level
				vo.getDataFields().forEach(field->{
					Log.log("\t\t"+field.getVarName()+","+field.getDatas().size());
				});
			});
		});

		return data;
	}

	/**
	 * 下載定義檔與範本檔
	 * @param reporCategory
	 * @param rootdir
	 * @return
	 */
	public Map<String,String> pullConfig(String reporCategory,String revision,String rootdir) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		String currentDate = sdf.format(new Date());
		String configDir=rootdir+File.separator+currentDate+"_config";
		new File(configDir).mkdirs();
		Map<String,String> result=this.pullConfigAndTemplateData(String.format(getSql("./resource/pullTemplateAndDefine.sql"), reporCategory), configDir);
		return result;
	}
	
	/**
	 * 下載圖片
	 * @param applyNumber
	 * @param photos
	 * @param rootdir
	 * @return
	 */
	public String pullImages(String applyNumber, Set<DefineVO> photos,String rootdir) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		String currentDate = sdf.format(new Date());
		String imgDir=rootdir+File.separator+currentDate+"_img";
		new File(imgDir).mkdirs();
		photos.forEach(photo->{
			String sqlTemplate=getSql("./resource/pullImages.sql");
			String sql=String.format(sqlTemplate, photo.getViewName(),applyNumber);
			this.pullImageData(sql,imgDir);
		});
		return imgDir;
	}
	/**
	 * 上傳資料
	 * @param applyNumber
	 * @param fileName
	 * @param sourceFile
	 */
	public String updateResultToDb(String applyNumber,String revision,String fileName,String sourceFile) {
		String sql=getSql("./resource/updateResult.sql");
		Connection con=null;
		PreparedStatement pst=null;
		try {
			con=this.dataSource.getConnection();
			Log.log(sql);
			pst=con.prepareStatement(sql);
			pst.setString(1, fileName);
			pst.setBytes(2,Files.readAllBytes(Paths.get(sourceFile)));
			pst.setString(3, applyNumber);
			pst.setString(4, revision);
			pst.executeUpdate();
			return "SUCCESS";
		} catch (Exception e) {
			Log.error(e);
			return "ERROR:"+e.getMessage();
		}finally{
			try {
				if(pst != null) {
					pst.close();
				}
				if(con != null) {
					con.close();
				}
			} catch (SQLException e) {
				Log.error(e);
			}
		}
	}

}

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
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Random;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Collectors;

import javax.sql.DataSource;

import com.arcare.document.docx.vo.CaseVO;
import com.arcare.document.docx.vo.DefineVO;
import com.arcare.document.docx.wrap.HashUtil;
import com.arcare.document.docx.wrap.Log;
import com.google.gson.JsonObject;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class CommonDAOV2 extends BaseDAO{

	public CommonDAOV2(DataSource dataSource) {
		super(dataSource);
	}

	/**
	 * 依據SQL 下載設定檔及範本檔
	 * @param sql
	 * @param targetDDir
	 * @return
	 */
	public List<CaseVO> pullConfigAndTemplateData(String sql,String targetDDir){
		List<CaseVO> result=new ArrayList<>();
		Connection con=null;
		ResultSet rs=null;
		try {
			con=this.dataSource.getConnection();
			Log.log(sql);
			PreparedStatement pst=con.prepareStatement(sql);
			rs=pst.executeQuery();
			while(rs.next()){
				CaseVO vo=new CaseVO();
				String caseNo =rs.getString(1);
				vo.setCaseNo(caseNo.trim());
				Integer type = rs.getInt(4);
				vo.setType(type);
				String bookMark = rs.getString(5);
				vo.setBookMark(bookMark.trim());
				
				if(type == 1) {
					//底稿csv
					{//scope start
						String fileName = "base";
						byte[] fileBytes = rs.getBytes(2);
						OutputStream targetFile=null;
						try {
							if(fileName!=null) {
								fileName=fileName.trim();
							}
							targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".csv");
							targetFile.write(fileBytes);
							targetFile.close();
							vo.setConfigFilePath(targetDDir+File.separator+fileName+".csv");
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
					//底稿docx
					{//scope start
						String fileName = "base";
						byte[] fileBytes = rs.getBytes(3);
						OutputStream targetFile=null;
						try {
							if(fileName!=null) {
								fileName=fileName.trim();
							}
							targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".docx");
							targetFile.write(fileBytes);
							targetFile.close();
							vo.setTemplateFilePath(targetDDir+File.separator+fileName+".docx");
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
				}else if(type==2) {
					//章節
					{//scope start
						String fileName = rs.getString(5).trim();
						byte[] fileBytes = rs.getBytes(2);
						OutputStream targetFile=null;
						try {
							if(rs.getBytes(2)!=null) {
								if(fileName!=null) {
									fileName=fileName.trim();
								}
								targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".csv");
								targetFile.write(fileBytes);
								targetFile.close();
								vo.setConfigFilePath(targetDDir+File.separator+fileName+".csv");
							}
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
					{//scope start
						String fileName = rs.getString(5).trim();
						byte[] fileBytes = rs.getBytes(3);
						OutputStream targetFile=null;
						try {
							if(fileName!=null) {
								fileName=fileName.trim();
							}
							targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".docx");
							targetFile.write(fileBytes);
							targetFile.close();
							vo.setTemplateFilePath(targetDDir+File.separator+fileName+".docx");
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
				}else if(type == 5) {
					//excel
					{//scope start
						String fileName = rs.getString(5).trim();
						byte[] fileBytes = rs.getBytes(2);
						OutputStream targetFile=null;
						try {
							if(rs.getBytes(2)!=null) {
								if(fileName!=null) {
									fileName=fileName.trim();
								}
								targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".csv");
								targetFile.write(fileBytes);
								targetFile.close();
								vo.setConfigFilePath(targetDDir+File.separator+fileName+".csv");
							}
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
					{//scope start
						String fileName = rs.getString(5).trim();
						byte[] fileBytes = rs.getBytes(3);
						OutputStream targetFile=null;
						try {
							if(fileName!=null) {
								fileName=fileName.trim();
							}
							targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".xlsx");
							targetFile.write(fileBytes);
							targetFile.close();
							vo.setTemplateFilePath(targetDDir+File.separator+fileName+".xlsx");
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
				}else if(type == 6) {
					//word
					{//scope start
						//csv
						String fileName = rs.getString(5).trim();
						byte[] fileBytes = rs.getBytes(2);
						OutputStream targetFile=null;
						try {
							if(rs.getBytes(2)!=null) {
								if(fileName!=null) {
									fileName=fileName.trim();
								}
								targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".csv");
								targetFile.write(fileBytes);
								targetFile.close();
								vo.setConfigFilePath(targetDDir+File.separator+fileName+".csv");
							}
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
					{//scope start
						String fileName = rs.getString(5).trim();
						byte[] fileBytes = rs.getBytes(3);
						OutputStream targetFile=null;
						try {
							if(fileName!=null) {
								fileName=fileName.trim();
							}
							targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".docx");
							targetFile.write(fileBytes);
							targetFile.close();
							vo.setTemplateFilePath(targetDDir+File.separator+fileName+".docx");
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
				}else if(type == 7) {
					//jpg
					{//scope start
						String fileName = rs.getString(5).trim();
						byte[] fileBytes = rs.getBytes(2);
						OutputStream targetFile=null;
						try {
							if(rs.getBytes(2)!=null) {
								if(fileName!=null) {
									fileName=fileName.trim();
								}
								targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".csv");
								targetFile.write(fileBytes);
								targetFile.close();
								vo.setConfigFilePath(targetDDir+File.separator+fileName+".csv");
							}
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
					{//scope start
						String fileName = rs.getString(5).trim();
						byte[] fileBytes = rs.getBytes(3);
						OutputStream targetFile=null;
						try {
							if(fileName!=null) {
								fileName=fileName.trim();
							}
							targetFile=new FileOutputStream(targetDDir+File.separator+fileName+".jpg");
							targetFile.write(fileBytes);
							targetFile.close();
							vo.setTemplateFilePath(targetDDir+File.separator+fileName+".jpg");
						}catch(Exception e) {
							Log.error(e);
						}finally {
							if(targetFile!=null) {
								targetFile.close();
							}
						}
					}//scope end
				}
				
				result.add(vo);
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
	 * 依據定義查詢資料 塞入VO
	 * @param caseNo
	 * @param data
	 * @return
	 * @throws Exception 
	 */
	public Map<String,List<DefineVO>> queryBindDataMap(String caseNo, Map<String,List<DefineVO>> data,String corpDBName) throws Exception{
		/**
		 * type
		 *   varName data
		 */
		final Map<String,Map<String,String>> result=new TreeMap<>();
		
		try {
			//prepare first layer result set
			data.forEach((k1,v1)->{
				result.put(k1, new TreeMap<>());
					v1.forEach(type->{
						if(type.getViewName()==null || "".equals(type.getViewName().trim())) {
							//ignore
						}else {
							Optional<JsonObject> json=this.query("select * from "+corpDBName+".."+type.getViewName()+" where CASENO='"+caseNo+"' ");
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
//															result.get(k1).put(dvos.get(i).getVarName(),entity.getValue().getAsString().substring(0,10).trim()+" ");
															result.get(k1).put(dvos.get(i).getVarName()," ");
														}else {
															dvos.get(i).setData(entity.getValue().getAsString().trim()+" ");
//															result.get(k1).put(dvos.get(i).getVarName(),entity.getValue().getAsString().trim()+" ");
															result.get(k1).put(dvos.get(i).getVarName()," ");
														}
													}else {
														if(dvos.get(i).getFieldType().equals("date") && entity.getValue().getAsString().length()>=10) {//process date
//															result.get(k1)
//															.put(dvos.get(i).getVarName(),
//																 result.get(k1).get(dvos.get(i).getVarName())+HashUtil.getSplitString()+entity.getValue().getAsString().substring(0, 10).trim()+" "
//																);
															result.get(k1).put(dvos.get(i).getVarName()," "+HashUtil.getSplitString()+" ");
															dvos.get(i).setData(entity.getValue().getAsString().substring(0,10).trim()+" ");
														}else {
//															result.get(k1)
//																.put(dvos.get(i).getVarName(),
//																	 result.get(k1).get(dvos.get(i).getVarName())+HashUtil.getSplitString()+entity.getValue().getAsString().trim()+" "
//																	);
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
		}catch(Exception e) {
			String s=Log.error(e);
			throw new Exception(s);
		}
		return data;
	}

	
	
	/**
	 * 下載定義檔與範本檔
	 * @param reporCategory
	 * @param rootdir
	 * @return
	 */
	public List<CaseVO> pullConfig(String caseNo,String rootdir,String corpDBName) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		String currentDate = sdf.format(new Date());
		String configDir=rootdir+File.separator+currentDate+"_config_"+Math.abs(new Random().nextInt());
		new File(configDir).mkdirs();
		List<CaseVO> result=this.pullConfigAndTemplateData(String.format(getSql("./resource/pullTemplateAndDefineV2.sql"),corpDBName, caseNo), configDir);
		return result;
	}
	
	/**
	 * 上傳資料
	 * @param applyNumber
	 * @param fileName
	 * @param sourceFile
	 */
	public String updateResultToDbV2(String caseNo,String revision,String fileName,String sourceFile) {
		String sql=getSql("./resource/updateResultV2.sql");
		Connection con=null;
		PreparedStatement pst=null;
		try {
			con=this.dataSource.getConnection();
			Log.log(sql);
			pst=con.prepareStatement(sql);
			pst.setString(1, fileName);
			pst.setBytes(2,Files.readAllBytes(Paths.get(sourceFile)));
			pst.setString(3, caseNo);
			pst.setString(4, revision);
			int r=pst.executeUpdate();
			if(r>0) {
				return "SUCCESS";
			}else {
				return String.format("ERROR:update result fail caseNo %s , revision %s", caseNo,revision);
			}
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
	/**
	 * 下載圖片
	 * @param caseNo
	 * @param corpDBName
	 * @param photos
	 * @param rootdir
	 * @return
	 */
	public String pullImages(String caseNo, String corpDBName, Set<DefineVO> photos, String rootdir) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		String currentDate = sdf.format(new Date());
		String imgDir=rootdir+File.separator+currentDate+"_"+Math.abs(new Random().nextInt())+"_img";
		new File(imgDir).mkdirs();
		photos.forEach(level1->{
			level1.getDataFields().forEach(level2->{
				if(level2.getFieldType().equals("photo")) {
					String sql=level2.getDownloadImageSql(corpDBName, caseNo);
					this.pullImageData(sql,imgDir);
				}
			});
		});
		return imgDir;
	}
	/**
	 * 下載圖片 依據bookmark命名
	 * @param sql
	 * @param targetDir
	 */
	private void pullImageData(String sql, String targetDir) {
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

}

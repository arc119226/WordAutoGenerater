package com.arcare.document.docx.wrap;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Queue;
import java.util.stream.Collectors;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;

import com.arcare.document.docx.vo.DefineVO;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class WordDocUtil {
	/**
	 * 移除書籤上的範本文字
	 * @param doc
	 * @param bookMarks
	 */
    public static void removeStringByBookMarks(XWPFDocument doc,List<String> bookMarks) {
    	bookMarks.forEach(bookMark->{
            doc.getParagraphs().stream()
            .filter(p->p.getCTP().toString().matches(String.format("[.\\S\\s]*<w:bookmarkStart w:id=\"\\d+\" w:name=\"%s\"\\s?\\/>[.\\S\\s]*", bookMark)))
            .forEach(p->{
            	p.getRuns().forEach(r->{
            		r.setText("",0);//清空文字
            	});	
            });
    	});
    }
   
	/**
	 * 依據書籤定位在段落中增加圖片
	 * @param docx
	 * @param imgFilePath
	 * @param bookMark
	 * @param widthCm
	 * @param heightCm
	 */
	public static void addPictureByBookMarkInParagraphs(XWPFDocument docx,String imgFilePath,String bookMark,double widthCm,double heightCm) {
		docx.getParagraphs().stream().filter(paragraph->paragraph.getCTP().toString().contains("\""+bookMark+"\"")).forEach(paragraph->{
			Log.log("addPictureByBookMark "+bookMark);
			File image = new File(imgFilePath.trim());
			double width = WordDocUtil.cmToP(widthCm);
			double height = WordDocUtil.cmToP(heightCm);
			int imgFormat = WordDocUtil.getImageFormat(image.getName());

			FileInputStream inputStream=null;
			try {
				inputStream=new FileInputStream(image);
				paragraph.createRun().addPicture(inputStream,imgFormat,image.getName(),Units.toEMU(width),Units.toEMU(height));
			}catch(Exception e1){
				Log.error(e1);
			}finally {
				if(inputStream!=null) {
					try {
						inputStream.close();
					} catch (IOException e) {
						Log.error(e);
					}
				}
			}
			Log.log("addPictureByBookMark end "+bookMark);
		});
	}
	
	
	 /**
	  * 
	  * @param imgFileName
	  * @return
	  */
	public static int getImageFormat(String imgFileName) {
		int format;
		if (imgFileName.toLowerCase().endsWith(".emf")) {
			format = XWPFDocument.PICTURE_TYPE_EMF;
		}else if (imgFileName.toLowerCase().endsWith(".wmf")) {
			format = XWPFDocument.PICTURE_TYPE_WMF;
		}else if (imgFileName.toLowerCase().endsWith(".pict")) {
			format = XWPFDocument.PICTURE_TYPE_PICT;
		}else if (imgFileName.toLowerCase().endsWith(".jpeg") || imgFileName.toLowerCase().endsWith(".jpg")) {
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		}else if (imgFileName.toLowerCase().endsWith(".png")) {
			format = XWPFDocument.PICTURE_TYPE_PNG;
		}else if (imgFileName.toLowerCase().endsWith(".dib")) {
			format = XWPFDocument.PICTURE_TYPE_DIB;
		}else if (imgFileName.toLowerCase().endsWith(".gif")) {
			format = XWPFDocument.PICTURE_TYPE_GIF;
		}else if (imgFileName.toLowerCase().endsWith(".tiff")) {
			format = XWPFDocument.PICTURE_TYPE_TIFF;
		}else if (imgFileName.toLowerCase().endsWith(".eps")) {
			format = XWPFDocument.PICTURE_TYPE_EPS;
		}else if (imgFileName.toLowerCase().endsWith(".bmp")) {
			format = XWPFDocument.PICTURE_TYPE_BMP;
		}else if (imgFileName.toLowerCase().endsWith(".wpg")) {
			format = XWPFDocument.PICTURE_TYPE_WPG;
		}else {
			return 0;
		}
		return format;
	}
	/**
	 * scale 
	 * @param cm
	 * @return
	 */
	public static double cmToP(double cm) {
		double unit=28.346457;
		return unit*cm;
	}
	
	/**
	 * 合併word 
	 * @param src
	 * @param append
	 * @throws Exception
	 */
    public static void appendBody(XWPFDocument src, XWPFDocument append) throws Exception {
        CTBody src1Body = src.getDocument().getBody();
        CTBody src2Body = append.getDocument().getBody();
        List<XWPFPictureData> allPictures = append.getAllPictures();
        Map<String,String> map = new HashMap<>();
        for (XWPFPictureData picture : allPictures) {
            String before = append.getRelationId(picture);
            String after = src.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
            map.put(before, after);
        }
        appendBody(src1Body, src2Body,map);
    }
    /**
     * 
     * @param src
     * @param append
     * @param map
     * @throws Exception
     */
    private static void appendBody(CTBody src, CTBody append,Map<String,String> map) throws Exception {  
        XmlOptions optionsOuter = new XmlOptions();  
        optionsOuter.setSaveOuter();  
        String appendString = append.xmlText(optionsOuter);  
        String srcString = src.xmlText();  
        String prefix = srcString.substring(0,srcString.indexOf(">")+1);  
        String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));  
        String sufix = srcString.substring(srcString.lastIndexOf("<"));  
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));  
        if (map != null && !map.isEmpty()) {  
            for (Map.Entry<String, String> set : map.entrySet()) {  
                addPart = addPart.replace(set.getKey(), set.getValue());  
            }  
        }  
        CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);  
        src.set(makeBody);  
    }  
	
	/**
	 * footer填值
	 * @param docx
	 * @param footers
	 */
	public static void replaceFooter(XWPFDocument docx,List<DefineVO> footers) {
		Log.log("replaceFooter start");
		if(footers==null) {
			Log.log("replaceFooter end");
			return;
		}
		footers.forEach(prefix->{
			prefix.getDataFields().forEach(d->{
				final String bookMark=d.getVarName();
				d.getDatasList().forEach(data->{
					docx.getFooterList().forEach(h->{
						h.getParagraphs().forEach(p->{
							if(p.getCTP().toString().contains("\""+bookMark+"\"")) {
								String newData=p.getParagraphText().replace(bookMark, data);
								p.getRuns().forEach(r->{
									r.setText("",0);
								});
								p.getRuns().stream().findFirst().ifPresent(r->{
									r.setText(newData, 0);
									Log.log("found key in footer:"+bookMark+", replace footer key to ->"+newData);
								});
							}
						});
					});
				});
			});
		});
		Log.log("replaceFooter end");
	}
	/**
	 * body填值 及 插入圖片
	 * @param docx
	 * @param singleRowV2
	 * @param imageDir
	 */
	public static void replaceBody(XWPFDocument docx,List<DefineVO> singleRowV2,String imageDir) {
		Log.log("replaceBody start");
		if(singleRowV2==null) {
			Log.log("replaceBody end");
			return;
		}
		//add table paragraph
		final List<XWPFParagraph> allP=docx.getTables().stream()
			.flatMap(tab->tab.getRows().stream())
			.flatMap(row->row.getTableCells().stream())
			.flatMap(cel->cel.getParagraphs().stream()).collect(Collectors.toList());
		//add body paragraph
		allP.addAll(docx.getParagraphs());
		Log.log("body paragraph size:"+allP.size());
		singleRowV2.stream().flatMap(prefix->prefix.getDataFields().stream()).forEach(d->{
			d.getDatasList().forEach(data->{
				allP.stream().filter(_p->_p.getCTP().toString().contains("\""+d.getVarName()+"\"")).forEach(p->{
					String newData=p.getParagraphText().replace(d.getVarName(), data);
					p.getRuns().forEach(pr->{
						pr.setText("",0);
					});
					p.getRuns().stream().findFirst().ifPresent(pr->{
						if("photo".equals(d.getFieldType())) {

							String imgFilePath= new File(imageDir,d.getVarName()+".jpg").getAbsolutePath();
							double widthCm=Double.parseDouble(d.getWidth());
							double heightCm=Double.parseDouble(d.getHeight());

							File image = new File(imgFilePath.trim());
							double width = WordDocUtil.cmToP(widthCm);
							double height = WordDocUtil.cmToP(heightCm);
							int imgFormat = WordDocUtil.getImageFormat(image.getName());

							FileInputStream inputStream=null;
							try {
								inputStream=new FileInputStream(image);
								pr.addPicture(inputStream,imgFormat,image.getName(),Units.toEMU(width),Units.toEMU(height));
							}catch(Exception e1){
								Log.error(e1);
							}finally {
								if(inputStream!=null) {
									try {
										inputStream.close();
									} catch (IOException e) {
										Log.error(e);
									}
								}
							}
							Log.log("found key in para:"+d.getVarName()+", replace key to a picture->"+imgFilePath);
						}else {
							pr.setText(newData,0);
							Log.log("found key in para:"+d.getVarName()+", replace key to ->"+newData);
						}
					});
				});
			});
		});
		Log.log("replaceBody end");
	}
	/**
	 * table 填值
	 * @param table
	 * @param bookmark
	 * @param data
	 */
	public static void replaceTable(XWPFTable table,String bookmark,Queue<String> data) {
		table.getRows().stream()
		.flatMap(row->row.getTableCells().stream())
		.flatMap(cel->cel.getParagraphs().stream()).collect(Collectors.toList()).forEach(p->{
			if(p.getParagraphText().trim().equals(bookmark)) {
				String newData=p.getParagraphText().replace(bookmark, data.poll());
				p.getRuns().forEach(pr->{
					pr.setText("",0);
				});
				p.getRuns().stream().findFirst().ifPresent(pr->{
					pr.setText(newData,0);
					Log.log("found key in para:"+bookmark+", replace key to ->"+newData);
				});
			}
		});
		
//		final List<XWPFParagraph> allP=table.getRows().stream()
//		.flatMap(row->row.getTableCells().stream())
//		.flatMap(cel->cel.getParagraphs().stream()).collect(Collectors.toList());
//		allP.stream().filter(_p->_p.getCTP().toString().contains("\""+bookmark+"\"")).forEach(p->{
//			String newData=p.getParagraphText().replace(bookmark, data.poll());
//			p.getRuns().forEach(pr->{
//				pr.setText("",0);
//			});
//			p.getRuns().stream().findFirst().ifPresent(pr->{
//				pr.setText(newData,0);
//				Log.log("found key in para:"+bookmark+", replace key to ->"+newData);
//			});
//		});
	}
	/**
	 * body 填值
	 * @param docx
	 * @param singleRowV2
	 */
	public static void replaceBody(XWPFDocument docx,List<DefineVO> singleRowV2) {
		Log.log("replaceBody start");
		//add table paragraph
		final List<XWPFParagraph> allP=docx.getTables().stream()
			.flatMap(tab->tab.getRows().stream())
			.flatMap(row->row.getTableCells().stream())
			.flatMap(cel->cel.getParagraphs().stream()).collect(Collectors.toList());
		//add body paragraph
		allP.addAll(docx.getParagraphs());
		Log.log("body paragraph size:"+allP.size());
		singleRowV2.stream().flatMap(prefix->prefix.getDataFields().stream()).forEach(d->{
			d.getDatasList().forEach(data->{
				allP.stream().filter(_p->_p.getCTP().toString().contains("\""+d.getVarName()+"\"")).forEach(p->{
					String newData=p.getParagraphText().replace(d.getVarName(), data);
					p.getRuns().forEach(pr->{
						pr.setText("",0);
					});
					p.getRuns().stream().findFirst().ifPresent(pr->{
						pr.setText(newData,0);
						Log.log("found key in para:"+d.getVarName()+", replace key to ->"+newData);
					});
				});
			});
		});
		Log.log("replaceBody start");
	}
	
	/**
	 * header 填值
	 * @param docx
	 * @param headers
	 */
	public static void replaceHeader(XWPFDocument docx,List<DefineVO> headers) {
		Log.log("replaceHeader start");
		if(headers==null) {
			Log.log("replaceHeader end");
			return;
		}
		headers.stream().flatMap(prefix->prefix.getDataFields().stream()).forEach(d->{
			d.getDatasList().forEach(data->{
				docx.getHeaderList().stream().flatMap(h->h.getParagraphs().stream()).forEachOrdered(p->{
					if(p.getCTP().toString().contains("\""+d.getVarName()+"\"")) {
						String newData=p.getParagraphText().replace(d.getVarName(), data);
						p.getRuns().forEach(r->{
							r.setText("",0);
						});
						p.getRuns().stream().findFirst().ifPresent(r->{
							r.setText(newData,0);
							Log.log("found key in header:"+d.getVarName()+", replace header key to ->"+newData);
						});
					}
				});
			});
		});
		Log.log("replaceHeader end");
	}
	
	/**
	 * 重建目錄
	 * @param docx
	 * @param keyword
	 */
	public static void refreshIndex(XWPFDocument docx,String keyword) {
		docx.getParagraphs().forEach(p->{
			//refresh index
			if(p.getCTP().toString().contains("\""+keyword+"\"")) {
				String newData=p.getParagraphText().replace(keyword, "");
				Log.log("generate index");
				p.getRuns().forEach(r->{
					r.setText("",0);
				});
				p.getRuns().stream().findFirst().ifPresent(r->{
					r.setText(newData,0);
					Log.log("found key in body paragraphs:"+keyword+", replace key to ->"+newData);
				});
				CTP ctP = p.getCTP();
				CTSimpleField toc = ctP.addNewFldSimple();
				toc.setInstr("TOC \\h");
				toc.setDirty(STOnOff.TRUE);
			}
		});
	}
	/**
	 * 複製某一行的table style
	 * @param index
	 * @param currentTable
	 */
	public static void copyTableStyleFromIndexRow(int index,XWPFTable currentTable) {
		XWPFTableRow opfirstRow=currentTable.getRows().get(index);
		if(opfirstRow!=null) {
			XWPFTableRow firstRow =opfirstRow;
			Optional<XWPFTableCell> opcell=firstRow.getTableCells().stream().findFirst();
			Map<Integer,Object> _indexAlignment=new HashMap<>();
			if(opcell.isPresent()) {
				XWPFTableCell cell=opcell.get();
				//copy cell style
				final CTTcPr tcpr=(CTTcPr) cell.getCTTc().getTcPr().copy();
				if(tcpr.getShd()!=null) {
					CTShd shd=(CTShd) tcpr.getShd().copy();
//					shd.setFill(null);
					tcpr.setShd(shd);
				}

				for(int i=0;i<currentTable.getRows().get(index).getTableCells().size();i++) {
					Optional<XWPFParagraph> p=currentTable.getRows().get(index).getTableCells().get(i).getParagraphs().stream().findFirst();
					if(p.isPresent()) {
						XWPFParagraph _p=p.get();
						_indexAlignment.put(i, _p.getAlignment());
					}
				}

				for(int i=index;i<currentTable.getRows().size();i++) {
					XWPFTableRow r = currentTable.getRows().get(i);
					for(int j=0;j<r.getTableCells().size();j++) {
						XWPFTableCell c=r.getTableCells().get(j);
						c.getCTTc().setTcPr(tcpr);
						XWPFParagraph cellp=cell.addParagraph();
						if(_indexAlignment.get(j)!=null) {
							cellp.setAlignment(ParagraphAlignment.class.cast(_indexAlignment.get(j)));
						}
					}
		
				}
			}
		}
	}
	/**
	 * 複製table的style
	 * @param currentTable
	 * @param cellSize
	 * @param skipFillColor
	 */
	public static void copyTableStyleFull(XWPFTable currentTable,final int cellSize,final boolean skipFillColor) {
		Map<Integer,CTTcPr> styleMap=new HashMap<>();
		currentTable.getRows().stream().filter(r->r.getTableCells().size()==cellSize).findFirst().ifPresent(firstRow->{
			for(int i=0;i<firstRow.getTableCells().size();i++) {
				XWPFTableCell cell=firstRow.getTableCells().get(i);
				if(cell.getCTTc()!=null && cell.getCTTc().getTcPr()!=null) {
					final CTTcPr tcpr=(CTTcPr) cell.getCTTc().getTcPr().copy();
					if(tcpr.getShd()!=null) {
						CTShd shd=(CTShd) tcpr.getShd().copy();
						if(skipFillColor) {
							shd.setFill(null);
						}
						tcpr.setShd(shd);
					}
					styleMap.put(i, tcpr);
				}
			}
		});
		List<XWPFTableRow> dataRows=currentTable.getRows().stream().filter(r->r.getTableCells().size()==cellSize).collect(Collectors.toList());
		for(int i=1;i<dataRows.size();i++) {
			XWPFTableRow r = dataRows.get(i);
			for(int j=0;j<r.getTableCells().size();j++) {
				XWPFTableCell cell=r.getTableCells().get(j);
				cell.getCTTc().setTcPr(styleMap.get(j));
			}
		}
	}
	/**
	 * 
	 * @param currentTable
	 * @param cellSize
	 */
	public static void copyTableStyleFromFirstTarget(XWPFTable currentTable,final int cellSize) {
		currentTable.getRows().stream().filter(r->r.getTableCells().size()==cellSize).findFirst().ifPresent(firstRow->{
			firstRow.getTableCells().stream().findFirst().ifPresent(cell->{
				//copy cell style
				if(cell.getCTTc()!=null && cell.getCTTc().getTcPr()!=null) {
					final CTTcPr tcpr=(CTTcPr) cell.getCTTc().getTcPr().copy();
					if(tcpr.getShd()!=null) {
						CTShd shd=(CTShd) tcpr.getShd().copy();
						shd.setFill(null);
						tcpr.setShd(shd);
					}
					List<XWPFTableRow> dataRows=currentTable.getRows().stream().filter(r->r.getTableCells().size()==cellSize).collect(Collectors.toList());
					for(int i=1;i<dataRows.size();i++) {
						XWPFTableRow r = dataRows.get(i);
						r.getTableCells().forEach(c->{
							c.getCTTc().setTcPr(tcpr);
						});
					}
				}
			});
		});
	}
	
	/**
	 *開啟文檔
	 * @param filePath
	 * @return
	 */
	public static XWPFDocument readDocx(String filePath){
		try {
			Log.log("readDocx:"+filePath);
			Path path = Paths.get(filePath);
			byte[] byteData = Files.readAllBytes(path);
			XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(byteData));
			return doc;
		} catch (IOException e) {
			Log.error(e);
		}
		return null;
	}
	
	/**
	 * word 另存新檔
	 * @param input
	 * @param wordFilePath
	 */
	public static void savDocx(XWPFDocument input,String wordFilePath) {
		 FileOutputStream outputStream=null;
		try {
            //save file
			outputStream = new FileOutputStream(wordFilePath);
            input.write(outputStream);
			outputStream.flush();
			input.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }finally {
        	if(outputStream!=null) {
        		 try {
        			 outputStream.close();
				} catch (IOException e) {
					Log.error(e);
				}
        	}
        }
	}
	
	/**
	 *複製段落
	 * @param clone
	 * @param source
	 */
	public static void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
		CTP ctp=CTP.Factory.newInstance();
		ctp.set(source.getCTP());
		clone.getCTP().set(ctp);
	}
	/**
	 * 給定書籤 找到文黨中的段落
	 * @param docx
	 * @param headerBookmark
	 * @return
	 */
	public static Optional<XWPFParagraph> findHeaderParagraphByBookMark(XWPFDocument docx,String headerBookmark){
		for(XWPFParagraph p:docx.getParagraphs()){
			for(XWPFRun r:p.getRuns()){
				if(r.getCTR().toString().contains(headerBookmark)){
					return Optional.of(p);
				}
			}
		}
		return Optional.empty();
	}
	/**
	 *  依據來源 複製段落
	 * @param source
	 * @param target
	 */
	public static void copyParagraphDeep(XWPFParagraph source, XWPFParagraph target) {
		CTPPr _ctppr=CTPPr.Factory.newInstance();
		_ctppr.set(source.getCTP().getPPr().copy());
	    target.getCTP().setPPr(_ctppr);
	    for (int i=0; i<source.getRuns().size(); i++ ) {
	        XWPFRun sourceRun = source.getRuns().get(i);
	        XWPFRun targetRun = target.createRun();
	        CTR _ctr=CTR.Factory.newInstance();
	        _ctr.set(sourceRun.getCTR().copy());
	        targetRun.getCTR().set(_ctr);
	    }
	}
	/**
	 *  給定cursor座標 複製段落 並更改標題
	 * @param p
	 * @param docx
	 * @param cursor
	 * @return
	 */
	public static XWPFParagraph copyParagraphToCurserAndUpdateText(XWPFParagraph p,XWPFDocument docx,XmlCursor cursor,String text){
		cursor.toEndToken();
		while(cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START);
		XWPFParagraph newParagraph=docx.insertNewParagraph(cursor);
		copyParagraphDeep(p,newParagraph);
		newParagraph.getRuns().get(0).setText(text, 0);
		return newParagraph;
	}
	/**
	 *  給定書籤找到文黨中的table
	 * @param docx
	 * @param tableBookmark
	 * @return
	 */
	public static Optional<XWPFTable> findBodyTableByBookMark(XWPFDocument docx,String tableBookmark){
		//找template table
		for(XWPFTable table:docx.getTables()){
			if(table.getCTTbl().toString().contains(tableBookmark)){
				return Optional.of(table);
			}
		}
		return Optional.empty();
	}
	/**
	 *  給定cursor 複製來源table
	 * @param template
	 * @param docx
	 * @param cursor
	 */
	public static XWPFTable copyTable(XWPFTable template,XWPFDocument docx,XmlCursor cursor){
		cursor.toEndToken();
		while(cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START);
		XWPFTable targetTable=docx.insertNewTbl(cursor);
		targetTable.setWidth(template.getWidth());
		copyTableDeep(template,targetTable);
		return targetTable;
	}
	/**
	 * 複製table
	 * @param source
	 * @param target
	 */
	private static void copyTableDeep(XWPFTable source, XWPFTable target) {
		CTTblPr cttblpr=CTTblPr.Factory.newInstance();
		cttblpr.set(source.getCTTbl().getTblPr().copy());
		CTTblGrid cttblgrid=CTTblGrid.Factory.newInstance();
		cttblgrid.set(source.getCTTbl().getTblGrid().copy());
		
	    target.getCTTbl().setTblPr(cttblpr);
	    target.getCTTbl().setTblGrid(cttblgrid);
	    
	    for (int r = 0; r<source.getRows().size(); r++) {
	        XWPFTableRow targetRow = target.createRow();
	        XWPFTableRow sourceRow = source.getRows().get(r);
	        
	        CTTrPr _cttrPr=CTTrPr.Factory.newInstance();
	        _cttrPr.set(sourceRow.getCtRow().getTrPr().copy());
	        targetRow.getCtRow().setTrPr(_cttrPr);
	        
	        for (int c=0; c<sourceRow.getTableCells().size(); c++) {
	            XWPFTableCell targetCell = c==0 ? targetRow.getTableCells().get(0) : targetRow.createCell();
	            XWPFTableCell sourceCell = sourceRow.getTableCells().get(c);
	            CTTcPr _cttcpr=CTTcPr.Factory.newInstance();
	            _cttcpr.set(sourceCell.getCTTc().getTcPr().copy());
	            targetCell.getCTTc().setTcPr(_cttcpr); 
	            XmlCursor cursor = targetCell.getParagraphArray(0).getCTP().newCursor();
	            for (int p = 0; p < sourceCell.getBodyElements().size(); p++) {
	                IBodyElement elem = sourceCell.getBodyElements().get(p);
	                if (elem instanceof XWPFParagraph) {
	                    XWPFParagraph targetPar = targetCell.insertNewParagraph(cursor);
	                    cursor.toNextToken();
	                    XWPFParagraph par = (XWPFParagraph) elem;
	                    WordDocUtil.copyParagraphDeep(par, targetPar);
	                } else if (elem instanceof XWPFTable) {
	                    XWPFTable targetTable = targetCell.insertNewTbl(cursor);
	                    XWPFTable table = (XWPFTable) elem;
	                    copyTableDeep(table, targetTable);
	                    cursor.toNextToken();
	                }
	            }
	            targetCell.removeParagraph(targetCell.getParagraphs().size()-1);
	        }
	    }
	    target.removeRow(0);
	}
	/**
	 * 合併章節
	 * @param baseDocument
	 * @param sectionMapping
	 * @param resultMsg
	 */
	public static void mergeSectionToOneDoc(XWPFDocument baseDocument,Map<String,String> sectionMapping,StringBuffer resultMsg) {
		sectionMapping.forEach((bookmark,filePath)->{
			Log.log("section:"+bookmark+" -> "+filePath);
			try {
				WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx(filePath), bookmark);
			} catch (Exception e) {
				Log.error(e);
				resultMsg.append("ERROR:"+e.getMessage()+","+bookmark+","+filePath+"\n");
			}
		});
	}
	/**
	 * 合併水平表格
	 * @param table
	 * @param row
	 * @param startCell
	 * @param endCell
	 */
	public static void mergeCellInRow(XWPFTable table, int row, int startCell, int endCell) {  
        for (int cellIndex = startCell; cellIndex <= endCell; cellIndex++) {  
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);  
            if ( cellIndex == startCell ) {  
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);  
            } else {  
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);  
            }  
        }  
    }  
	/**
	 *  合併垂直表格
	 * @param table
	 * @param col
	 * @param startRow
	 * @param endRow
	 */
	public static void mergeCellsInColumn(XWPFTable table, int col, int startRow, int endRow) {  
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {  
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);  
            if ( rowIndex == startRow ) {   
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);  
            } else {   
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);  
            }  
        }  
    }

	/**
	 * (x1,y1) row
	 * --------------------
	 * |                  |
	 * |                  |column
	 * |                  |
	 * --------------------
	 *                  (x2,y2)
	 * from (x1,y1)left top point to (x2,y2) right button
	 * @param table
	 * @param row
	 * @param column
	 * @param endRowCell
	 * @param endColumnCell
	 */
	public static void mergCelllAndRow(XWPFTable table,int row,int column,int endRowCell,int endColumnCell) {
		for(int t=row;t<=endRowCell;t++) {
//			//合併第x row 第 start 到 end 
//			String script=String.format("WordDocUtil.mergeCellInRow(table, %s, %s, %s);", x,row,endColumnCell);
//			System.out.println("合併第"+x+" row 第 "+column+" column 起向右合併到第 "+endColumnCell+" cell ");
//			System.out.println(script);
			WordDocUtil.mergeCellInRow(table, t, column, endColumnCell);
		}
		for(int t=column;t<=endColumnCell;t++) {
//			String script=String.format("WordDocUtil.mergeCellsInColumn(table, %s, %s, %s);", t,column,endRowCell);
//			System.out.println("合併第"+t+" column 第 "+row+" row起 向下合併到第"+endRowCell+" cell ");
//			System.out.println(script);
	    	WordDocUtil.mergeCellsInColumn(table, t, row, endRowCell);
		}
	}
}

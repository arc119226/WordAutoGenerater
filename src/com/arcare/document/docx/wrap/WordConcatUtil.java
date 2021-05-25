package com.arcare.document.docx.wrap;

import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.select.Elements;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

/**
 * 合併word 限制: 不包括嵌入式物件
 * 
 * @author FUHSIANG_LIU
 *
 */
public class WordConcatUtil {
	/**
	 * 重新生成Footer
	 * 
	 * @param result
	 * @param footers
	 */
	public static void reGenerateFooter(XWPFDocument result, List<XWPFFooter> footers) {
		CTSectPr sectPr = result.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(result, sectPr);
		footers.forEach(footer -> {
			List<XWPFPictureData> allPictures = footer.getAllPictures();
			Map<String, String> footerImgMap = new HashMap<>();
			for (XWPFPictureData picture : allPictures) {
				String before = footer.getRelationId(picture);
				String after;
				try {
					after = footer.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
					footerImgMap.put(before, after);
				} catch (InvalidFormatException e) {
					Log.error(e);
					return;
				}
			}
			XWPFParagraph[] parsFooter = new XWPFParagraph[footer.getParagraphs().size()];
			for (int i = 0; i < footer.getParagraphs().size(); i++) {
				XWPFParagraph xwpf = footer.getParagraphs().get(i);
				String xml = xwpf.getCTP().xmlText();
				if (footerImgMap != null && !footerImgMap.isEmpty()) {
					for (Map.Entry<String, String> set : footerImgMap.entrySet()) {
						xml = xml.replace(set.getKey(), set.getValue());
					}
				}
				CTP ctp;
				try {
					ctp = CTP.Factory.parse(xml);
					xwpf.getCTP().copy().set(ctp);
					parsFooter[i] = xwpf;
				} catch (XmlException e) {
					Log.error(e);
					return;
				}
			}
			XWPFFooter newFooter = policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);

			for (XWPFParagraph p : parsFooter) {
				if (!"".equals(p.getParagraphText().trim()) || p.getCTP().xmlText().contains("rId")) {
					newFooter.createParagraph().getCTP().set(p.getCTP().copy());
				}
			}

			footer.getAllPictures().forEach(picture -> {
				try {
					newFooter.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
				} catch (InvalidFormatException e) {
					Log.error(e);
					return;
				}
			});
		});
	}

	/**
	 * 重新生成Header
	 * 
	 * @param result
	 * @param headers
	 */
	public static void reGenerateHeader(XWPFDocument result, List<XWPFHeader> headers) {
		CTSectPr sectPr = result.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(result, sectPr);
		headers.forEach(header -> {
			List<XWPFPictureData> allPictures = header.getAllPictures();
			Map<String, String> headerImgMap = new HashMap<>();
			for (XWPFPictureData picture : allPictures) {
				String before = header.getRelationId(picture);
				String after;
				try {
					after = header.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
					headerImgMap.put(before, after);
				} catch (InvalidFormatException e) {
					Log.error(e);
					return;
				}
			}
			XWPFParagraph[] parsHeader = new XWPFParagraph[header.getParagraphs().size()];
			for (int i = 0; i < header.getParagraphs().size(); i++) {
				XWPFParagraph xwpf = header.getParagraphs().get(i);
				String xml = xwpf.getCTP().xmlText();
				if (headerImgMap != null && !headerImgMap.isEmpty()) {
					for (Map.Entry<String, String> set : headerImgMap.entrySet()) {
						xml = xml.replace(set.getKey(), set.getValue());
					}
				}
				CTP ctp;
				try {
					ctp = CTP.Factory.parse(xml);
					xwpf.getCTP().copy().set(ctp);
					parsHeader[i] = xwpf;
				} catch (XmlException e) {
					Log.error(e);
					return;
				}
			}
			XWPFHeader newHeader = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

			for (XWPFParagraph p : parsHeader) {
				if (!"".equals(p.getParagraphText().trim()) || p.getCTP().xmlText().contains("rId")) {
					newHeader.createParagraph().getCTP().set(p.getCTP().copy());
				}
			}

			header.getAllPictures().forEach(picture -> {
				try {
					newHeader.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
				} catch (InvalidFormatException e) {
					Log.error(e);
					return;
				}
			});
		});
	}

	/**
	 * 
	 * @param base     底稿
	 * @param append   要插入的檔案
	 * @param bookmark 插入到底稿的哪個書籤
	 * @throws Exception
	 */
	public static void appendBodyByBookMark(XWPFDocument base, XWPFDocument append, String bookmark) throws Exception {
		CTBody baseBody = base.getDocument().getBody();
		CTBody appendBody = append.getDocument().getBody();
		List<XWPFPictureData> allPictures = append.getAllPictures();
		Map<String, String> imgMap = new HashMap<>();
		for (XWPFPictureData picture : allPictures) {
			String before = append.getRelationId(picture);
			String after = base.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
			imgMap.put(before, after);
		}
		
		WordConcatUtil.appendBody(baseBody, appendBody, imgMap, bookmark);
	}

	/**
	 * 
	 * @param base
	 * @param append
	 * @param bodyImgMap
	 * @param bookmark
	 * @throws Exception
	 */
	private static void appendBody(CTBody base, CTBody append, Map<String, String> bodyImgMap, String bookmark)
			throws Exception {
		XmlOptions optionsOuter = new XmlOptions();
		optionsOuter.setSaveOuter();
		String appendString = append.xmlText(optionsOuter);
		//各自section去除header及footer
		appendString = appendString.replaceAll("<w:footerReference r:id=\"rId\\d+\" w:type=\"default\"\\s?\\/>", "") //除去footer
								   .replaceAll("<w:footerReference r:id=\"rId\\d+\" w:type=\"first\"\\s?\\/>", "") //除去footer
								   .replaceAll("<w:headerReference r:id=\"rId\\d+\" w:type=\"default\"\\s?\\/>", "") //除去header 
								   .replaceAll("<w:headerReference r:id=\"rId\\d+\" w:type=\"first\"\\s?\\/>", ""); //除去header
		String srcString = base.xmlText();
		// xml頭
		String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
		// 主要xml
		String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));

		org.jsoup.nodes.Document doc = Jsoup.parse(mainPart, "", org.jsoup.parser.Parser.xmlParser());

		Set<String> addSet = new HashSet<>();
		addSet.add("#root");// only select root element

		Elements mainElements = JsoupUtil.ignoreNameSpaceSelect(doc, addSet);

		// xml尾
		String sufix = srcString.substring(srcString.lastIndexOf("<"));

		// 章節檔 body部分
		String addBody = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
		if (bodyImgMap != null && !bodyImgMap.isEmpty()) {
			for (Map.Entry<String, String> set : bodyImgMap.entrySet()) {
				addBody = addBody.replace(set.getKey(), set.getValue());
			}
		}

		int indexNeedInsert = -1;
		for (int i = 0; i < mainElements.size(); i++) {
			Element em = mainElements.get(i);
			List<Node> children = em.childNodes();
			for (int j = 0; j < children.size(); j++) {
				if (children.get(j).toString().matches(String.format(
						"[.\\S\\s]*<w:bookmarkStart w:id=\"\\d+\" w:name=\"%s\"\\s?\\/>[.\\S\\s]*", bookmark))) {
					indexNeedInsert = j;
				}
			}
		}

		if (indexNeedInsert != -1) {
			org.jsoup.nodes.Document appendDoc = Jsoup.parse(addBody, "", org.jsoup.parser.Parser.xmlParser());
			appendDoc.outputSettings().prettyPrint(false);
			Elements appendEms = JsoupUtil.ignoreNameSpaceSelect(appendDoc, addSet);
			mainElements.get(0).insertChildren(indexNeedInsert + 1, appendEms.get(0).childNodes());
		}
		try {		
			//replace something use regx
			String strMainElements=mainElements.toString()
				.replaceAll("&nbsp;","");
//				.replaceAll("<wp:posOffset>\\s?(\r\n|\n)\\s+", "<wp:posOffset>")
//				.replaceAll("(\r\n|\n)?\\s+<\\/wp:posOffset>", "</wp:posOffset>");				
//				.replaceAll("<w:t>\\s?(\r\n|\n)\\s+", "<w:t>")
//				.replaceAll("<w:t xml:space=\"preserve\">\\s?(\r\n|\n)\\s+", "<w:t xml:space=\"preserve\">")
//				.replaceAll("(\r\n|\n)?\\s+<\\/w:t>", " </w:t>")			
//				.replaceAll("<wp14:pctWidth>\\s?(\r\n|\n)\\s+", "<wp14:pctWidth>")
//				.replaceAll("(\r\n|\n)?\\s+<\\/wp14:pctWidth>", "</wp14:pctWidth>")
//				.replaceAll("<wp14:pctHeight>\\s?(\r\n|\n)\\s+", "<wp14:pctHeight>")
//				.replaceAll("(\r\n|\n)?\\s+<\\/wp14:pctHeight>", "</wp14:pctHeight>");
			CTBody makeBody = CTBody.Factory.parse(prefix + strMainElements + sufix);
			base.set(makeBody);
		}catch(Exception e) {
			Log.log(prefix + mainElements + sufix);
			throw new Exception(e);
		}
	}

	/**
	 * test case
	 * 
	 * @param args
	 * @throws Exception
	 */
	public static void main(String args[]) throws Exception {
		// step1. 合併檔案 儲存
//		XWPFDocument baseDocument = WordDocUtil.readDocx("./Base.docx");
//
//		WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx("./Doc1.docx"), "section1");
//		WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx("./Doc2.docx"), "section2");
//		WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx("./Doc3.docx"), "section3");
//		WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx("./Doc4.docx"), "section4");
//		WordConcatUtil.reGenerateHeader(baseDocument, baseDocument.getHeaderList());// 要重新產生Header才會顯示
//		WordConcatUtil.reGenerateFooter(baseDocument, baseDocument.getFooterList());// 要重新產生Footer才會顯示
//		baseDocument.write(new FileOutputStream("./atstract.docx"));
//		XWPFDocument result = WordDocUtil.readDocx("./atstract.docx");
//
//		// step2 塞資料 處理細節
//		Optional<XWPFTable> opTable = result.getTables().stream().findFirst();
//		if (opTable.isPresent()) {
//			WordDocUtil.mergCelllAndRow(opTable.get(), 0, 0, 5, 1);
//			WordDocUtil.mergCelllAndRow(opTable.get(), 6, 0, 9, 1);
//			WordDocUtil.mergCelllAndRow(opTable.get(), 2, 3, 8, 7);
//		}
//
//		WordDocUtil.removeStringByBookMarks(result, Arrays.asList("section1", "section2", "section3", "section4"));
//		result.write(new FileOutputStream("./result.docx"));

		
//		XWPFDocument baseDocument = WordDocUtil.readDocx("./output/20181018091512148.docx");

//		WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx("./output/20181018091510766.docx"), "Chapter_ConductPower");
//		WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx("./Doc2.docx"), "section2");
//		WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx("./Doc3.docx"), "section3");
//		WordConcatUtil.appendBodyByBookMark(baseDocument, WordDocUtil.readDocx("./Doc4.docx"), "section4");
//		WordConcatUtil.reGenerateHeader(baseDocument, baseDocument.getHeaderList());// 要重新產生Header才會顯示
//		WordConcatUtil.reGenerateFooter(baseDocument, baseDocument.getFooterList());// 要重新產生Footer才會顯示
//		baseDocument.write(new FileOutputStream("./output/atstract.docx"));
//		XWPFDocument result = WordDocUtil.readDocx("./output/atstract.docx");

//		// step2 塞資料 處理細節
//		Optional<XWPFTable> opTable = result.getTables().stream().findFirst();
//		if (opTable.isPresent()) {
//			WordDocUtil.mergCelllAndRow(opTable.get(), 0, 0, 5, 1);
//			WordDocUtil.mergCelllAndRow(opTable.get(), 6, 0, 9, 1);
//			WordDocUtil.mergCelllAndRow(opTable.get(), 2, 3, 8, 7);
//		}
//
//		WordDocUtil.removeStringByBookMarks(result, Arrays.asList("section1", "section2", "section3", "section4"));
//		result.write(new FileOutputStream("./result.docx"));
		
//		org.jsoup.nodes.Document d=Jsoup.parse("<wp>         asdfasdf.</wp>".replaceAll(" ", "##SP##ARC"), "", org.jsoup.parser.Parser.xmlParser());

	}
}

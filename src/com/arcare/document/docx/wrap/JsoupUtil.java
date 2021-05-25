package com.arcare.document.docx.wrap;

import java.util.Set;

import org.jsoup.nodes.Document;
import org.jsoup.nodes.Document.OutputSettings;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class JsoupUtil {
	/**
	 * find doc by tagName ignore namespace
	 * @param doc
	 * @param tagName
	 * @return
	 */
	public static Elements ignoreNameSpaceSelect(Document doc, Set<String> tagName){
		//不要幫忙排版
		OutputSettings outputSettings = new OutputSettings();
		outputSettings.prettyPrint(false);
		doc.outputSettings(outputSettings);
		Elements withTypes = new Elements();
		for( Element element : doc.select("*") ){
		    final String s[] = element.tagName().split(":");
		    if( s.length == 1 && tagName.contains(s[0]) == true ){
		        withTypes.add(element);
		    }
		}
		return withTypes;
	}
	/**
	 * find sub element by tagName ignore namespace
	 * @param doc
	 * @param tagName
	 * @return
	 */
	public static Elements ignoreNameSpaceSelect(Element doc, String tagName){
		Elements withTypes = new Elements();
		for( Element element : doc.select("*") ){
		    final String s[] = element.tagName().split(":"); 
		    if( s.length > 1 && s[1].equals(tagName) == true ){
		        withTypes.add(element);
		    }
		}
		return withTypes;
	}
}

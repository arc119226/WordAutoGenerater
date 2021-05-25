package com.arcare.document.docx.wrap;

import java.security.MessageDigest;
import java.util.Formatter;

/**
 * 
 * @author FUHSIANG_LIU
 *
 */
public class HashUtil {
	
	private static final String magic="arc";
	/**
	 * 取得資料分隔字串HASH
	 * @return
	 */
	public static String getSplitString() {
		return HashUtil.encryptSha256(magic).substring(0,10);
	}
	private static String encryptSha256(String password) {
		String sha1 = "";
		try {
			MessageDigest crypt = MessageDigest.getInstance("SHA-256");
			crypt.reset();
			crypt.update(password.getBytes("UTF-8"));
			sha1 = byteToHex(crypt.digest());
		} catch (Exception e) {
			Log.error(e);
		}
		return sha1;
	}
	@Deprecated
	private static String byteToHex(final byte[] hash) {
		Formatter formatter = new Formatter();
		for (byte b : hash) {
			formatter.format("%02x", b);
		}
		String result = formatter.toString();
		formatter.close();
		return result;
	}
}

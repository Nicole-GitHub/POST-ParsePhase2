package gss.Tools;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;

public class FileTools {
	private static final String className = FileTools.class.getName();


	// 檔案路徑 名稱
	private static String filenameTemp;

	/**
	 * 建立檔案
	 * 
	 * @param path			檔路徑
	 * @param fileName		檔名稱
	 * @param extension		副檔名
	 * @param fileContent	檔案內容
	 * @return 是否建立成功，成功則返回true
	 */
	public static boolean createFile(String path, String fileName, String extension, String fileContent) throws Exception {
		String funcName = "createFile";
		Boolean bool = false;
		File file ;
		
		try {
			file = new File(path);
			if(!file.exists()) file.mkdirs();
			
			filenameTemp = path + fileName + "." + extension;// 檔案路徑 名稱 檔案型別
			file = new File(filenameTemp);
			// 如果檔案不存在，則建立新的檔案
			if (!file.exists()) {
				file.createNewFile();
				bool = true;
//				System.out.println("success create file: " + filenameTemp);
			}
			// 建立檔案成功後，寫入內容到檔案裡
			writeFileContent(filenameTemp, fileContent, false);
		} catch (Exception ex) {
			throw new Exception(className + " " + funcName + " Error: \n" + ex);
		}
		return bool;
	}

	/**
	 * 向檔案中寫入內容並將參數直接寫入對應值
	 * 
	 * @param filePathName 檔名稱
	 * @param newstr   寫入的內容
	 * @param append   是否保留原內容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeFileContent(String filePathName, String newstr, boolean append) throws Exception {
		String funcName = "writeFileContent";
		Boolean bool = false;
		String filein = "\n" + newstr + "\n";// 新寫入的行，換行
		String temp = "";
		FileInputStream fis = null;
		InputStreamReader isr = null;
		BufferedReader br = null;
		FileOutputStream fos = null;
		PrintWriter pw = null;
//		Map<String,String> map = new HashMap<String,String>();
		
		try {
			File file = new File(filePathName);// 檔案路徑(包括檔名稱)
			// 將原檔案內容讀入輸入流
			fis = new FileInputStream(file);
			isr = new InputStreamReader(fis, "UTF8");
			br = new BufferedReader(isr);
			StringBuffer buffer = new StringBuffer();
			// 寫入檔案原有內容
			if (append) {
				while ((temp = br.readLine()) != null) {
					buffer.append(temp);
					// 行與行之間的分隔符 相當於“\n”
					buffer = buffer.append(System.getProperty("line.separator"));
//					if (temp.contains("$ARGV[")) {
//						String key = temp.substring(temp.indexOf("$") + 1, temp.indexOf(" "));
//						String value = temp.substring(temp.lastIndexOf("[") + 1, temp.lastIndexOf("]"));
//						map.put(key, value);
//					}
				}

//				int usridValue = Integer.parseInt(map.get("USRID").toString());
//				int passwdValue = Integer.parseInt(map.get("PASSWD").toString());
//				int finalValue = 0;
//				// forEach
//				for (Entry<String, String> entry : map.entrySet()) {
////					System.out.println("key:" + entry.getKey() + ",value:" + entry.getValue());
//					String key = entry.getKey();
//					if (!"USRID".equals(key) && !"PASSWD".equals(key)) {
//						int value = Integer.parseInt(entry.getValue());
//						finalValue = value;
//						if (value > usridValue)
//							finalValue--;
//						if (value > passwdValue)
//							finalValue--;
//						map.put(key, String.valueOf(finalValue));
//					}
//				}
			}
			buffer.append(filein);
			String str = buffer.toString();

//			if (append) {
//				String[] newstrArray = newstr.split(" ");
//				// forEach
//				for (Entry<String, String> entry : map.entrySet()) {
////					System.out.println("key:" + entry.getKey() + ",value:" + entry.getValue());
//					String key = entry.getKey();
//					if (!"USRID".equals(key) && !"PASSWD".equals(key)
//							&& (Integer.parseInt(entry.getValue()) + 1) < newstrArray.length) {
//						String value = newstrArray[Integer.parseInt(entry.getValue()) + 1];
//						value = value.contains("後面的參數") ? "" : value;
//						if (!StringUtils.isBlank(value))
//							str = str.replace("${" + key + "}", value);
//					}
//				}
//			}
			
			fos = new FileOutputStream(file);
			pw = new PrintWriter(fos);
			pw.write(str.toCharArray());
			pw.flush();
			bool = true;
		} catch (Exception ex) {
			throw new Exception(className + " " + funcName + " Error: \n" + ex);
		} finally {
			if (pw != null)	pw.close();
			if (fos != null) fos.close();
			if (br != null) br.close();
			if (isr != null) isr.close();
			if (fis != null) fis.close();
		}
		return bool;
	}
	
	/**
	 * 向檔案中寫入內容並另存新檔
	 * 
	 * @param filePathName 檔名稱
	 * @param newstr   寫入的內容
	 * @return
	 * @throws IOException
	 */
	public static boolean writeFileContent(String filePathName, String newFilePathName, String newstr) throws Exception {
		String funcName = "writeFileContent";
		Boolean bool = false;
		FileOutputStream fos = null;
		PrintWriter pw = null;
		try {
			File file = new File(filePathName);// 檔案路徑(包括檔名稱)
			StringBuffer buffer = new StringBuffer();
			buffer.append(newstr);
			fos = new FileOutputStream(file);
			pw = new PrintWriter(fos);
			pw.write(buffer.toString().toCharArray());
			pw.flush();
			bool = true;
		} catch (Exception ex) {
			throw new Exception(className + " " + funcName + " Error: \n" + ex);
		} finally {
			if (pw != null)	pw.close();
			if (fos != null) fos.close();
		}
		return bool;
	}

	/**
	 * 讀取檔案內容
	 */
	public static String readFileContent(String filePathName) throws Exception {
		String funcName = "readFileContent";
		String temp = "";
		FileInputStream fis = null;
		InputStreamReader isr = null;
		BufferedReader br = null;
		
		try {
			File file = new File(filePathName);// 檔案路徑(包括檔名稱)
			// 將檔案內容讀入輸入流
			fis = new FileInputStream(file);
			isr = new InputStreamReader(fis);
			br = new BufferedReader(isr);
			StringBuffer buffer = new StringBuffer();
			// 讀取檔案內容
			while((temp = br.readLine()) != null) {
				buffer.append(temp);
				// 行與行之間的分隔符 相當於“\n”
				buffer.append(System.getProperty("line.separator"));
			}
			
//			System.out.println(buffer);
			return buffer.toString();
		} catch (Exception ex) {
			throw new Exception(className + " " + funcName + " Error: \n" + ex);
		} finally {
			if (br != null) br.close();
			if (isr != null) isr.close();
			if (fis != null) fis.close();
		}
//		return "";
	}
	
	/**
	 * 刪除路徑下的所有資料夾與資料
	 * @param path
	 */
	public static void deleteFolder(String path) {
		File f = new File(path);
		
		if (!f.exists())
			f.mkdirs();
		
		FileTools.deleteFolder(f);
	}
	/**
	 * 刪除路徑下的所有資料夾與資料
	 * @param file
	 */
	private static void deleteFolder(File file) {
		for (File subFile : file.listFiles()) {
			if (subFile.isDirectory()) {
				deleteFolder(subFile);
			} else {
				subFile.delete();
			}
		}
		file.delete();
	}

	/**
	 * 複製檔案
	 * @param file
	 * @throws Exception 
	 */
	public static void copyFile(String srcDirPath, String destDirPath) throws Exception {
		try {
			FileUtils.copyDirectory(new File(srcDirPath), new File(destDirPath));
		} catch (IOException ex) {
			throw new Exception(className + " copyFile Error: \n" + ex);
		}
	}

}

package gss;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import gss.Tools.FileTools;
import gss.Tools.Tools;

public class RunParse {
	private static final String className = RunParse.class.getName();
	
	/**
	 * 取得 Excel 內容
	 * 
	 * @param tableLayoutPath
	 * @param map
	 * @throws Exception
	 */
	public static void run( String path ) throws Exception {

		try {

			Workbook workbook = Tools.getWorkbook(path+"排程整理.xlsx");
 
			Sheet sheet = workbook.getSheet("JOB STEP (單一JOB)");
//			Sheet sheet = workbook.getSheet("JOB STEP");

			Row row = null;
			for (int r = 2; r <= sheet.getLastRowNum(); r++) {

				System.out.println(r);
//				if(r == 1780) {
//					System.out.println("===================");
////					continue;
//				}
				
				row = sheet.getRow(r);
				if (row == null || !Tools.isntBlank(row.getCell(2)))
					break;

				String jobName = Tools.getCellValue(row, 2, "jobName");
				String cmd = Tools.getCellValue(row, 5, "Command");
				String iswork = Tools.getCellValue(row, 8, "作用中?");
				String params = Tools.getCellValue(row, 10, "參數說明");
				String cmdPath = Tools.getCellValue(row, 11, "指令");

//				System.out.println(params);
				if(StringUtils.isBlank(cmd) || StringUtils.isBlank(params) || StringUtils.isBlank(cmdPath)
						|| "N".equals(iswork))
					continue;
				
				Map<String, String> paramsArrayMap = new HashMap<String, String>();
				String[] cmdArray = cmd.split(" ");
				String[] paramsArray = params.split("\n");
				int i = 1;
				for(String str : paramsArray) {
					String[] paramsArray2 = str.indexOf("(") > 0
							? str.substring(4,str.indexOf("(")).split("=")
							: str.substring(4).split("=");
					if(cmdArray.length > i)
						paramsArrayMap.put(paramsArray2[0], cmdArray[i++]);
				}
				String[] cmdPathArray = cmdPath.split("%");
				for(String str : cmdPathArray) {
					if(paramsArrayMap.get(str) != null) {
						cmdPath = cmdPath.replace("%"+str+"%", paramsArrayMap.get(str));
					}
//					System.out.println(cmdPath);
				}
				
				// 實際要抓的ETL PERL程式路徑
				String etlPLPath = "C:/SVN/dw2209/COLLECTION/郵政整體資訊管理系統/現行郵政整體資訊管理系統SourceCode";
				cmdPath = cmdPath.replace("D:",etlPLPath); // win
//				cmdPath = cmdPath.replace("D:","/Users/nicole/22"); // mac
//				cmdPath = cmdPath.replace("C:\\SVN","/Users/nicole/22"); // mac 
				
				cmdPath = cmdPath.replace("\\", "/");
				CellStyle style = Tools.setStyle(workbook);
				Tools.setStringCell(style, null, row, 11, cmdPath);
//				String filePathAndParams = cmdPath.substring(cmdPath.indexOf("/Users")); // mac
				String filePathAndParams = cmdPath.substring(cmdPath.indexOf("C:/SVN")); // win
				String filePath = filePathAndParams.substring(0,filePathAndParams.indexOf(" "));

//				filePathAndParams = filePathAndParams.replace("/", ".").replace("\\", ".").replace(":", "");
//				filePathAndParams = filePathAndParams.substring(filePathAndParams.indexOf("ETL."));

				String perlName = filePathAndParams.substring(0,filePathAndParams.indexOf(" "));
				perlName = perlName.substring(perlName.lastIndexOf("/",perlName.lastIndexOf(".")-1)+1);
//				String perlNameAndParams = perlName + filePathAndParams.substring(filePathAndParams.indexOf(" "));

				String outputPath = path + "Output/" + jobName + "/";
//				String outputFileName = (r+1) + "_" + perlNameAndParams;
				String outputFileName = (r+1) + "_" + perlName; // 檔名改只到.PL就好
				outputFileName = outputFileName.replace("2>&1","");
				
				// 將excel內對應到的Perl檔複製到output下對應JobName目錄並變更副檔名為txt
				try {
					FileTools.createFile(outputPath, outputFileName, "pl",	FileTools.readFileContent(filePath));
					// 將cmd(含params)一起寫入上述檔案中
					FileTools.writeFileContent(outputPath + outputFileName + ".pl", "="+filePathAndParams, true);
				} catch (Exception e) {
					String NoSuchFile = "(No such file or directory)";
					if(e.getMessage().contains(NoSuchFile)) {
						System.out.println(r + NoSuchFile);
						continue;
					}
				}
				
			}

			Tools.output(workbook, path, "排程整理_Output");

			System.out.println("Done!");
		} catch (Exception ex) {
			throw new Exception(className + " Error: \n" + ex);
		}
	}
}

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

			Sheet sheet = workbook.getSheet("JOB STEP");

			Row row = null;
			for (int r = 2; r <= sheet.getLastRowNum(); r++) {

				System.out.println(r);
				if(r == 708) {
					System.out.println("===================");
//					continue;
				}
				
				row = sheet.getRow(r);
				if (row == null || !Tools.isntBlank(row.getCell(2)))
					break;

				String jobName = Tools.getCellValue(row, 2, "jobName");
				String cmd = Tools.getCellValue(row, 5, "參數說明");
				String params = Tools.getCellValue(row, 9, "參數說明");
				String cmdPath = Tools.getCellValue(row, 10, "指令");

//				System.out.println(params);
				if(StringUtils.isBlank(cmd) || StringUtils.isBlank(params) || StringUtils.isBlank(cmdPath))
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
				
				cmdPath = cmdPath.replace("D:","C:/SVN"); // win
//				cmdPath = cmdPath.replace("D:","/Users/nicole/22"); // mac
//				cmdPath = cmdPath.replace("C:\\SVN","/Users/nicole/22"); // mac

				cmdPath = cmdPath.replace("\\", "/");
				CellStyle style = Tools.setStyle(workbook);
				Tools.setStringCell(style, null, row, 11, cmdPath);
//				String filePathAndParams = cmdPath.substring(cmdPath.indexOf("/Users")); // mac
				String filePathAndParams = cmdPath.substring(cmdPath.indexOf("C:/SVN")); // win
				String filePath = filePathAndParams.substring(0,filePathAndParams.indexOf(" "));

				filePathAndParams = filePathAndParams.replace("/", ".").replace("\\", ".").replace(":", "");
				filePathAndParams = filePathAndParams.substring(filePathAndParams.indexOf("ETL."));
				String outputPath = path + "Output/" + jobName + "/";
				String outputFileName = r + "_" + filePathAndParams;
				outputFileName = outputFileName.replace("2>&1","");
				
				// 將excel內對應到的Perl檔複製到output下對應JobName目錄並變更副檔名為txt
				try {
					FileTools.createFile(outputPath, outputFileName, "txt",	FileTools.readFileContent(filePath));
					// 將cmd(含params)一起寫入上述檔案中
					FileTools.writeFileContent(outputPath + outputFileName + ".txt", filePathAndParams, true);
				} catch (Exception e) {
					if(e.getMessage().contains("(No such file or directory)")) {
						System.out.println(r + "(No such file or directory)");
						continue;
					}
				}
				
			}

			Tools.output(workbook, path, "排程整理2.xlsx");

			System.out.println("Done!");
		} catch (Exception ex) {
			throw new Exception(className + " Error: \n" + ex);
		}
	}
}

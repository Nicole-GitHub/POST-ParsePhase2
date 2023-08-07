package gss;

import java.io.File;
import java.util.Date;

import gss.Tools.FileTools;

public class Main {

	public static void main(String[] args) throws Exception {
		String className = "Main";
		try {
			String os = System.getProperty("os.name");

			System.out.println("=== NOW TIME: " + new Date());
			System.out.println("=== os.name: " + os);
			System.out.println("=== Parser.class.Path: "
					+ Main.class.getProtectionDomain().getCodeSource().getLocation().getPath());
			// 判斷當前執行的啟動方式是IDE還是jar
			// 若放檔的路徑中有中文時執行bat會讓中文變亂碼導致使用.isFile()會失效，故用.endsWith判斷
			String runPath = Main.class.getProtectionDomain().getCodeSource().getLocation().getPath();
			boolean isStartupFromJar = runPath.endsWith(".jar");
//		boolean isStartupFromJar = new File(Parser.class.getProtectionDomain().getCodeSource().getLocation().getPath()).isFile();
			System.out.println("=== isStartupFromJar: " + isStartupFromJar);
			String path = System.getProperty("user.dir") + File.separator; // Jar
			if (!isStartupFromJar) {// IDE
				path = os.contains("Mac") ? "/Users/nicole/Dropbox/POST/JavaTools/Phase2/" // Mac
						: "C:/Users/nicole_tsou/Dropbox/POST/JavaTools/Phase2/"; // win
//						: "C:/Users/Nicole/Dropbox/POST/POST-ParseExcel2HQL/"; // win(MSI)
			}

			FileTools.deleteFolder(path + "Output/");
			RunParse.run(path);
			
		} catch (Exception ex) {
			throw new Exception(className + " Error: \n" + ex);
		}
	}

}

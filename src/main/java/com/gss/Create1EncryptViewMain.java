package com.gss;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

public class Create1EncryptViewMain {

	public static void main(String[] args) {		

		String folderName = "", fileNamePath = "";
		String os = System.getProperty("os.name");
		System.out.println("===os.name===" + os);
		
		// 取properties的路徑
//		path = os.contains("Mac") ? "/Users/nicole/22/GitHub/" : "D:/GitHub/"; // mac : win
//		path += "ETL_MD/src/main/resources/";
		
		// 判斷當前執行的啟動方式是IDE還是jar
		boolean isStartupFromJar = new File(Create1EncryptViewMain.class.getProtectionDomain().getCodeSource().getLocation().getPath()).isFile();
		System.out.println("isStartupFromJar: " + isStartupFromJar);

		String path = System.getProperty("user.dir") + File.separator; // Jar
		if(!isStartupFromJar) {// IDE
			path = os.contains("Mac") ? "/Users/nicole/Dropbox/ETL/RUN_BAT/" // Mac
					: "C:/Users/Nicole/Dropbox/ETL/RUN_BAT/"; // win
			folderName = "1110705_002 一次加密VIEW勾稽外來人士統一證號異動檔";
			fileNamePath = "一次加密VIEW表勾稽外來人口new.xlsx";
		}

		/**
		 * 透過windows的cmd執行時需將System.in格式轉為big5才不會讓中文變亂碼
		 * 即使在cmd下chcp 65001轉成utf-8也沒用
		 * 但在eclipse執行時不能轉為big5
		 */
		try (Scanner s =  isStartupFromJar ? new Scanner(System.in, "big5") : new Scanner(System.in)) {
			System.out.println("請輸入需求單目錄名稱: ");
			folderName = "".equals(folderName) ? s.nextLine() : folderName;
			folderName += "/";
			System.out.println("請輸入一次加密VIEW表勾稽外來人口檔案名稱(含副檔名): ");
			fileNamePath = "".equals(fileNamePath) ? s.nextLine() : fileNamePath;
		}

		System.out.println("getPropertypath: " + path);
		Map<String, String> mapProp = Property.getProperties(path);
		// 切回ETL目錄
		path += "../";
		
		String work = mapProp.get("work");
		String path2folder = path + work + folderName;
		
		fileNamePath = path2folder + fileNamePath;
		String outputPath = path2folder + mapProp.get("outputPathBuild");
		
		System.out.println(""
				+ "\n,path2folder:\t\t" + path2folder
				+ "\n,fileNamePath:\t\t\t" + fileNamePath
				+ "\n,outputPath:\t\t\t" + outputPath
				+ "\n"
				);

		// User提供檔案資料分析
		List<List<Map<String, String>>> list = Parser.runParser(fileNamePath);
		WriteViewExcel.write(outputPath, list);
		// 將分析資料寫入對應的MD上傳檔
//		String selectField = REQandTableLayout.write(reqPath, outputPathMD, outputPathBuild, list, "REQ", true); // 順便產生SQL.sql與TableDefinition.sql
//		REQandTableLayout.write(tableDefPath, outputPathMD, "", list, "TableLayout", false);
//		REQandTableLayout.write(tableDefPath, outputPathMD, "", list, "ST", false);
//		DataMart.write(dataMartPath, outputPathMD, list, "資料來源與目的對照");
//		BatchImp.write(batchImpPath, outputPathMD, list, "檔案資料補充說明");
//		QcDef.write(qcDefPath, outputPathMD, list, "檢核定義");
//		RawData.write(rawDataPath, outputPathMD, list, "資料處理關聯");
//		
//		// 產生建置所需的檔案
//		ParserTableLayout.Parser(parserPath,outputPathMD, list);
//		GenSQLFile.write(outputPathBuildHP, outputPathBuildIQ, parserPath, list, selectField, "GenSQLFile"); // 產VIEW
//		GenBuildDB.write(hpPath, outputPathBuildHP, list, "HP");
//		GenBuildDB.write(iqPath, outputPathBuildIQ, list, "IQ");
//		GenBuild.write(apPath, outputPathBuild, list, folderName, "AP");
//		GenBuild.write(etlPath, outputPathBuild, list, folderName, "ETL");
//		GenSD.write(outputPathBuild, list, "JOB_SD"); // 產生Info.txt
		
		System.out.println("=== 已完成! ===");
	}

}

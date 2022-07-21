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
		
		// 判斷當前執行的啟動方式是IDE還是jar
		boolean isStartupFromJar = new File(Create1EncryptViewMain.class.getProtectionDomain().getCodeSource().getLocation().getPath()).isFile();
		System.out.println("isStartupFromJar: " + isStartupFromJar);

		String path = System.getProperty("user.dir") + File.separator; // Jar
		if(!isStartupFromJar) {// IDE
			path = os.contains("Mac") ? "/Users/nicole/Dropbox/ETL/RUN_BAT/" // Mac
					: "C:/Users/Nicole/Dropbox/ETL/RUN_BAT/"; // win
			folderName = "1110705_002 一次加密VIEW勾稽外來人士統一證號異動檔";
			fileNamePath = "一次加密VIEW表勾稽外來人口_DWM_COPD_DATA.xlsx";
		}

		/**
		 * 透過windows的cmd執行時需將System.in格式轉為big5才不會讓中文變亂碼
		 * 即使在cmd下chcp 65001轉成utf-8也沒用
		 * 但在eclipse與mac執行時不能轉為big5
		 */
		try (Scanner s =  isStartupFromJar && !os.contains("Mac") ? new Scanner(System.in, "big5") : new Scanner(System.in)) {
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

		// Build Sample
		String buildSample = path + mapProp.get("buildSample");
		String iqPath = buildSample + mapProp.get("iqPath");
		String oraclePath = buildSample + mapProp.get("oraclePath");
		// 建置用
		String outputPathBuild = path2folder + mapProp.get("outputPathBuild");
		// 資料倉儲SybaseIQ(DB)建置表單
		String outputPathBuildIQ = outputPathBuild + mapProp.get("outputPathBuildIQ");
		// 資料倉儲SybaseIQ(DB)建置表單
		String outputPathBuildOracle = outputPathBuild + mapProp.get("outputPathBuildOracle");

		System.out.println(""
				+ "\n,path2folder:\t" + path2folder
				+ "\n,fileNamePath:\t" + fileNamePath
				+ "\n,outputPathBuildIQ:\t" + outputPathBuildIQ
				+ "\n"
				);

		/**
		 * list.get(0) : Table Column List(含IsID)
		 * list.get(1) : GRANT
		 */
		List<List<Map<String, String>>> list = Parser.runParser(fileNamePath);
		WriteViewExcel.write(iqPath, outputPathBuildIQ, list);
		WriteMDSql.write(oraclePath, outputPathBuildOracle, list);
		WriteRCPT.write(outputPathBuild, list, folderName);
		
		System.out.println("=== 已完成! ===");
	}

}

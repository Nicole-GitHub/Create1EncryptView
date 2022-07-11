package com.gss;

import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteViewExcel {
	private static final String className = WriteViewExcel.class.getName();

	public static void write(String iqPath, String outputPath, List<List<Map<String, String>>> list) {

		XSSFWorkbook workbook = new XSSFWorkbook();
		CellStyle cellStyle = Tools.getStyle(workbook);
		Sheet sheet = null;
		Cell cell = null;
		Row row = null;
		String tableOldName = "", tableNewName = "", columnName = "", isID = "", sql = "", joinColName = "",
				joinNumStr = "", tableOldViewName = "", tableNewViewName = "";
		List<String> tableNewViewNameArr = new LinkedList<String>();
		List<Map<String, String>> isIDColList = new LinkedList<Map<String, String>>();
		Map<String, String> isIDColMap;

		try {
			int rowNum = 0;
			int cellNum = 0;
			int joinNum = 2;
			boolean sqlhead = true;

			// Table Column List
			for (Map<String, String> listDetail : list.get(0)) {
				cellNum = 0;
				tableNewName = listDetail.get("TableName");
				columnName = listDetail.get("ColumnName");
				isID = listDetail.get("IsID");

				// 是否為新table
				if (!tableOldName.equals(tableNewName)) {
					tableOldViewName = "V_" + tableOldName + "_REPID";
					tableNewViewName = "V_" + tableNewName + "_REPID";
					tableNewViewNameArr.add(tableNewViewName);

					// 組完最後的left join
					if (!"".equals(tableOldName))
						writeFile(sql, tableOldName, tableOldViewName, isIDColList, outputPath);

					// 開始新table的頭
					sheet = workbook.createSheet(tableNewName);
					sheet.setDefaultColumnWidth(30);
					sheet.setColumnWidth(2, 10 * 256);
					sql = "create or replace view nhiadm." + tableNewViewName + " as \nselect \n";
					sqlhead = true;
					isIDColList = new LinkedList<Map<String, String>>();
					joinNum = 2;
					rowNum = 0;
					row = sheet.createRow(rowNum++);
					Tools.setCell(Tools.getStyleRed(workbook), cell, row, cellNum, "Column");
				}
				tableOldName = tableNewName;
				row = sheet.createRow(rowNum++);
				Tools.setCell(cellStyle, cell, row, cellNum++, tableNewName);
				Tools.setCell(cellStyle, cell, row, cellNum++, columnName);
				Tools.setCell(cellStyle, cell, row, cellNum++, isID);

				// 是否為身分證欄位
				if ("Y".equals(isID)) {
					// 判斷這個id是屬於第幾個left join
					joinNumStr = "";
					for (Map<String, String> map : isIDColList) {
						if (columnName.startsWith(map.get("ColName").toString())
								|| map.get("ColName").startsWith(columnName))
							joinNumStr = map.get("JoinNum");
					}
					joinNumStr = "".equals(joinNumStr) ? String.valueOf(joinNum++) : joinNumStr;

					isIDColMap = new HashMap<String, String>();
					isIDColMap.put("ColName", columnName);
					isIDColMap.put("JoinNum", joinNumStr);
					isIDColList.add(isIDColMap);

					joinColName = columnName.contains("_JOIN") ? "_JOIN" : "";
					sql += (sqlhead ? " " : ",") + "CASE WHEN T" + joinNumStr + ".NEW_ID" + joinColName
							+ " IS NOT NULL THEN T" + joinNumStr + ".NEW_ID" + joinColName + " ELSE T1." + columnName
							+ " END AS " + columnName + "\n,T1." + columnName + " AS ORIG_" + columnName + "\n";
					sqlhead = false;
				} else {
					sql += (sqlhead ? " " : ",") + "T1." + columnName + "\n";
					sqlhead = false;
				}
			}
			// 最後一個table的尾部
			writeFile(sql, tableOldName, tableNewViewName, isIDColList, outputPath);

			// GRANT
			String grantee = "";
			for (Map<String, String> listDetail : list.get(1)) {
				cellNum = 0;
				tableNewName = listDetail.get("TableName");
				grantee = listDetail.get("Grantee");

				// 是否為新table
				if (!tableOldName.equals(tableNewName)) {
					tableOldViewName = "V_" + tableOldName + "_REPID";
					tableNewViewName = "V_" + tableNewName + "_REPID";

					if (!"".equals(tableOldName))
						FileTools.createFile(outputPath, tableOldViewName, "sql", sql);

					// 開始新table的頭
					sql = "";
					sheet = workbook.getSheet(tableNewName);
					rowNum = sheet.getLastRowNum() + 1;
					row = sheet.createRow(rowNum++); // 讓column 與 grant 之間多空一行
					row = sheet.createRow(rowNum++);
					Tools.setCell(Tools.getStyleRed(workbook), cell, row, cellNum, "GRANT");
				}
				tableOldName = tableNewName;
				row = sheet.createRow(rowNum++);
				Tools.setCell(cellStyle, cell, row, cellNum++, tableNewName);
				Tools.setCell(cellStyle, cell, row, cellNum++, grantee);

				sql += "grant select on nhiadm." + tableNewViewName + " to " + grantee + "; \n";
			}
			// 最後一個table的尾部
			FileTools.createFile(outputPath, tableNewViewName, "sql", sql);

			Tools.output(workbook, outputPath + "../", "Table Column.xlsx");

			buildIQ(iqPath, outputPath, tableNewViewNameArr);
		} catch (Exception ex) {
			throw new RuntimeException(className + " write Error: \n" + ex);
		} finally {
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException ex) {
				throw new RuntimeException(className + " finally Error: \n" + ex);
			}
		}
	}

	/**
	 * 將from與join資訊組進sql裡
	 * 
	 * @param sql
	 * @param tableOldName
	 * @param tableViewName
	 * @param isIDColList
	 * @param outputPath
	 */
	private static void writeFile(String sql, String tableOldName, String tableViewName,
			List<Map<String, String>> isIDColList, String outputPath) {
		String joinColName = "", joinNumStr = "";
		String fileContent = sql + "\nFROM " + tableOldName + " T1 \n";
		for (Map<String, String> map : isIDColList) {
			joinColName = map.get("ColName");
			if (joinColName.contains("_JOIN")) {
				joinNumStr = map.get("JoinNum");
				fileContent += " LEFT JOIN ( select distinct ID,NEW_ID,NEW_ID_JOIN from nhiadm.DWU_FOREIGN_ID_MAP) T" + joinNumStr
						+ " on T1." + joinColName.substring(0, joinColName.indexOf("_JOIN")) + " = T" + joinNumStr
						+ ".ID \n";
			}
		}
		fileContent += ";";
		
		FileTools.createFile(outputPath, tableViewName, "sql", fileContent);
	}

	/**
	 * 建立IQ建置單
	 * 
	 * @param iqPath
	 * @param outputPathBuild
	 * @param tableNewViewNameArr
	 */
	private static void buildIQ(String iqPath, String outputPathBuild, List<String> tableNewViewNameArr) {

		Workbook workbook = Tools.getWorkbook(iqPath);
		Sheet sheet = workbook.getSheetAt(0);
		CreationHelper createHelper = workbook.getCreationHelper();
		Row row = null;
		Cell cell = null;
		Hyperlink link = null;

		try {
			CellStyle defStyle = Tools.getStyle(workbook);
			CellStyle hLinkStyle = Tools.getStyleHLink(workbook);

			int rowNum = 2;
			int cellNum = 0;
			for (String tableViewName : tableNewViewNameArr) {
				if(tableViewName.startsWith("V_V_"))
					continue;
				
				cellNum = 0;
				row = sheet.createRow(rowNum++);
				Tools.setCell(defStyle, cell, row, cellNum++, "");
//				Tools.setCell(defStyle, cell, row, cellNum++, String.valueOf(rowNum-2));
				cell = row.createCell(cellNum++);
				cell.setCellFormula("ROW()-2");
				cell.setCellStyle(defStyle);
				Tools.setCell(defStyle, cell, row, cellNum++, "DW");
				Tools.setCell(defStyle, cell, row, cellNum++, "Alter");
				Tools.setCell(defStyle, cell, row, cellNum++, "procedure");
				Tools.setCell(defStyle, cell, row, cellNum++, tableViewName);
				Tools.setCell(defStyle, cell, row, cellNum++, "");
				Tools.setCell(defStyle, cell, row, cellNum++, "");
				Tools.setCell(defStyle, cell, row, cellNum++, "");

				// 相關檔案
				cell = row.createCell(cellNum++);
				cell.setCellValue(tableViewName + ".sql");
				link = (Hyperlink) createHelper.createHyperlink(Hyperlink.LINK_FILE);
				link.setAddress("\\\\192.168.3.52\\scms\\iqTable\\" + tableViewName + ".sql");
				cell.setHyperlink(link);
				cell.setCellStyle(hLinkStyle);

				Tools.setCell(defStyle, cell, row, cellNum++, "");
				Tools.setCell(defStyle, cell, row, cellNum++, "");
			}

			String outputFileBuild = outputPathBuild.substring(
					outputPathBuild.lastIndexOf("/", outputPathBuild.lastIndexOf("/") - 1) + 1,
					outputPathBuild.length() - 1) + ".xls";

			Tools.output(workbook, outputPathBuild, outputFileBuild);
		} catch (Exception ex) {
			throw new RuntimeException(className + " write Error: \n" + ex);
		} finally {
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException ex) {
				throw new RuntimeException(className + " finally Error: \n" + ex);
			}
		}
	}
	
}

package com.gss;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteRCPT {
	private static final String className = WriteViewExcel.class.getName();

	public static void write(String outputPath, List<List<Map<String, String>>> list, String folderName) {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		CellStyle cellStyle = Tools.getStyle(workbook);
		CellStyle cellStyleTitle = Tools.getStyleTitle(workbook);
		CellStyle cellStyleRed = Tools.getStyleRed(workbook);
		Sheet sheet = null;
		Cell cell = null;
		Row row = null;
		String selCntSql = "", selMDSql = "", selIQSql = "", idCol = "", tableOldName = "",
				tableNewName = "", columnName = "", isID = "", tableNewViewName = "";
		
		try {
			int rowNum = 0;
			int cellNum = 0;
			sheet = workbook.createSheet("列表");
			workbook.createSheet("MD資訊");
			workbook.createSheet("IQ Schema");
			sheet.setDefaultColumnWidth(50);
			sheet.setColumnWidth(2, 30 * 256);
			row = sheet.createRow(rowNum++);
			Tools.setCell(cellStyleTitle, cell, row, cellNum++, "原table/view");
			Tools.setCell(cellStyleTitle, cell, row, cellNum++, "勾稽後之View");
			Tools.setCell(cellStyleTitle, cell, row, cellNum++, "影響欄位");

			selMDSql = "select * from ( \n";
			selIQSql = "select * from ( \n";
			// Table Column List
			for (Map<String, String> listDetail : list.get(0)) {
				tableNewName = listDetail.get("TableName");
				columnName = listDetail.get("ColumnName");
				isID = listDetail.get("IsID");

				tableNewViewName = "V_" + tableNewName + "_REPID";
				// 是否為新table
				if (!tableOldName.equals(tableNewName)) {
					selCntSql += "select 'nhiadm." + tableNewName + "' as t_name,count(1) cnt from nhiadm." + tableNewName + " UNION \n"
							+ "select 'nhiadm." + tableNewViewName + "' as t_name,count(1) cnt from nhiadm." + tableNewViewName + " UNION \n";

					cellNum = 0;
					row = sheet.createRow(rowNum++);
					Tools.setCell(cellStyleRed, cell, row, cellNum++, "nhiadm."+tableNewName);
					Tools.setCell(cellStyle, cell, row, cellNum++, "nhiadm."+tableNewViewName);

					// 組完TABLE與TABLE間的union
					if(!"".equals(tableOldName)) {
						idCol = idCol.substring(0, idCol.lastIndexOf(","));
						selMDSql += idCol + ") \n union \n";
						selIQSql += idCol + ") \n union \n";
					}
					
					selMDSql += "select TABLENM, FIELDNM, DATA_CAT, DATA_LENGTH, FIELD_LOGIC \n" + 
							"from md_field \n" + 
							"where tablenm = '"+tableNewViewName+"' \n" + 
							"and fieldnm in ( ";
					selIQSql += "SELECT creator, tname, cname, coltype, nulls, length \n" + 
							"FROM SYS.SYSCOLUMNS \n" + 
							"WHERE CREATOR='nhiadm' AND TNAME = upper('"+tableNewViewName+"') \n" + 
							"and cname in ( ";
					idCol = "";
				}
				tableOldName = tableNewName;


				// 是否為身分證欄位
				if ("Y".equals(isID)) {
					if(!"".equals(idCol))
						row = sheet.createRow(rowNum++);
					Tools.setCell(cellStyle, cell, row, cellNum, columnName);
					idCol += "'" + columnName + "', 'ORIG_" + columnName + "',";
				}
			}

			idCol = idCol.substring(0, idCol.lastIndexOf(","));
			selMDSql += idCol + ") \n) a \nORDER BY TABLENM, FIELD_LOGIC, FIELDNM ; \n";
			selIQSql += idCol + ") \n) a \nORDER BY TNAME, CNAME ; \n";
			selCntSql = selCntSql.substring(0, selCntSql.lastIndexOf("UNION")) + "; \n";
			
			FileTools.createFile(outputPath, "selMDSql", "sql", selMDSql);
			FileTools.createFile(outputPath, "selIQSql", "sql", selIQSql);
			FileTools.createFile(outputPath, "selCntSql", "sql", selCntSql);
			
			folderName = folderName.substring(0,folderName.indexOf(" "));
			Tools.output(workbook, outputPath, "系統維護問題紀錄單_" + folderName + " 欄位異動清單.xlsx");
			
		} catch (Exception ex) {
			throw new RuntimeException(className + " write Error: \n" + ex);
		}
	}
}

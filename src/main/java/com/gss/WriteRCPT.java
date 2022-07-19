package com.gss;

import java.io.IOException;
import java.util.Arrays;
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

public class WriteRCPT {
	private static final String className = WriteViewExcel.class.getName();

	public static void write(String outputPath, List<List<Map<String, String>>> list) {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		CellStyle cellStyle = Tools.getStyle(workbook);
		CellStyle cellStyleRed = Tools.getStyleRed(workbook);
		Sheet sheet = null;
		Cell cell = null;
		Row row = null;
		String selCntSql = "", selMDSql = "", selIQSql = "", tableOldName = "", tableNewName = "", columnName = "", isID = "",
				tableNewViewName = "";

		try {
			int rowNum = 0;
			int cellNum = 0;
			sheet = workbook.createSheet("列表");
			sheet.setDefaultColumnWidth(30);
//			sheetList.setColumnWidth(2, 30 * 256);
			row = sheet.createRow(rowNum++);
			Tools.setCell(cellStyleRed, cell, row, cellNum++, "原table/view");
			Tools.setCell(cellStyleRed, cell, row, cellNum++, "勾稽後之View");
			Tools.setCell(cellStyleRed, cell, row, cellNum++, "影響欄位");
			
			// Table Column List
			for (Map<String, String> listDetail : list.get(0)) {
				cellNum = 0;
				tableNewName = listDetail.get("TableName");
				columnName = listDetail.get("ColumnName");
				isID = listDetail.get("IsID");

				tableNewViewName = "V_" + tableNewName + "_REPID";
				// 是否為新table
				if (!tableOldName.equals(tableNewName)) {
					selCntSql += "select 'nhiadm." + tableNewName + "' as t_name,count(1) cnt from nhiadm." + tableNewName + " UNION \n"
							+ "select 'nhiadm." + tableNewViewName + "' as t_name,count(1) cnt from nhiadm." + tableNewViewName + " UNION \n";

					row = sheet.createRow(rowNum++);
					Tools.setCell(Tools.getStyleRed(workbook), cell, row, cellNum, "Column");
					

					// 組完TABLE與TABLE間的union
					if(!"".equals(tableOldName)) {
						selMDSql += ") \n union \n";
						selIQSql += ") \n union \n";
					}
					
					selMDSql += "select TABLENM, FIELDNM, DATA_CAT, DATA_LENGTH, FIELD_LOGIC \n" + 
							"from md_field \n" + 
							"where tablenm = '"+tableNewViewName+"' \n" + 
							"and fieldnm in ( ";
					selIQSql += "SELECT creator, tname, cname, coltype, nulls, length \n" + 
							"FROM SYS.SYSCOLUMNS \n" + 
							"WHERE CREATOR='nhiadm' AND TNAME = upper('"+tableNewViewName+"') \n" + 
							"and cname in ( ";
				}
				tableOldName = tableNewName;

				// 是否為身分證欄位
				if ("Y".equals(isID)) {
//					selMDSql += ""
//					// 3 InsFieldOrigSql
//					insFieldOrigSql += "insert into md_field "
//							+ "select SYSTEM_TYPE, '" + tableNewViewName + "', 'ORIG_'||FIELDNM, '原始'||FIELDCNM, DATA_CAT, DATA_LENGTH, PRIMARY_KEY, "
//							+ "NULL_FLAG, INIT_VALUE, ENCRYPT_FLAG, "
//							+ "(SELECT MAX(FIELD_SEQ_NO)+1 FROM MD_FIELD WHERE TABLENM = '" + tableNewViewName + "' AND FIELD_SEQ_NO < 999), "
//							+ "BLURRY_FLAG, VALID_S_DATE, VALID_E_DATE, DATA_CODE, "
//							+ "'原始申報資料之'||FIELDCNM||'。'||DATA_DESC, FORMULA, FIELD_LOGIC, 'GSS3526', sysdate "
//							+ "from md_field where tablenm = '" + tableNewName + "' and fieldnm = '" + columnName + "'; \n";
//					// 4 UpdFieldIDSql
//					updFieldIDSql += "update md_field "
//							+ "set DATA_DESC = '因應外來人口統一證號異動作業，將原始" + columnName + "與外來人口統一證號異動檔勾稽，若勾稽到則以新ID取代，否則維持原始" + columnName + "'||'。'||DATA_DESC, "
//							+ "FIELD_LOGIC = '與DWU_FOREIGN_ID_MAP勾稽得到，否則為原始" + columnName + "'||'。'||FIELD_LOGIC "
//							+ "where tablenm = '" + tableNewViewName + "' and fieldnm = '" + columnName + "'; \n";
				}
			}

			FileTools.createFile(outputPath, "selCntSql", "sql", selCntSql += " ; \n");
			
//			buildOracle(outputPath, fieldNameArr);
		} catch (Exception ex) {
			throw new RuntimeException(className + " write Error: \n" + ex);
		}
	}
	
	/**
	 * 建立ORACLE建置單
	 * 
	 * @param oraclePath
	 * @param outputPathBuild
	 * @param tableNewViewNameArr
	 */
	private static void buildOracle(String oraclePath, String outputPathBuild, List<String> fieldNameArr) {

		Workbook workbook = Tools.getWorkbook(oraclePath);
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
			for (String fieldName : fieldNameArr) {
				cellNum = 0;
				row = sheet.createRow(rowNum++);
				Tools.setCell(defStyle, cell, row, cellNum++, "");
				cell = row.createCell(cellNum++);
				cell.setCellFormula("ROW()-2");
				cell.setCellStyle(defStyle);
				Tools.setCell(defStyle, cell, row, cellNum++, "SC");
				Tools.setCell(defStyle, cell, row, cellNum++, "Alter");
				Tools.setCell(defStyle, cell, row, cellNum++, "Table");
				Tools.setCell(defStyle, cell, row, cellNum++, fieldName);
				Tools.setCell(defStyle, cell, row, cellNum++, "");

				// 相關檔案
				cell = row.createCell(cellNum++);
				cell.setCellValue(fieldName + ".sql");
				link = (Hyperlink) createHelper.createHyperlink(Hyperlink.LINK_FILE);
				link.setAddress("http://192.168.102.12/scms/med2table/" + fieldName + ".sql");
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

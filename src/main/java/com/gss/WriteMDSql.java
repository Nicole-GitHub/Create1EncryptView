package com.gss;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class WriteMDSql {
	private static final String className = WriteViewExcel.class.getName();

	public static void write(String oraclePath, String outputPath, List<List<Map<String, String>>> list) {

		String insTableSql = "", insFieldSql = "", insFieldOrigSql = "", updFieldIDSql = "", insTableVerSql = "",
				insFieldVerSql = "", tableOldName = "", tableNewName = "", columnName = "", isID = "",
				tableNewViewName = "";

		try {
			// Table Column List
			for (Map<String, String> listDetail : list.get(0)) {
				tableNewName = listDetail.get("TableName");
				columnName = listDetail.get("ColumnName");
				isID = listDetail.get("IsID");

				tableNewViewName = "V_" + tableNewName + "_REPID";
				// 是否為新table
				if (!tableOldName.equals(tableNewName)) {
					// 1 InsTableSql
					insTableSql += "insert into md_table " 
							+ "select SYSTEM_TYPE, DATA_TYPE, '" + tableNewViewName + "', TABLECNM||'(含新ID)', 'V', IMPORT_FRQ, ACCESS_LIMITED, "
							+ "'Y', AUTHORIZE, RESERVE_TIME, RESERVE_TYPE, RESERVE_FLOAT, RESERVE_FIELDNM, TRANSFER_TYPE, "
							+ "RESERVE_STATUS, DATA_SOURCE, BUSINESS_PURP, DATA_MEAN||'，並勾稽外來人口統一證號異動對照檔', OTHERS, FILE_LOGIC, APPLY_PIC, "
							+ "'GSS3526', sysdate, 'Y' "
							+ "from md_table where tablenm = '" + tableNewName + "'; \n";
					// 2 InsFieldSql
					insFieldSql += "insert into md_field "
							+ "select SYSTEM_TYPE, '" + tableNewViewName + "', FIELDNM, FIELDCNM, DATA_CAT, DATA_LENGTH, PRIMARY_KEY, "
							+ "NULL_FLAG, INIT_VALUE, ENCRYPT_FLAG, FIELD_SEQ_NO, BLURRY_FLAG, VALID_S_DATE, VALID_E_DATE, "
							+ "DATA_CODE, DATA_DESC, FORMULA, FIELD_LOGIC, 'GSS3526', sysdate "
							+ "from md_field where tablenm = '" + tableNewName + "'; \n";
					// 5 InsTableVerSql
					insTableVerSql += "insert into md_table_ver "
							+ "select SYSTEM_TYPE, DATA_TYPE, TABLENM, TABLECNM, 1, 'Y', TABLE_CAT, IMPORT_FRQ, ACCESS_LIMITED, "
							+ "DATA_LIMITED, AUTHORIZE, RESERVE_TIME, RESERVE_TYPE, RESERVE_FLOAT, RESERVE_FIELDNM, TRANSFER_TYPE, "
							+ "RESERVE_STATUS, DATA_SOURCE, BUSINESS_PURP, DATA_MEAN, OTHERS, FILE_LOGIC, APPLY_PIC, USER_ID, "
							+ "TRAN_DATE, DATA_LIMITED_IA "
							+ "from md_table where tablenm = '" + tableNewViewName + "'; \n";
					// 6 InsFieldVerSql
					insFieldVerSql += "insert into md_field_ver "
							+ "select SYSTEM_TYPE, TABLENM, FIELDNM, FIELDCNM, 1, 'Y', DATA_CAT, DATA_LENGTH, PRIMARY_KEY, NULL_FLAG, "
							+ "INIT_VALUE, ENCRYPT_FLAG, FIELD_SEQ_NO, BLURRY_FLAG, VALID_S_DATE, VALID_E_DATE, DATA_CODE, DATA_DESC, "
							+ "FORMULA, FIELD_LOGIC, USER_ID, TRAN_DATE "
							+ "from md_field where tablenm = '" + tableNewViewName + "'; \n";
				}
				tableOldName = tableNewName;

				// 是否為身分證欄位
				if ("Y".equals(isID)) {
					// 3 InsFieldOrigSql
					insFieldOrigSql += "insert into md_field "
							+ "select SYSTEM_TYPE, '" + tableNewViewName + "', 'ORIG_'||FIELDNM, '原始'||FIELDCNM, DATA_CAT, DATA_LENGTH, PRIMARY_KEY, "
							+ "NULL_FLAG, INIT_VALUE, ENCRYPT_FLAG, "
							+ "(SELECT MAX(FIELD_SEQ_NO)+1 FROM MD_FIELD WHERE TABLENM = '" + tableNewViewName + "' AND FIELD_SEQ_NO < 999), "
							+ "BLURRY_FLAG, VALID_S_DATE, VALID_E_DATE, DATA_CODE, "
							+ "'原始申報資料之'||FIELDCNM||'。'||DATA_DESC, FORMULA, FIELD_LOGIC, 'GSS3526', sysdate "
							+ "from md_field where tablenm = '" + tableNewName + "' and fieldnm = '" + columnName + "'; \n";
					// 4 UpdFieldIDSql
					updFieldIDSql += "update md_field "
							+ "set DATA_DESC = '因應外來人口統一證號異動作業，將原始" + columnName + "與外來人口統一證號異動檔勾稽，若勾稽到則以新ID取代，否則維持原始" + columnName + "'||'。'||DATA_DESC, "
							+ "FIELD_LOGIC = '與DWU_FOREIGN_ID_MAP勾稽得到，否則為原始" + columnName + "'||'。'||FIELD_LOGIC "
							+ "where tablenm = '" + tableNewViewName + "' and fieldnm = '" + columnName + "'; \n";
				}
			}

			FileTools.createFile(outputPath, "1_InsTableSql", "sql", insTableSql += "\nCOMMIT; \n");
			FileTools.createFile(outputPath, "2_InsFieldSql", "sql", insFieldSql += "\nCOMMIT; \n");
			FileTools.createFile(outputPath, "3_InsFieldOrigSql", "sql", insFieldOrigSql += "\nCOMMIT; \n");
			FileTools.createFile(outputPath, "4_UpdFieldIDSql", "sql", updFieldIDSql += "\nCOMMIT; \n");
			FileTools.createFile(outputPath, "5_InsTableVerSql", "sql", insTableVerSql += "\nCOMMIT; \n");
			FileTools.createFile(outputPath, "6_InsFieldVerSql", "sql", insFieldVerSql += "\nCOMMIT; \n");
			
			List<String> fieldNameArr = Arrays.asList(new String[] { "1_InsTableSql", "2_InsFieldSql",
					"3_InsFieldOrigSql", "4_UpdFieldIDSql", "5_InsTableVerSql", "6_InsFieldVerSql" });
			
			buildOracle(oraclePath, outputPath, fieldNameArr);
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

package com.gss;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Parser {
	private static final String className = Parser.class.getName();

	protected static List<List<Map<String, String>>> runParser(String path) {
		String tableName = "", columnName = "", grantee = "";
		boolean isID = false;
		Workbook workbook = null;
		List<List<Map<String, String>>> list = new LinkedList<List<Map<String, String>>>();
		try {
			List<Map<String, String>> listDetail = new LinkedList<Map<String, String>>();
			Map<String, String> map = null;
			workbook = Tools.getWorkbook(path);
			Sheet sheetTableColumn = workbook.getSheet("Table Column List");
			Sheet sheetIDColumn = workbook.getSheet("身分證");
			Sheet sheetGrant = workbook.getSheet("GRANT");

			// Table Column List
			for (Row row : sheetTableColumn) {
				isID = false;
				if (row.getRowNum() > 0 && Tools.isntBlank(row.getCell(1))) {

					tableName = Tools.getCellValue(row, 1, "SheetTableColumn Table Name");
					columnName = Tools.getCellValue(row, 2, "SheetTableColumn Column Name");

					// 判斷此否為id
					for (Row idRow : sheetIDColumn) {
						if (idRow.getRowNum() > 0 && Tools.isntBlank(idRow.getCell(1))
								&& tableName.equals(Tools.getCellValue(idRow, 1, "SheetID Table Name"))
								&& columnName.equals(Tools.getCellValue(idRow, 2, "SheetID Column Name"))) {
							isID = true;
							break;
						}
					}

					map = new LinkedHashMap<String, String>();
					map.put("TableName", tableName);
					map.put("ColumnName", columnName);
					map.put("IsID", isID ? "Y" : "N");
					listDetail.add(map);
				}
			}
			list.add(listDetail);

			// GRANT
			listDetail = new LinkedList<Map<String, String>>();
			for (Row row : sheetGrant) {
				if (row.getRowNum() > 0 && Tools.isntBlank(row.getCell(1))) {

					tableName = Tools.getCellValue(row, 2, "SheetGrant Table Name");
					grantee = Tools.getCellValue(row, 1, "SheetGrant Grantee");

					map = new LinkedHashMap<String, String>();
					map.put("TableName", tableName);
					map.put("Grantee", grantee);
					listDetail.add(map);
				}
			}
			list.add(listDetail);
			System.out.println("Parse Done!");
		} catch (Exception ex) {
			throw new RuntimeException(className + " Error: \n" + ex.getMessage());
		} finally {
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException ex) {
				throw new RuntimeException(className + " finally Error: \n" + ex);
			}
		}

		return list;
	}

}

package com.gss;

import java.io.IOException;
import java.util.Iterator;
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
		List<Map<String, String>> listDetail = new LinkedList<Map<String, String>>();
		List<Map<String, String>> listIDDetail = new LinkedList<Map<String, String>>();
		Map<String, String> map = null;
		Map<String, String> idMap = null;

		try {
			workbook = Tools.getWorkbook(path);
			Sheet sheetTableColumn = workbook.getSheet("Table Column List");
			Sheet sheetIDColumn = workbook.getSheet("身分證");
			Sheet sheetGrant = workbook.getSheet("GRANT");

			// 判斷是否為id
			for (Row idRow : sheetIDColumn) {
				if (idRow.getRowNum() > 0 && Tools.isntBlank(idRow.getCell(1))) {

					idMap = new LinkedHashMap<String, String>();
					idMap.put("TableName", Tools.getCellValue(idRow, 1, "SheetID Table Name"));
					idMap.put("ColumnName", Tools.getCellValue(idRow, 2, "SheetID Column Name"));
					listIDDetail.add(idMap);
				}
			}
			// 整理listIDDetail
			refreshIDList(listIDDetail);

			// Table Column List
			for (Row row : sheetTableColumn) {
				isID = false;
				if (row.getRowNum() > 0 && Tools.isntBlank(row.getCell(1))) {

					tableName = Tools.getCellValue(row, 1, "SheetTableColumn Table Name");
					columnName = Tools.getCellValue(row, 2, "SheetTableColumn Column Name");

					// 判斷是否為id
					for (Map<String, String> listIDDetailMap : listIDDetail) {
						if (tableName.equals(listIDDetailMap.get("TableName"))
								&& columnName.equals(listIDDetailMap.get("ColumnName"))) {
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

	/**
	 * 整理listIDDetail 將剩餘沒對應到的ColumnName加上_JOIN再加回原list中
	 * 
	 * @param listIDDetail
	 */
	private static void refreshIDList(List<Map<String, String>> listIDDetail) {
		List<Map<String, String>> list = new LinkedList<Map<String, String>>(listIDDetail);
		List<Map<String, String>> list2 = new LinkedList<Map<String, String>>(listIDDetail);
		Iterator<Map<String, String>> it = list.iterator();
		Iterator<Map<String, String>> it2;
		Map<String, String> map = null;
		Map<String, String> map2 = null;
		String tableName1 = "", tableName2 = "", columnName1 = "", columnName2 = "";

		while (it.hasNext()) { // 遍历Iterator对象
			map = it.next();
			
			it2 = list2.iterator();
			while (it2.hasNext()) { // 遍历Iterator对象
				map2 = it2.next();

				tableName1 = map.get("TableName");
				tableName2 = map2.get("TableName");
				columnName1 = map.get("ColumnName");
				columnName2 = map2.get("ColumnName");

				if (tableName1.equals(tableName2)
						&& columnName2.replace("_JOIN", "").equals(columnName1.replace("_JOIN", ""))
						&& !columnName2.equals(columnName1)) {
					it.remove();
					break;
				}

			}
		}

		/**
		 * 將剩餘沒對應到的ColumnName加上_JOIN再加回原list中
		 * 但若直接改原map再加回去時會發生list原本的值也被覆蓋掉(等於會出現兩個一樣的_JOIN欄位) 故再新建一個list3與map3來存放剩餘的資訊
		 */
		List<Map<String, String>> list3 = new LinkedList<Map<String, String>>();
		Map<String, String> map3 = null;
		it = list.iterator();
		while (it.hasNext()) { // 遍历Iterator对象
			map = it.next();

			map3 = new LinkedHashMap<String, String>();
			map3.put("TableName", map.get("TableName"));
			map3.put("ColumnName", map.get("ColumnName") + "_JOIN");
			list3.add(map3);
		}
		// 將list3 append進原list中，並加在最後面
		listIDDetail.addAll(listIDDetail.size(), list3);

	}
}

package com.gss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Tools {
	private static final String className = Tools.class.getName();
//	private static Workbook workbook;
//	private static File f;
//	private static String excelVersion;
//	private static CellStyle style;

	/**
	 * 取得 Excel的Workbook
	 * 
	 * @param path
	 * @return
	 */
	public static Workbook getWorkbook(String path) {
		InputStream inputStream = null;
		Workbook workbook = null;
//		String excelVersion = "";
		File f;
		try {
			f = new File(path);
			inputStream = new FileInputStream(f);
			String aux = path.substring(path.lastIndexOf(".") + 1);
			if ("XLS".equalsIgnoreCase(aux)) {
//				excelVersion = "2003";
				workbook = new HSSFWorkbook(inputStream);
			} else if ("XLSX".equalsIgnoreCase(aux)) {
//				excelVersion = "2007";
				workbook = new XSSFWorkbook(inputStream);
			} else {
				throw new Exception("檔案格式錯誤");
			}
		} catch (Exception ex) {
			// 因output時需要用到，故不可寫在finally內
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				throw new RuntimeException(className + " getWorkbook Error: \n" + e);
			}

			throw new RuntimeException(className + " getWorkbook Error: \n" + ex);
		} finally {
			try {
				if (inputStream != null)
					inputStream.close();
			} catch (IOException e) {
				throw new RuntimeException(className + " getWorkbook Error: \n" + e);
			}
		}
		return workbook;
	}

//	/**
//	 * 取得 Excel的Sheet
//	 * 
//	 * @param path
//	 * @return
//	 */
//	public static Sheet getSheet(String path, String sheetName) {
//		return getWorkbook(path).getSheet(sheetName);
//	}

	/**
	 * 寫出整理好的Excel檔案
	 * 
	 * @param outputPath
	 * @param outputFileName
	 */
	public static void output(Workbook workbook, String outputPath, String outputFileName) {
		OutputStream output = null;
		try {
			File f = new File(outputPath);
			if (!f.exists())
				f.mkdirs();

			f = new File(outputPath + outputFileName);
			output = new FileOutputStream(f);
			workbook.write(output);
		} catch (Exception ex) {
			throw new RuntimeException(className + " output Error: \n" + ex);
		} finally {
			try {
				if (workbook != null)
					workbook.close();
				if (output != null)
					output.close();
			} catch (IOException ex) {
				throw new RuntimeException(className + " output finally Error: \n" + ex);
			}
		}
	}

	/**
	 * 設定寫出檔案時的Style
	 */
	protected static CellStyle getStyle(Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		short borderStyle = CellStyle.BORDER_THIN;
		style.setBorderBottom(borderStyle); // 儲存格格線(下)
		style.setBorderLeft(borderStyle); // 儲存格格線(左)
		style.setBorderRight(borderStyle); // 儲存格格線(右)
		style.setBorderTop(borderStyle); // 儲存格格線(上)
		return style;
	}

	/**
	 * 設定寫出檔案時的Style
	 */
	protected static CellStyle getStyleRed(Workbook workbook) {
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 14);
		font.setColor(Font.COLOR_RED);
		font.setBold(true);
		font.setFontName("微軟正黑體");
		
		CellStyle style = workbook.createCellStyle();
		style.setFont(font);
		return style;
	}

	/**
	 * 設定Cell內容(含Style)
	 * 
	 * @param cell
	 * @param row
	 * @param cellNum
	 * @param cellValue
	 */
	public static void setCell(CellStyle style, Cell cell, Row row, int cellNum, String cellValue) {
		cell = row.createCell(cellNum);
		cell.setCellValue(cellValue);
		cell.setCellStyle(style);
	}

	/**
	 * 不為空
	 */
	protected static boolean isntBlank(Cell cell) {
		return cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK;
	}

//	/**
//	 * 中文欄位
//	 *  部份、身份 -> 部分、身分
//	 *  計劃 -> 計畫
//	 * 	迄 -> 訖
//	 *  記錄 -> 紀錄
//	 */
//	protected static String replaceFieldCName(String fieldCName) {
//
//		fieldCName = fieldCName.replace("部份", "部分");
//		fieldCName = fieldCName.replace("身份", "身分");
//		fieldCName = fieldCName.replace("計劃", "計畫");
//		fieldCName = fieldCName.replace("迄", "訖");
//		fieldCName = fieldCName.replace("記錄", "紀錄");
//		return fieldCName;
//	}

//	/**
//	 * 是否為需加密的欄位
//	 * sourceFieldEName.equals("HOSP_ID") || fieldCName.equals("ID")
//	 *	|| fieldCName.contains("身分證字號") || fieldCName.contains("身分證號") 
//	 *	|| fieldCName.contains("醫事機構代碼") || fieldCName.contains("特約藥局代號") 
//	 *	|| fieldCName.contains("保險對象健保ID") || fieldCName.contains("醫師ID") 
//	 *	|| fieldCName.contains("藥師ID")
//	 *	|| fieldCName.contains("投保單位或扣費單位") || fieldCName.contains("統一編號")
//	 */
//	public static boolean isEncrypt(String fieldCName, String sourceFieldEName) {
//		List<String> list = Arrays.asList(
//				new String[] {"身分證字號","醫事機構代碼","保險對象健保ID","藥師ID","投保單位或扣費單位","身分證號","特約藥局代號","醫師ID","統一編號","醫事人員代號"});
//		
//		if(sourceFieldEName.equals("HOSP_ID") || fieldCName.equals("ID"))
//			return true;
//		
//		for(String str : list) {
//			if(fieldCName.contains(str))
//				return true;
//		}
//		
//		return false;
//	}

//	/**
//	 * 寫入 REQandTableLayout Cell
//	 */
//	public static void setREQandTLCell(Cell cell, Row row, String dataLineSeq, String dataStartDate, String fieldEName, String fieldCName, String dataType, 
//			String dataLength, String pk, String nullable, String initValue, String isEncrypt, String type, String dataClass, 
//			String dataDesc, String sourceTableEName, String sourceFieldEName, String procRule, String chgDesc) {
//
//		setCell(cell, row, 0, dataLineSeq);
//		setCell(cell, row, 1, fieldEName);
//		setCell(cell, row, 2, fieldCName);
//		setCell(cell, row, 3, dataType);
//		setCell(cell, row, 4, dataLength);
//		setCell(cell, row, 5, pk);
//		setCell(cell, row, 6, nullable);
//		setCell(cell, row, 7, initValue);
//		setCell(cell, row, 8, isEncrypt);
//		if ("REQ".equals(type)) {
//			setCell(cell, row, 9, dataClass);
//			setCell(cell, row, 10, dataDesc);
//			setCell(cell, row, 11, sourceTableEName);
//			setCell(cell, row, 12, sourceFieldEName);
//			setCell(cell, row, 13, procRule);
//			setCell(cell, row, 14, chgDesc);
//		} else {
//			setCell(cell, row, 9, dataLineSeq);
//			setCell(cell, row, 10, "");
//			setCell(cell, row, 11, dataStartDate);
//			setCell(cell, row, 12, "");
//			setCell(cell, row, 13, dataClass);
//			setCell(cell, row, 14, dataDesc);
//			setCell(cell, row, 15, "");
//			setCell(cell, row, 16, "");
//		}
//	}

//	/**
//	 * 寫入 DataMart Cell
//	 */
//	public static void setDataMartCell(Cell cell, Row row, String dwTableEName, String fieldEName, String dataSource, String sourceTableEName, 
//			String tableCName, String sourceFieldEName, String fieldCName, String procRule) {
//
//		setCell(cell, row, 0, dwTableEName);
//		setCell(cell, row, 1, fieldEName);
//		setCell(cell, row, 2, dataSource);
//		setCell(cell, row, 3, sourceTableEName);
//		setCell(cell, row, 4, tableCName);
//		setCell(cell, row, 5, sourceFieldEName);
//		setCell(cell, row, 6, fieldCName);
//		setCell(cell, row, 7, procRule);
//		setCell(cell, row, 8, "DW");
//		
//	}

	/**
	 * 取今日日期 YYYY/MM/DD
	 * 
	 * @return
	 */
	public static String getToDay() {
		return new SimpleDateFormat("yyyy/MM/dd").format(new Date());
	}

	/**
	 * 取Excel欄位值
	 * 
	 * @param sheet
	 * @param rownum
	 * @param cellnum
	 * @param fieldName
	 * @return
	 */
	public static String getCellValue(Row row, int cellnum, String fieldName) {
		try {
			if (!Tools.isntBlank(row.getCell(cellnum)) || row.getCell(cellnum).getCellType() == Cell.CELL_TYPE_BLANK) {
				return "";
			} else if (row.getCell(cellnum).getCellType() == Cell.CELL_TYPE_NUMERIC) {
				return String.valueOf((int) row.getCell(cellnum).getNumericCellValue()).trim();
			} else if (row.getCell(cellnum).getCellType() == Cell.CELL_TYPE_STRING) {
				return row.getCell(cellnum).getStringCellValue().trim();
			}
		} catch (Exception ex) {
			throw new RuntimeException(className + " getCellValue " + fieldName + " 格式錯誤");
		}
		return "";
	}
}

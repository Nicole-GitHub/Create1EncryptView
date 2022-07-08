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
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Tools {
	private static final String className = Tools.class.getName();

	/**
	 * 取得 Excel的Workbook
	 * 
	 * @param path
	 * @return
	 */
	public static Workbook getWorkbook(String path) {
		InputStream inputStream = null;
		Workbook workbook = null;
		File f;
		try {
			f = new File(path);
			inputStream = new FileInputStream(f);
			String aux = path.substring(path.lastIndexOf(".") + 1);
			if ("XLS".equalsIgnoreCase(aux))
				workbook = new HSSFWorkbook(inputStream);
			else if ("XLSX".equalsIgnoreCase(aux))
				workbook = new XSSFWorkbook(inputStream);
			else
				throw new Exception("檔案格式錯誤");
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
	 * 設定寫出檔案時的Style (一般)
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
	 * 設定寫出檔案時的Style (粗紅字體)
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
	 * 設定寫出檔案時的Style (超連結)
	 */
	protected static CellStyle getStyleHLink(Workbook workbook) {
		CellStyle hLinkstyle = workbook.createCellStyle();
		short borderStyle = CellStyle.BORDER_THIN;
		hLinkstyle.setBorderBottom(borderStyle); // 儲存格格線(下)
		hLinkstyle.setBorderLeft(borderStyle); // 儲存格格線(左)
		hLinkstyle.setBorderRight(borderStyle); // 儲存格格線(右)
		hLinkstyle.setBorderTop(borderStyle); // 儲存格格線(上)

		Font hLinkfont = workbook.createFont();
		hLinkfont.setUnderline(Font.U_SINGLE); // 底線
		hLinkfont.setColor(HSSFColor.BLUE.index); // 顏色
		hLinkfont.setFontName("新細明體"); // 字型
		hLinkfont.setFontHeightInPoints((short) 12); // 大小
		hLinkstyle.setFont(hLinkfont);

		return hLinkstyle;
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
			if (!Tools.isntBlank(row.getCell(cellnum)) || row.getCell(cellnum).getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else if (row.getCell(cellnum).getCellType() == Cell.CELL_TYPE_NUMERIC)
				return String.valueOf((int) row.getCell(cellnum).getNumericCellValue()).trim();
			else if (row.getCell(cellnum).getCellType() == Cell.CELL_TYPE_STRING)
				return row.getCell(cellnum).getStringCellValue().trim();
		} catch (Exception ex) {
			throw new RuntimeException(className + " getCellValue " + fieldName + " 格式錯誤");
		}
		return "";
	}
}

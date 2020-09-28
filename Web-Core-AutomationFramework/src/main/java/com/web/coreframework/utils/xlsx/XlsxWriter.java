package com.web.coreframework.utils.xlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// delete column deleterow cell formatting
public class XlsxWriter {
	private String path;
	private FileInputStream fis = null;
	private FileOutputStream fout = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private Row row = null;
	private Cell cell = null;

	public XlsxWriter(String path) throws IOException {
		this.path = path;
		File f = new File(path);
		try {
			fis = new FileInputStream(f);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

	}

	public boolean setCellData(String sheetname, int rownum, String columnheader, String data) throws IOException {
		int columnindex = -1;
		try {
			if (workbook.getSheetIndex(sheetname) == -1) {
				throw new NullPointerException("Sheet does not exist");

			} else {
				sheet = workbook.getSheet(sheetname);
				row = sheet.getRow(0);
				for (int i = 1; i < row.getLastCellNum(); i++) {
					cell = row.getCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);
					if (cellToString(cell).equals(columnheader)) {
						columnindex = i;
					}
				}
				if (columnindex == -1) {
					return false;
				}
				row = sheet.getRow(rownum);
				if (row == null) {
					row = sheet.createRow(rownum);
				}
				cell = row.getCell(columnindex);
				if (cell == null) {
					cell = row.createCell(columnindex);
				}
				cell.setCellValue(data);
				fout = new FileOutputStream(path);
				workbook.write(fout);
				fout.close();
				return true;
			}
		} catch (NullPointerException e) {
			e.printStackTrace();
			return false;
		}

	}

	public boolean setColumnData(String sheetname, String columnheader, String[] columndata) throws IOException {
		XlsxReader sheetreader = new XlsxReader(path);
		int columnindex = -1;
		try {
			if (!sheetreader.isSheetExist(sheetname)) {
				throw new NullPointerException("Sheet not found");
			} else {
				columnindex = sheetreader.getColumnIndex(sheetname, columnheader);
				if (columnindex == -1) {
					row = sheet.getRow(0);
					cell=row.createCell(row.getLastCellNum());
					cell.setCellValue(columnheader);
					fout = new FileOutputStream(path);
					workbook.write(fout);
					fout.close();
					columnindex = row.getLastCellNum()-1;
				}
				for(int i=1;i<=sheet.getLastRowNum();i++) {
					row=sheet.getRow(i);
					if (row == null) {
						row = sheet.createRow(i);
					}
					cell = row.getCell(columnindex);
					if (cell == null) {
						cell = row.createCell(columnindex);
					}
					cell.setCellValue(columndata[i-1]);
					fout = new FileOutputStream(path);
					workbook.write(fout);
					fout.close();
				}
				return true;
			}
		} catch (NullPointerException e) {
			e.printStackTrace();
			return false;
		}
	}

	public boolean setRowData(String sheetname, int rownum,ArrayList<String> rowdata) throws IOException {
		rownum=rownum-1;
		XlsxReader sheetreader = new XlsxReader(path);
		try {
			if (!sheetreader.isSheetExist(sheetname)) {
				throw new NullPointerException("Sheet not found");
			} else {
				sheet=workbook.getSheet(sheetname);
				row=sheet.getRow(rownum);
				if(row==null) {
					row=sheet.createRow(rownum);
				}
				for(int i=0;i<rowdata.size();i++) {
					cell=row.getCell(i);
					if(cell==null) {
						cell=row.createCell(i);
					}
					cell.setCellValue(rowdata.get(i));
					fout = new FileOutputStream(path);
					workbook.write(fout);
					fout.close();
				}
				return true;
			}
			
		} catch (NullPointerException e) {
			e.printStackTrace();
			return false;
		}
	}
	
	
	public boolean addSheet(String sheetname) {
		FileOutputStream fOut;
		try {
			workbook.createSheet(sheetname);
			fOut = new FileOutputStream(path);
			workbook.write(fOut);
			fOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	public boolean removeSheet(String sheetname) {
		int index = workbook.getSheetIndex(sheetname);
		if (index == -1)
			return false;

		FileOutputStream fOut;
		try {
			workbook.removeSheetAt(index);
			fOut = new FileOutputStream(path);
			workbook.write(fOut);
			fOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	private String cellToString(Cell cell) {
		switch (cell.getCellType()) {
		case NUMERIC:
			return ((int) cell.getNumericCellValue() + "");
		case STRING:
			return cell.getStringCellValue();
		case BLANK:
			return "";
		default:
			return "Not a valid data type";
		}
	}

}

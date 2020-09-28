package com.web.coreframework.utils.xlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxReader {
	private FileInputStream fis = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private Row row = null;
	private Cell cell = null;

	// Constructor to initialize File and sheet variables
	public XlsxReader(String path) throws IOException {
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

	// Returns Cell data based on sheet name, row number and column number
	public String getCellData(String sheetname, int rownum, int colnum) throws IOException {
		try {
			if (workbook.getSheetIndex(sheetname) == -1) {
				throw new NullPointerException("Sheet does not exist");

			} else {
				sheet = workbook.getSheet(sheetname);
				if (rownum < 0 || rownum > sheet.getLastRowNum()) {
					System.out.println("Row out of bound");
				} else {
					for (int i = 0; i <= sheet.getLastRowNum(); i++) {
						row = sheet.getRow(i);
						if (colnum < 0 || colnum > row.getLastCellNum()) {
							System.out.println("Column out of bound");
							break;
						} else {
							for (int j = 0; j < row.getLastCellNum(); j++) {
								cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
								if (i == rownum && j == colnum) {
									return cellToString(cell);
								}
							}
						}
					}
				}
				workbook.close();
				return null;
			}
		} catch (NullPointerException e) {
			e.printStackTrace();
			return null;
		}
	}

	// Returns an ArrayList of all row data based on row number
	public ArrayList<String> getRowData(String sheetname, int rownum) {
		try {
			if (workbook.getSheetIndex(sheetname) == -1) {
				throw new NullPointerException("Sheet does not exist");
			} else {
				sheet = workbook.getSheet(sheetname);
				ArrayList<String> ar = new ArrayList<String>();
				row = sheet.getRow(rownum);
				for (int i = 0; i < row.getLastCellNum(); i++) {
					cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					ar.add(cellToString(cell));
				}
				return ar;
			}

		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}

	/* Returns a HashMap of key=column header and column data=ArrayList of values */
	public HashMap<String,ArrayList<String>> getColumnData(String sheetname, String columnheader) {
		int columnindex = -1;
		ArrayList<String> ar = new ArrayList<String>();
		HashMap<String,ArrayList<String>>hmap=new HashMap<String,ArrayList<String>>();
		try {
			if (workbook.getSheetIndex(sheetname) == -1) {
				throw new NullPointerException("Sheet does not exist");
			} else {
				columnindex = getColumnIndex(sheetname, columnheader);
			}
			try {
				if (columnindex == -1) {
					throw new Exception("Invalid column header");
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				row = sheet.getRow(i);
				if(row==null) {
					row=sheet.createRow(i);
				}
				cell = row.getCell(columnindex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				ar.add(cellToString(cell));
			}
			hmap.put(columnheader, ar);
			return hmap;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}

	// Returns a 2-D String Array of all row data based on column header and column
	public String[][] getRowData(String sheetname, String columnheader, String columndata) {
		sheet = workbook.getSheet(sheetname);
		row = sheet.getRow(0);
		ArrayList<Integer> rowindex = new ArrayList<Integer>();
		HashMap<String,ArrayList<String>>hmap=new HashMap<String,ArrayList<String>>();
		hmap.putAll(getColumnData(sheetname, columnheader));
		if (hmap.get(columnheader).contains(columndata)) {
			for (int i = 0; i < hmap.get(columnheader).size(); i++) {
				if (hmap.get(columnheader).get(i).equals(columndata)) {
					rowindex.add(i);
				}
			}
			String[][] rowsdata = new String[rowindex.size()][row.getLastCellNum()];
			for (int i = 0; i < rowindex.size(); i++) {
				ArrayList<String> ar1 = getRowData(sheetname, rowindex.get(i) + 1);
				for (int j = 0; j < row.getLastCellNum(); j++) {
					rowsdata[i][j] = ar1.get(j);
				}
			}
			return rowsdata;
		} else {
			return null;
		}

	}

	// convert the cell values to string to be used by all returning functions
	private String cellToString(Cell cell) {
		switch (cell.getCellType()) {
		case NUMERIC:
			return ((int) cell.getNumericCellValue() + "");
		case STRING:
			return cell.getStringCellValue();
		case BLANK:
			return null;
		default:
			return "Not a valid data type";
		}
	}

	// returns row count in integers
	public int getRowCount(String sheetname) {
		int index = workbook.getSheetIndex(sheetname);
		if (index == -1)
			return 0;
		else {
			sheet = workbook.getSheetAt(index);
			int number = sheet.getLastRowNum() + 1;
			return number;
		}
	}

	// Returns Boolean for sheet exist
	public Boolean isSheetExist(String sheetname) {
		if (workbook.getSheetIndex(sheetname) == -1) {
			return false;
		} else {
			return true;
		}
	}

	// Returns Column header index
	public int getColumnIndex(String sheetname, String columnheader) {
		int columnindex = -1;
		sheet = workbook.getSheet(sheetname);
		row = sheet.getRow(0);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			cell = row.getCell(i);
			if (cellToString(cell).equals(columnheader)) {
				columnindex = cell.getColumnIndex();
			}
		}
		return columnindex;
	}
	/*
	 * MethodsTo be added...... getColumnCount and getCellData(method overloading)
	 * based on sheet name,row number and column header
	 */
}

package cn.kingkang.office.until;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.kingkang.pojo.ExcelSheet;

/**
 * Excel表格工具类
 * 
 * @author kingkang
 */
public class ExcelUntil {
	private static Workbook workbook;
	private static FormateCell formateCell;// 单元格数据处理类

	/**
	 * 获得一个Excel表格所有的数据 TODO
	 * 
	 * @param file
	 *            文件
	 * @param formate
	 *            Cell表格数据转换结果
	 * @param startDataRow
	 *            sheet中从第几行数据开始是数据，包含cell表头，但不包含标题
	 * @return
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public static List<ExcelSheet> getExcelData(File file, FormateCell formate, int startDataRow)
			throws InvalidFormatException, IOException {
		workbook = WorkbookFactory.create(file);
		List<ExcelSheet> sheets = new ArrayList<>();
		if (formateCell == null) {
			formateCell = new FormateCell() {
			};
		} else {
			formateCell = formate;
		}
		/**
		 * 对每个sheet工作表解析
		 */
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			Sheet sheet = workbook.getSheetAt(i);
			sheets.add(getSheetData(sheet, startDataRow));
		}
		return sheets;
	}
	/**
	 * 向Excel表格中写入数据 TODO
	 * @param file
	 *            文件名
	 * @param object
	 *            数据
	 * @return
	 * @throws IOException
	 */
	public static boolean write2Excel(File file, Object[][] object) throws IOException {
		createWorkBook(file);// 创建workbook
		Sheet sheet = createSheet(null);
		writeData(sheet, object, 0);
		return flushExcel(file);
	}
	/**
	 * 向Excel表格中写入数据
	 * TODO
	 * @param excelSheet
	 * @param file
	 * @return
	 * @throws IOException
	 */
	public static boolean write2Excel(ExcelSheet excelSheet, File file) throws IOException {
		if (excelSheet == null) {
			return false;
		}
		if (excelSheet.getCell_title() == null && excelSheet.getData() == null) {
			return false;
		}
		Object[][] objects;
		if (excelSheet.getCell_title() != null) {
			if (excelSheet.getData() != null) {
				objects = new Object[excelSheet.getData().length + 1][excelSheet.getCell_title().length];
				objects[0] = excelSheet.getCell_title();
				for (int i = 1; i < objects.length; i++) {
					objects[i] = excelSheet.getData()[i - 1];
				}
			} else {
				objects = new Object[1][excelSheet.getCell_title().length];
				objects[0] = excelSheet.getCell_title();
			}
		} else {
			objects = excelSheet.getData();
		}
		int endCol = objects[0] == null ? 5 : objects[0].length;
		return write2Excel(file, objects, 0, 0, 0, endCol, excelSheet.getTitle());
	}
	/**
	 * 向Excel表格中写入带有合并行的数据 TODO
	 * 
	 * @param file
	 *            文件
	 * @param object
	 *            数据
	 * @param startRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param startCol
	 *            开始列
	 * @param endCol
	 *            结束列
	 * @param mergarData
	 *            合并行的数据
	 * @return
	 * @throws IOException
	 */
	public static boolean write2Excel(File file, Object[][] object, int startRow, int endRow, int startCol, int endCol,
			Object mergarData) throws IOException {
		createWorkBook(file);
		Sheet sheet = createSheet(null);
		writeDataMarger(sheet, startRow, endRow, startCol, endCol, mergarData);
		writeData(sheet, object, endRow + 1);
		return flushExcel(file);
	}
	/**
	 * 获得工作表中的数据 TODO
	 * 
	 * @param sheet
	 * @param startDataRow
	 * @return
	 */
	private static ExcelSheet getSheetData(Sheet sheet, int startDataRow) {
		ExcelSheet excelSheet = new ExcelSheet();
		if (sheet == null) {
			return null;
		}
		excelSheet.setSheetName(sheet.getSheetName());// 获得单元表名称
		excelSheet.setTitle(getTitle(sheet, startDataRow).toString());// 表头
		Object[] cell_Titile = getRowData(sheet.getRow(startDataRow));
		excelSheet.setCell_title(cell_Titile);// 单元格表头
		int lastRowNum = sheet.getLastRowNum();
		Object[][] objects = new Object[lastRowNum - startDataRow][cell_Titile.length];
		for (int i = startDataRow+1; i < lastRowNum; i++) {
			System.out.println((i-startDataRow)+""+sheet.getLastRowNum()+""+(lastRowNum-startDataRow));
			objects[i-startDataRow-1] = getRowData(sheet.getRow(i));
		}
		excelSheet.setData(objects);// 数据
		return excelSheet;
	}

	/**
	 * 读取每行的数据 TODO
	 * 
	 * @param row
	 *            Excel行
	 * @return 该行的数据
	 */
	private static Object[] getRowData(Row row) {
		if (row == null || row.getLastCellNum() <= 0) {
			return null;
		}
		int cell_Num = row.getLastCellNum() - row.getFirstCellNum();// 获得该行单元格数量
		Object[] objects = new Object[cell_Num];
		for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
			Cell cell = row.getCell(i);
			objects[i] = formateCell.formateCellValue(cell);
		}
		return objects;
	}
	/**
	 * 获得表头 TODO
	 * 
	 * @param sheet
	 * @param startDataRow
	 * @return
	 */
	private static Object getTitle(Sheet sheet, int startDataRow) {
		Row row = null;
		// 获得表头
		if (startDataRow > sheet.getFirstRowNum()) {
			for (int i = sheet.getFirstRowNum(); i < startDataRow; i++) {
				row = sheet.getRow(i);
				for (Cell cell : row) {
					return getMergedValues(sheet, row.getRowNum(), cell.getColumnIndex());
				}
			}
		}
		return null;
	}

	/**
	 * 获得合并行数据 TODO
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	private static Object getMergedValues(Sheet sheet, int row, int column) {
		int marger_Num = sheet.getNumMergedRegions();// 获得所有合并的单元格数量
		for (int i = 0; i < marger_Num; i++) {
			CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
			int firstColumn = cellRangeAddress.getFirstColumn();
			int lastColumn = cellRangeAddress.getLastColumn();
			int firstRow = cellRangeAddress.getFirstRow();
			int lastRow = cellRangeAddress.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					Row fRow = sheet.getRow(firstRow);
					Cell fCell = fRow.getCell(firstColumn);
					return formateCell.formateCellValue(fCell);
				}
			}
		}
		return null;
	}

	/**
	 * 判断是否是合并的行 TODO
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public static boolean isMergedRow(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row == firstRow && row == lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
	}

	/**
	 * 创建一个sheet TODO
	 * 
	 * @param workbook
	 * @param sheetName
	 *            工作表名称
	 * @return
	 * @throws IOException
	 * @throws Exception
	 */
	private static Sheet createSheet(String sheetName) throws IOException {
		if (workbook == null) {
			throw new IOException("workbook create filed");
		}
		if (sheetName == null) {
			return workbook.createSheet();
		} else {
			return workbook.createSheet(sheetName);
		}
	}

	/**
	 * 创建workBook TODO
	 * 
	 * @param file
	 *            文件名
	 * @throws FileNotFoundException
	 */
	private static void createWorkBook(File file) throws FileNotFoundException {
		if (file == null) {
			throw new FileNotFoundException("不能将file设为空");
		}
		if (file.getName().endsWith(".xlsx")) {
			workbook = new XSSFWorkbook();
		} else {
			workbook = new HSSFWorkbook();
		}
	}

	/**
	 * 向Excel表格中写入数据 TODO
	 * 
	 * @param sheet
	 *            工作表
	 * @param data
	 *            数据
	 * @return
	 */
	private static boolean writeData(Sheet sheet, Object[][] data, int startRow) {
		int rowNum = startRow;
		for (int i = 0; i < data.length; i++) {
			Row row = sheet.createRow(rowNum++);
			for (int j = 0; j < data[i].length; j++) {
				Cell cell = row.createCell(j);
				if (data[i][j] == null) {
					cell.setCellValue("");
				} else {
					cell.setCellValue(data[i][j].toString());
				}
			}
		}
		return true;
	}

	/**
	 * 刷新Excel表格 TODO
	 * 
	 * @param file
	 *            写入的文件
	 * @return
	 * @throws IOException
	 */
	private static boolean flushExcel(File file) throws IOException {
		if (workbook == null) {
			return false;
		}
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		if (fos != null) {
			fos.close();
			fos.flush();
		}
		return true;

	}

	/**
	 * 写入合并行数据 TODO
	 * 
	 * @param sheet
	 *            单元表
	 * @param startRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param startCol
	 *            开始列
	 * @param endCol
	 *            结束咧
	 * @param mergarData
	 *            合并行数据
	 * @return
	 */
	private static boolean writeDataMarger(Sheet sheet, int startRow, int endRow, int startCol, int endCol,
			Object mergarData) {
		Row row;
		if (startRow <= endRow) {
			row = sheet.createRow(startRow);
			for (int i = startRow + 1; i <= endRow; i++) {
				sheet.createRow(i);
			}
			sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startCol, endCol));// 设置单元格合并
		} else {
			row = sheet.createRow(startRow);
			throw new IndexOutOfBoundsException("Row has error:" + startRow + ">" + endRow);
		}
		Cell cell = row.createCell(startCol);
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cell.setCellStyle(cellStyle);// 设置居中
		cell.setCellValue(mergarData.toString());
		return true;
	}

}

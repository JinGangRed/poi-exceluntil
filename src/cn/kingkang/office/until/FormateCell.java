package cn.kingkang.office.until;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
/**
 * 转换Excel表格中的数据
 * @author kingkang
 */
public interface FormateCell {
	/**
	 * 默认全部转换成string类型
	 * TODO
	 * @param cell
	 * @return
	 */
	public default Object formateCellValue(Cell cell){
		int cellType = -1;
		if (cell == null) {
			return null;
		} else {
			cellType = cell.getCellType();
		}
		switch (cellType) {
		case Cell.CELL_TYPE_BLANK:// 空值单元格
			return "";
		case Cell.CELL_TYPE_BOOLEAN:// boolean类型数据
			return String.valueOf(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_FORMULA:// 公式类型数据
			return String.valueOf(cell.getCellFormula());
		case Cell.CELL_TYPE_NUMERIC:// 数值类型数据
			DataFormatter dataFormatter = new DataFormatter();
			return dataFormatter.formatCellValue(cell);
		case Cell.CELL_TYPE_STRING:// 字符串类型数据
			return String.valueOf(cell.getRichStringCellValue());
		default:
			return null;
		}
	}
}

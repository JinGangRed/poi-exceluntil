package cn.kingkang.pojo;


/**
 * Excel表格
 * @author kingkang
 */
public class ExcelSheet {

	private String sheetName;//工作表单名称
	private String title;//标题
	private Object[] cell_title;//表头
	private Object[][] data;//数据
	public ExcelSheet() {}
	
	public ExcelSheet(String sheetName) {
		this.sheetName = sheetName;
	}
	
	public ExcelSheet(String sheetName, String title) {
		this.sheetName = sheetName;
		this.title = title;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}


	public Object[] getCell_title() {
		return cell_title;
	}

	public void setCell_title(String[] cell_title) {
		this.cell_title = cell_title;
	}
	public void setCell_title(Object[] cell_title) {
		this.cell_title = cell_title;
	}
	
	public Object[][] getData() {
		return data;
	}

	public void setData(Object[][] data) {
		this.data = data;
	}

	@Override
	public String toString() {
		StringBuffer cellTitile = new StringBuffer("[");
		if(cell_title != null){
			for (int i = 0; i < cell_title.length; i++) {
				cellTitile.append(String.valueOf(cell_title[i])+" ");
			}
		}
		cellTitile.append("]");
		StringBuffer data_str = new StringBuffer("[");
		if(data != null){
			for (int i = 0; i < data.length; i++) {
				if(data[i] != null){
					for (int j = 0; j < data[i].length; j++) {
						data_str.append(data[i][j]+",");
					}
				}
			}
		}
		data_str.append("]");
		return "[ ExcelSheet :"
				+ "sheetName=" + sheetName 
				+ ", title=" + title 
				+ ", cell_title=" + cellTitile.toString()
				+ ", data=" + data_str.toString()+ "]";
	}
	
	
}

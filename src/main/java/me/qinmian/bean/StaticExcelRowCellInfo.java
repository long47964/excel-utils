package me.qinmian.bean;

public class StaticExcelRowCellInfo {

	private ExcelRowCellInfo[] cellInfoArray;
	
	public StaticExcelRowCellInfo() {}

	public StaticExcelRowCellInfo(ExcelRowCellInfo[] cellInfoArray) {
		super();
		this.cellInfoArray = cellInfoArray;
	}

	public ExcelRowCellInfo[] getCellInfoArray() {
		return cellInfoArray;
	}

	public void setCellInfoArray(ExcelRowCellInfo[] cellInfoArray) {
		this.cellInfoArray = cellInfoArray;
	}
	
}

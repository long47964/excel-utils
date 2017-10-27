package me.qinmian.bean;

public class ExcelRowCellInfo {

	private String value;
	
	private Boolean single;
	
	private Boolean autoCol;
	
	private Integer startRow;
	
	private Integer endRow;
	
	private Integer startCol;
	
	private Integer endCol ;
	
	private Short rowHightInPoint;
	
	private ExportCellStyleInfo cellStyleInfo;

	public ExcelRowCellInfo() {
	}

	public ExcelRowCellInfo(String value, Boolean single, Boolean autoCol, Integer startRow, Integer endRow,
			Integer startCol, Integer endCol, Short rowHightInPoint, ExportCellStyleInfo cellStyleInfo) {
		super();
		this.value = value;
		this.single = single;
		this.autoCol = autoCol;
		this.startRow = startRow;
		this.endRow = endRow;
		this.startCol = startCol;
		this.endCol = endCol;
		this.rowHightInPoint = rowHightInPoint;
		this.cellStyleInfo = cellStyleInfo;
	}



	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}
	

	public Boolean getSingle() {
		return single;
	}


	public void setSingle(Boolean single) {
		this.single = single;
	}


	public Boolean getAutoCol() {
		return autoCol;
	}


	public void setAutoCol(Boolean autoCol) {
		this.autoCol = autoCol;
	}


	public Integer getStartRow() {
		return startRow;
	}

	public void setStartRow(Integer startRow) {
		this.startRow = startRow;
	}

	public Integer getEndRow() {
		return endRow;
	}

	public void setEndRow(Integer endRow) {
		this.endRow = endRow;
	}

	public Integer getStartCol() {
		return startCol;
	}

	public void setStartCol(Integer startCol) {
		this.startCol = startCol;
	}

	public Integer getEndCol() {
		return endCol;
	}

	public void setEndCol(Integer endCol) {
		this.endCol = endCol;
	}

	public ExportCellStyleInfo getCellStyleInfo() {
		return cellStyleInfo;
	}

	public void setCellStyleInfo(ExportCellStyleInfo cellStyleInfo) {
		this.cellStyleInfo = cellStyleInfo;
	}

	public Short getRowHightInPoint() {
		return rowHightInPoint;
	}

	public void setRowHightInPoint(Short rowHightInPoint) {
		this.rowHightInPoint = rowHightInPoint;
	}
	
	
}

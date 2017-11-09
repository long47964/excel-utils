package me.qinmian.bean;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

public class ExportInfo {

	private String sheetName;
	
	private Integer headRow;
	
	private Integer dataRow;
	
	private Integer maxSheetSize;
	
	private Map<Field,ExportFieldInfo> fieldInfoMap;
	
	private List<SortableField> sortableFields;
	
	private int headRowCount;
	
	private StaticExcelRowCellInfo staticExcelRowCellInfo;
	
	private Short dataHightInPoint;
	
	private Short headHightInPoint;
	
	public ExportInfo() {
		super();
	}
	public ExportInfo(String sheetName, Integer headRow, Integer dataRow,
			Map<Field, ExportFieldInfo> fieldInfoMap,
			List<SortableField> sortableFields) {
		super();
		this.sheetName = sheetName;
		this.headRow = headRow;
		this.dataRow = dataRow;
		this.fieldInfoMap = fieldInfoMap;
		this.sortableFields = sortableFields;
	}

	public String getSheetName() {
		return sheetName;
	}


	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}


	public Integer getHeadRow() {
		return headRow;
	}


	public void setHeadRow(Integer headRow) {
		this.headRow = headRow;
	}


	public Integer getDataRow() {
		return dataRow;
	}


	public void setDataRow(Integer dataRow) {
		this.dataRow = dataRow;
	}

	public Map<Field, ExportFieldInfo> getFieldInfoMap() {
		return fieldInfoMap;
	}


	public void setFieldInfoMap(Map<Field, ExportFieldInfo> fieldInfoMap) {
		this.fieldInfoMap = fieldInfoMap;
	}
	
	public List<SortableField> getSortableFields() {
		return sortableFields;
	}
	
	public void setSortableFields(List<SortableField> sortableFields) {
		this.sortableFields = sortableFields;
	}
	public int getHeadRowCount() {
		return headRowCount;
	}
	public void setHeadRowCount(int headRowCount) {
		this.headRowCount = headRowCount;
	}
	
	public StaticExcelRowCellInfo getStaticExcelRowCellInfo() {
		return staticExcelRowCellInfo;
	}
	public void setStaticExcelRowCellInfo(StaticExcelRowCellInfo staticExcelRowCellInfo) {
		this.staticExcelRowCellInfo = staticExcelRowCellInfo;
	}
	public Integer getMaxSheetSize() {
		return maxSheetSize;
	}
	public void setMaxSheetSize(Integer maxSheetSize) {
		this.maxSheetSize = maxSheetSize;
	}
	public Short getDataHightInPoint() {
		return dataHightInPoint;
	}
	public void setDataHightInPoint(Short dataHightInPoint) {
		this.dataHightInPoint = dataHightInPoint;
	}
	public Short getHeadHightInPoint() {
		return headHightInPoint;
	}
	public void setHeadHightInPoint(Short headHightInPoint) {
		this.headHightInPoint = headHightInPoint;
	}
	
}

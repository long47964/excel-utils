package me.qinmian.bean;

import java.util.Map;

public class ImportInfo {

	private Integer dataRow;
	private Integer headRow;
	private Map<String,ImportFieldInfo> fieldInfoMap;
	
	public ImportInfo() {}
	public ImportInfo(Integer dataRow, Integer headRow, Map<String,ImportFieldInfo> fieldInfoMap) {
		this.dataRow = dataRow;
		this.headRow = headRow;
		this.fieldInfoMap = fieldInfoMap;
	}
	public Integer getDataRow() {
		return dataRow;
	}
	public void setDataRow(Integer dataRow) {
		this.dataRow = dataRow;
	}
	public Integer getHeadRow() {
		return headRow;
	}
	public void setHeadRow(Integer headRow) {
		this.headRow = headRow;
	}
	public Map<String, ImportFieldInfo> getFieldInfoMap() {
		return fieldInfoMap;
	}
	public void setFieldInfoMap(Map<String, ImportFieldInfo> fieldInfoMap) {
		this.fieldInfoMap = fieldInfoMap;
	}
	
}

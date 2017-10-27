package me.qinmian.util;

import java.util.List;
import java.util.Map;

public class ImportResult<T> {

	private Integer suceess;
	private Integer fail;
	private Map<Integer,T> dataMap;
	private List<Integer> errorRows;
	public ImportResult() {
		super();
	}
	public ImportResult(Integer suceess, Integer fail, Map<Integer, T> dataMap, List<Integer> errorRows) {
		super();
		this.suceess = suceess;
		this.fail = fail;
		this.dataMap = dataMap;
		this.errorRows = errorRows;
	}
	public Integer getSuceess() {
		return suceess;
	}
	public void setSuceess(Integer suceess) {
		this.suceess = suceess;
	}
	public Integer getFail() {
		return fail;
	}
	public void setFail(Integer fail) {
		this.fail = fail;
	}
	public Map<Integer, T> getDataMap() {
		return dataMap;
	}
	public void setDataMap(Map<Integer, T> dataMap) {
		this.dataMap = dataMap;
	}
	public List<Integer> getErrorRows() {
		return errorRows;
	}
	public void setErrorRows(List<Integer> errorRows) {
		this.errorRows = errorRows;
	}
	@Override
	public String toString() {
		return "ImportResult [suceess=" + suceess + ", fail=" + fail + ", dataMap=" + dataMap + ", errorRows="
				+ errorRows + "]";
	}
	
	
}

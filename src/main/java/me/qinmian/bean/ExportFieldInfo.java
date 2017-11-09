package me.qinmian.bean;

import java.lang.reflect.Method;
import java.util.List;

import me.qinmian.bean.inter.ExportProcessor;
import me.qinmian.emun.DataType;

public class ExportFieldInfo {

	private List<Method> methodChain;
	
	private ExportCellStyleInfo headStyle;
	
	private ExportCellStyleInfo dataStyle;
	
	private String headName;
	
	private String dataFormat;
	
	private String dateFormat;
	
	private DataType dataType;
	
	private Short width;
	
	private Boolean autoWidth;
	
	private ExportProcessor exportProcessor;
			
	public ExportFieldInfo() {}
	
	public ExportFieldInfo(String headName) {
		this.headName =headName;
	}

	public ExportFieldInfo(List<Method> methodChain,
			ExportCellStyleInfo headStyle, ExportCellStyleInfo dataStyle,
			String headName, String dataFormat, String dateFormat,
			DataType dataType) {
		this.methodChain = methodChain;
		this.headStyle = headStyle;
		this.dataStyle = dataStyle;
		this.headName = headName;
		this.dataFormat = dataFormat;
		this.dateFormat = dateFormat;
		this.dataType = dataType;
	}

	public List<Method> getMethodChain() {
		return methodChain;
	}

	public void setMethodChain(List<Method> methodChain) {
		this.methodChain = methodChain;
	}

	public ExportCellStyleInfo getHeadStyle() {
		return headStyle;
	}

	public void setHeadStyle(ExportCellStyleInfo headStyle) {
		this.headStyle = headStyle;
	}

	public ExportCellStyleInfo getDataStyle() {
		return dataStyle;
	}

	public void setDataStyle(ExportCellStyleInfo dataStyle) {
		this.dataStyle = dataStyle;
	}

	public String getHeadName() {
		return headName;
	}

	public void setHeadName(String headName) {
		this.headName = headName;
	}

	public String getDataFormat() {
		return dataFormat;
	}

	public void setDataFormat(String dataFormat) {
		this.dataFormat = dataFormat;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}

	public DataType getDataType() {
		return dataType;
	}

	public void setDataType(DataType dataType) {
		this.dataType = dataType;
	}

	public Short getWidth() {
		return width;
	}

	public void setWidth(Short width) {
		this.width = width;
	}

	public Boolean getAutoWidth() {
		return autoWidth;
	}

	public void setAutoWidth(Boolean autoWidth) {
		this.autoWidth = autoWidth;
	}

	public ExportProcessor getExportProcessor() {
		return exportProcessor;
	}

	public void setExportProcessor(ExportProcessor exportProcessor) {
		this.exportProcessor = exportProcessor;
	}

}

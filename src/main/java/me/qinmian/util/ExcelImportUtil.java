package me.qinmian.util;

import java.beans.PropertyDescriptor;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.NumberUtils;
import org.springframework.util.StringUtils;

import me.qinmian.annotation.Excel;
import me.qinmian.annotation.ExcelField;

public  class ExcelImportUtil {

	private static String dataFormat = "yyyy-MM-dd HH:mm:ss";
	
	
	private static Map<Class<?>,ExcelInfo> excelInfoMap ;
	
	public static <T>  ImportResult<T> importExcel(String fileName,InputStream inputStream,Class<T> clazz) throws Exception{
		if(excelInfoMap == null ){
			excelInfoMap = new HashMap<Class<?>,ExcelInfo>();
		}
		Map<String,ExcelFieldInfo> fieldInfoMap;	
		Integer headNum = 0;
		Integer dataNum = 1;
		if(!excelInfoMap.isEmpty() && excelInfoMap.get(clazz) != null){
			ExcelInfo excelInfo = excelInfoMap.get(clazz);
			fieldInfoMap = excelInfo.getFieldInfoMap();
			dataNum = excelInfo.getDataNum();
			headNum = excelInfo.getHeadNum();
		}else{
			Excel excel = clazz.getAnnotation(Excel.class);
			if(excel != null ){
				dataNum = excel.dataRow();
				headNum = excel.headRow();
			}
			Field[] fields = clazz.getDeclaredFields();
			fieldInfoMap = new HashMap<String, ExcelFieldInfo>();
			for(Field field :fields){
				String name = field.getName();
				PropertyDescriptor pd = new PropertyDescriptor(name, clazz);
				Method setMethod = pd.getWriteMethod();
				ExcelField fieldInfo = field.getAnnotation(ExcelField.class);
				if(fieldInfo != null && !"".equals(fieldInfo.headName())){
					name = fieldInfo.headName();
				}	
				fieldInfoMap.put(name, new ExcelFieldInfo(field.getType(), setMethod, fieldInfo.required()));
			}	
			
			excelInfoMap.put(clazz, new ExcelInfo(dataNum,headNum,fieldInfoMap));
		}
		
		Workbook workbook ;
		if("xls".equalsIgnoreCase(fileName.substring(fileName.lastIndexOf(".")+1))){
			workbook = new HSSFWorkbook(inputStream);
		}else{
			workbook = new XSSFWorkbook(inputStream);
		}
		Sheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		if(rowCount < (headNum+1)){//
			return null;
		}
		List<String> cellNames;
		Row row;
		int cellCount;
		Cell cell;
		try {
			row = sheet.getRow(headNum);
			cellCount = row.getPhysicalNumberOfCells();
			cellNames = new ArrayList<String>();
			cell = null;
			for(int i = 0 ; i < cellCount ; i++){
				cell = row.getCell(i);
				cellNames.add(cell.getStringCellValue());			
			}
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
		Map<Integer,T> dataMap = new HashMap<Integer,T>();
		int suceess = 0 ; 
		int fail = 0;
		List<Integer> errorRows = new ArrayList<Integer>();
		rowLoop:
		for(int i = dataNum ; i < rowCount ; i++){
			row = sheet.getRow(i);
			T obj = clazz.newInstance();
			try {
				for(int j = 0 ; j < cellCount ; j++){					
					ExcelFieldInfo fieldInfo = fieldInfoMap.get(cellNames.get(j));
					cell = row.getCell(j);
					if(cell!=null){
						Object value = getValue(fieldInfo.getType(), cell);
						if(String.class.equals(fieldInfo.getType()) && StringUtils.isEmpty(value)){
							value = null;
						}
						if(value == null && fieldInfo.isRequired()){
							fail++;
							errorRows.add(i+1);
							continue rowLoop;
						}
						fieldInfo.getMethod().invoke(obj, value);															
					}
				} 
				suceess++;
				dataMap.put(i+1, obj);
			}catch (Exception e) {
				fail++;
				errorRows.add(i+1);
				e.printStackTrace();
				System.out.println(cell.toString());
			}
		}
		return new ImportResult<T>(suceess, fail, dataMap.isEmpty() ? null : dataMap, errorRows.isEmpty() ? null : errorRows);
	}
	
	public  static Object getValue(Class<?> type,Cell cell) throws Exception{
		int cellType = cell.getCellType();
		Object obj = null ;
		switch (cellType) {
		case Cell.CELL_TYPE_BLANK:
			return null;
			
		case Cell.CELL_TYPE_BOOLEAN:
			obj = cell.getBooleanCellValue();
			break;
			
		case Cell.CELL_TYPE_STRING:
			obj = cell.getStringCellValue();
			break;
			
		case Cell.CELL_TYPE_NUMERIC:
			obj = cell.getNumericCellValue();
			break;
			
		case Cell.CELL_TYPE_ERROR:
			return null;
		}
		obj = convertIfNeccesary(obj,type);
		return obj;
	}
	
	@SuppressWarnings("unchecked")
	private static Object convertIfNeccesary(Object obj, Class<?> type) throws ParseException {
		if(type.equals(obj.getClass())){
			return obj;
		}
		if(String.class.equals(type)){
			return obj.toString();
		}
		if(Number.class.isAssignableFrom(type)){
			if(Number.class.isAssignableFrom(obj.getClass())){
				return NumberUtils.convertNumberToTargetClass((Number)obj, (Class<Number>)type);
			}else if(String.class.equals(obj.getClass())){
				return NumberUtils.parseNumber(obj.toString(), (Class<Number>)type);
			}
		}
		if(String.class.equals(obj.getClass())){
			if(Date.class.equals(type)){
				SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dataFormat);
				return simpleDateFormat.parse(obj.toString());
			}
			if(Boolean.class.equals(type)){
				return Boolean.parseBoolean(obj.toString());
			}
		}
		return null;
	}

	public static class ExcelInfo{
		
		private Integer dataNum;
		private Integer headNum;
		private Map<String,ExcelFieldInfo> fieldInfoMap;
		
		public ExcelInfo() {
			super();
		}
		public ExcelInfo(Integer dataNum, Integer headNum, Map<String, ExcelFieldInfo> fieldInfoMap) {
			super();
			this.dataNum = dataNum;
			this.headNum = headNum;
			this.fieldInfoMap = fieldInfoMap;
		}

		public Integer getDataNum() {
			return dataNum;
		}

		public void setDataNum(Integer dataNum) {
			this.dataNum = dataNum;
		}

		public Integer getHeadNum() {
			return headNum;
		}

		public void setHeadNum(Integer headNum) {
			this.headNum = headNum;
		}

		public Map<String, ExcelFieldInfo> getFieldInfoMap() {
			return fieldInfoMap;
		}

		public void setFieldInfoMap(Map<String, ExcelFieldInfo> fieldInfoMap) {
			this.fieldInfoMap = fieldInfoMap;
		}
		
	}
	
	public static class ExcelFieldInfo{
		
		private Class<?> type;
		private Method method;
		private boolean required;
		public ExcelFieldInfo() {
			super();
		}
		public ExcelFieldInfo(Class<?> type, Method method, boolean required) {
			super();
			this.type = type;
			this.method = method;
			this.required = required;
		}
		public Class<?> getType() {
			return type;
		}
		public void setType(Class<?> type) {
			this.type = type;
		}
		public Method getMethod() {
			return method;
		}
		public void setMethod(Method method) {
			this.method = method;
		}
		public boolean isRequired() {
			return required;
		}
		public void setRequired(boolean required) {
			this.required = required;
		}
		
		
		
	}
}
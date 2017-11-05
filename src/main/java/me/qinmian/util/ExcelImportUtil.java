package me.qinmian.util;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import me.qinmian.annotation.Excel;
import me.qinmian.annotation.ExcelField;
import me.qinmian.bean.ImportFieldInfo;
import me.qinmian.bean.ImportInfo;

public  class ExcelImportUtil {

	private static final Logger LOG = LoggerFactory.getLogger(ExcelExportUtil.class);

//	private final static String DAFAULT_DATAFORMAT = "yyyy-MM-dd";
	
	private final static int DEFAULT_DATA_ROW = 1;
	
	private final static int DEFAULT_HEAD_ROW = 0;
	
	private final static DataFormatter DATA_FORMATTER = new DataFormatter();
	
	private final static Map<Class<?>,ImportInfo> importInfoMap = new HashMap<Class<?>,ImportInfo>();
	
	
	public static <T>  ImportResult<T> importExcel(String fileName,InputStream inputStream,Class<T> clazz) throws Exception{

		if(importInfoMap.get(clazz) == null){//初始化信息
			initForTargetClass(clazz);			
		}
		ImportInfo importInfo = importInfoMap.get(clazz);
		Integer headRow = importInfo.getHeadRow();
		Workbook workbook = createWorkbook(fileName, inputStream);
		
		int sheetNum = workbook.getNumberOfSheets();
		if(sheetNum < 1 ){
			return null;
		}
		Sheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		if(rowCount < (headRow+1)){//
			return null;
		}
		List<String> headNameList = createHeadNameList(sheet, headRow);
		return readData(clazz, importInfo, workbook,headNameList);
	}

	private static <T> ImportResult<T> readData(Class<T> clazz, ImportInfo importInfo, Workbook workbook, List<String> headNameList)
			throws InstantiationException, IllegalAccessException {
		int sheetNum = workbook.getNumberOfSheets();
		int dataRow = importInfo.getDataRow();
		Map<String, ImportFieldInfo> fieldInfoMap = importInfo.getFieldInfoMap();
		Sheet sheet;
		int rowCount;
		Row row;
		int cellCount = 0;
		Cell cell;		
		Map<Integer,T> dataMap = new HashMap<Integer,T>();
		int suceess = 0 ; 
		int fail = 0;
		List<Integer> errorRows = new ArrayList<Integer>();
		for(int sheetIndex = 0 ; sheetIndex < sheetNum ; sheetIndex++){
			sheet = workbook.getSheetAt(sheetIndex);
			rowCount = sheet.getPhysicalNumberOfRows();
			rowLoop:
				for(int i = dataRow ; i < rowCount ; i++){
					row = sheet.getRow(i);
					cellCount = row.getPhysicalNumberOfCells();
					T obj = clazz.newInstance();
					try {
						for(int j = 0 ; j < cellCount ; j++){					
							ImportFieldInfo fieldInfo = fieldInfoMap.get(headNameList.get(j));
							if(fieldInfo == null ){
								continue;
							}
							cell = row.getCell(j);
							if(cell != null){
								Object value = getValue(fieldInfo.getTypeChain().get(fieldInfo.getTypeChain().size() - 1), cell,"");
								if(String.class.equals(fieldInfo.getTypeChain().get(fieldInfo.getTypeChain().size() - 1)) && StringUtils.isEmpty(value)){
									value = null;
								}
								if(value == null && fieldInfo.isRequired()){
									fail++;
									errorRows.add(i+1);
									continue rowLoop;
								}
								invoke(obj,fieldInfo,value);
							}
						} 
						suceess++;
						dataMap.put(i+1, obj);
					}catch (Exception e) {
						fail++;
						errorRows.add(i+1);
						e.printStackTrace();
						if(LOG.isErrorEnabled()){
							LOG.error("第"+(i+1) + "行报错",e);
						}
					}
				}			
		}
		return new ImportResult<T>(suceess, fail, dataMap.isEmpty() ? null : dataMap, errorRows.isEmpty() ? null : errorRows);
	}
	
	private static void invoke(Object obj, ImportFieldInfo fieldInfo, Object value) throws Exception {
		List<Method> getMethodChain = fieldInfo.getGetMethodChain();
		List<Class<?>> typeChain = fieldInfo.getTypeChain();
		List<Method> setMethodChain = fieldInfo.getSetMethodChain();
		int i = 0 ;
		Object current;
		if(getMethodChain != null ){
			for(Method method : getMethodChain){
				current = method.invoke(obj);
				if(current == null ){
					current = typeChain.get(i+1).newInstance();
					setMethodChain.get(i).invoke(obj, current);
				}
				obj = current;
				i++;
			}
		}
		Method setMethod = setMethodChain.get(i);
		setMethod.invoke(obj, value);
	}


	private static List<String> createHeadNameList(Sheet sheet ,int headRow){
		List<String> headNameList;
		Row row;
		int cellCount;
		Cell cell;
		try {
			row = sheet.getRow(headRow);
			cellCount = row.getPhysicalNumberOfCells();
			headNameList = new ArrayList<String>();
			cell = null;
			for(int i = 0 ; i < cellCount ; i++){
				cell = row.getCell(i);
				headNameList.add(cell.getStringCellValue());			
			}
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
		return headNameList;
	}

	private static Workbook createWorkbook(String fileName, InputStream inputStream) throws IOException {
		Workbook workbook ;
		if("xls".equalsIgnoreCase(fileName.substring(fileName.lastIndexOf(".")+1))){
			workbook = new HSSFWorkbook(inputStream);
		}else{
			workbook = new XSSFWorkbook(inputStream);
		}
		return workbook;
	}

	private static <T> void initForTargetClass(Class<T> clazz)
			throws IntrospectionException {
		Integer headNum = DEFAULT_HEAD_ROW;
		Integer dataNum = DEFAULT_DATA_ROW;
		Map<String, ImportFieldInfo> fieldInfoMap;
		if(clazz.isAnnotationPresent(Excel.class)){
			Excel excel = clazz.getAnnotation(Excel.class);
			if(excel != null ){
				dataNum = excel.dataRow();
				headNum = excel.headRow();
			}			
		}
//		Field[] fields = clazz.getDeclaredFields();
		List<Field> fields = ClassUtils.getAllFieldFromClass(clazz);
		fieldInfoMap = new HashMap<String, ImportFieldInfo>();
		List<Class<?>> classChain = new ArrayList<Class<?>>();
		classChain.add(clazz);
		createFieldInfo(classChain, null , null ,fieldInfoMap, fields);		
		synchronized(importInfoMap){
			importInfoMap.put(clazz, new ImportInfo(dataNum,headNum,fieldInfoMap));			
		}
	}

	private static  void createFieldInfo(List<Class<?>> clazzChain, List<Method> setMethodChain, List<Method> getMethodChain ,
			Map<String, ImportFieldInfo> fieldInfoMap, List<Field> fields) throws IntrospectionException {
		Class<?> currentClazz = clazzChain.get(clazzChain.size()-1);
		for(Field field :fields){
			String name = field.getName();
			PropertyDescriptor pd = new PropertyDescriptor(name, currentClazz);
			Method setMethod = pd.getWriteMethod();
			List<Method> currentSetMethodChain = new ArrayList<Method>();
			if(setMethodChain != null && !setMethodChain.isEmpty()){
				currentSetMethodChain.addAll(setMethodChain);
			}
			currentSetMethodChain.add(setMethod);
			//clazzChain初始化就不为空
			List<Class<?>> currentClassChain = new ArrayList<Class<?>>(clazzChain);
			currentClassChain.add(field.getType());
			if(ClassUtils.isSimpleType(field.getType())){
				boolean required = false;
				String dateFormat = null ;
				if(field.isAnnotationPresent(ExcelField.class)){
					ExcelField excelField = field.getAnnotation(ExcelField.class);
					if(!"".equals(excelField.headName())){
						name = excelField.headName();
					}
					required = excelField.required();
					dateFormat = excelField.dateFormat();
				}
				fieldInfoMap.put(name, new ImportFieldInfo(currentClassChain,currentSetMethodChain,getMethodChain,required,dateFormat));				
			}else{
				
				List<Method> currentGetMehtodChain = new ArrayList<Method>();
				if(getMethodChain != null && !getMethodChain.isEmpty()){
					currentGetMehtodChain.addAll(getMethodChain);
				}
				currentGetMehtodChain.add(pd.getReadMethod());
				createFieldInfo(currentClassChain, currentSetMethodChain, currentGetMehtodChain, fieldInfoMap, 
								ClassUtils.getAllFieldFromClass(field.getType()));
			}
			
		}
	}
	
	public  static Object getValue(Class<?> type,Cell cell,String dateFormat) throws Exception{
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
			if(DateUtil.isCellDateFormatted(cell)){
				obj = DateUtil.getJavaDate(cell.getNumericCellValue());
			}else if(Number.class.isAssignableFrom(type) || ClassUtils.isBaseNumberType(type)){
				//当pojo字段类型是数字型时以数字形式获取
				obj = cell.getNumericCellValue();				
			}else{
				//其他类型都以string获取
				obj = DATA_FORMATTER.formatCellValue(cell);				
			}
			break;
			
		case Cell.CELL_TYPE_ERROR:
			return null;
		}
		obj = ConvertUtils.convertIfNeccesary(obj, type, dateFormat);
//		obj = convertIfNeccesary(obj,type);
		return obj;
	}
	
}
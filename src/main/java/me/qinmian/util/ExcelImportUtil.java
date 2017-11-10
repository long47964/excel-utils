package me.qinmian.util;

import java.beans.IntrospectionException;
import java.io.IOException;
import java.io.InputStream;
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

import me.qinmian.bean.ImportFieldInfo;
import me.qinmian.bean.ImportInfo;
import me.qinmian.emun.ExcelFileType;

public  class ExcelImportUtil {

	private static final Logger LOG = LoggerFactory.getLogger(ExcelExportUtil.class);
	
	private final static DataFormatter DATA_FORMATTER = new DataFormatter();
	
	private final static Map<Class<?>,ImportInfo> importInfoMap = new HashMap<Class<?>,ImportInfo>();
	
	
	
	/**
	 * @param clazz pojo对应class
	 * @param fileName 文件名
	 * @param inputStream 输入流
	 * @return 导入结果list
	 * @throws Exception
	 */
	public static <T>  List<T> importExcel(Class<T> clazz, String fileName,InputStream inputStream) throws Exception{	
		List<T> list = new ArrayList<T>();
		list.addAll(importExcel(fileName, inputStream, clazz).getDataMap().values());
		return list;
	}
		
	/**
	 * @param clazz pojo对应class
	 * @param fileType  excle文件类型
	 * @param inputStream 输入流
	 * @return 导入结果list
	 * @throws Exception
	 */
	public static <T>  List<T> importExcel(Class<T> clazz, ExcelFileType fileType,InputStream inputStream) throws Exception{	
		List<T> list = new ArrayList<T>();
		list.addAll(importExcel(fileType, inputStream, clazz).getDataMap().values());
		return list;
	}
	
	
	/**
	 * @param fileName 文件名
	 * @param inputStream 输入流
	 * @param clazz pojo对应class
	 * @return 导入结果
	 * @see me.qinmian.util.ImportResult
	 * @throws Exception
	 */
	public static <T>  ImportResult<T> importExcel(String fileName,InputStream inputStream,Class<T> clazz) throws Exception{	
		ExcelFileType fileType = getFileType(fileName);
		return importExcel(fileType, inputStream, clazz);
	}
	
	public static <T>  ImportResult<T> importExcel(ExcelFileType fileType,InputStream inputStream,Class<T> clazz) throws Exception{
		if(importInfoMap.get(clazz) == null){//初始化信息
			initTargetClass(clazz);			
		}
		ImportInfo importInfo = importInfoMap.get(clazz);
		Integer headRow = importInfo.getHeadRow();
		Workbook workbook = createWorkbook(fileType, inputStream);
		
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
	

	private static Workbook createWorkbook(ExcelFileType fileType, InputStream inputStream) throws IOException {
		Workbook workbook;
		if(ExcelFileType.XLS == fileType){
			workbook = new HSSFWorkbook(inputStream);
		}else{
			workbook = new XSSFWorkbook(inputStream);
		}
		return workbook;
	}
	
	private static ExcelFileType getFileType(String fileName) {
		if("xls".equalsIgnoreCase(fileName.substring(fileName.lastIndexOf(".")+1))){
			return ExcelFileType.XLS;
		}
		return ExcelFileType.XLSX;
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
								Object value = getValue(fieldInfo, cell);
								
								if(String.class.equals(fieldInfo.getTypeChain().get(fieldInfo.getTypeChain().size() - 1)) 
										&& StringUtils.isEmpty(value)){
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
	
	/** 根据获取的的数据执行set方法
	 * @param obj
	 * @param fieldInfo
	 * @param value
	 * @throws Exception
	 */
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


	/** 获取表头名
	 * @param sheet
	 * @param headRow
	 * @return
	 */
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

	/** 根据对象class初始化对应信息
	 * @param clazz
	 * @return
	 * @throws IntrospectionException
	 */
	private static <T> ImportInfo initTargetClass(Class<T> clazz)
			throws IntrospectionException {
		ImportInfo importInfo = ExcelUtils.getImportInfo(clazz);
		synchronized(importInfoMap){
			importInfoMap.put(clazz, importInfo);			
		}
		return importInfo;
	}
	
	/** 从表格之中读取数据并进行处理
	 * @param fieldInfo 
	 * @param cell 当前cell
	 * @return
	 * @throws Exception
	 */
	public  static Object getValue(ImportFieldInfo fieldInfo, Cell cell) throws Exception{
		int size = fieldInfo.getTypeChain().size();
		Class<?> type = fieldInfo.getTypeChain().get(size - 1);
		String dateFormat = fieldInfo.getDateFormat();
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
		if(fieldInfo.getImportProcessor() != null){
			obj = fieldInfo.getImportProcessor().process(obj);
		}
		obj = ConvertUtils.convertIfNeccesary(obj, type, dateFormat);
		return obj;
	}
	
}
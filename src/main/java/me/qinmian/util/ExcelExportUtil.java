package me.qinmian.util;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import me.qinmian.annotation.Excel;
import me.qinmian.annotation.ExcelField;
import me.qinmian.annotation.ExcelRowCell;
import me.qinmian.annotation.ExportCellStyle;
import me.qinmian.annotation.ExportFontStyle;
import me.qinmian.annotation.ExportStyle;
import me.qinmian.annotation.IgnoreField;
import me.qinmian.annotation.StaticExcelRow;
import me.qinmian.bean.ExcelFieldInfo;
import me.qinmian.bean.ExcelRowCellInfo;
import me.qinmian.bean.ExportCellStyleInfo;
import me.qinmian.bean.ExportFontStyleInfo;
import me.qinmian.bean.ExportInfo;
import me.qinmian.bean.SortComparator;
import me.qinmian.bean.SortableField;
import me.qinmian.bean.StaticExcelRowCellInfo;
import me.qinmian.emun.DataType;
import me.qinmian.emun.ExcelFileType;

public class ExcelExportUtil {
	
	private static final Logger LOG = LoggerFactory.getLogger(ExcelExportUtil.class);
	
	private final static int DEFAULT_HEAD_ROW = 0;
	
	private final static int DEFAULT_DATA_ROW = 1;
	
	private final static int DEFAULT_SORT = 100;
	
	private final static String DEFAULT_SHEET_NAME = "Sheet1";
	
	private final static String REGEX = "\\$\\{.*\\}";
	
	private final static String GET_METHOD_PREFIX = "get";
	
	private final static String SET_METHOD_PREFIX = "set";
	
	private final static String ANNO_KEY = "anno";
	
	private final static String STYLE_KEY = "style";
	
	private final static int XLX_MAX_SHEET_SIZE = 65536;
	
	private final static int XLXS_MAX_SHEET_SIZE = 1048576;
	
	private final static short DEFAULT_HIGHT_IN_POINT = 25;

	private static Map<Class<?> , ExportInfo> infoMap;	
	
	/**
	 *  key
	 *  value
	 */
	
	private static Map<Class<?>,Map<String,Map<Method,Method>>> classInfoCache;

	//初始化缓存信息
	static
	{
		try {
			Map<Method,Method> cellStyleMethodMap = new HashMap<Method,Method>();
			Map<Method,Method> cellAnno2InfoMap = new HashMap<Method,Method>();
			Map<Method,Method> fontStyleMethodMap = new HashMap<Method,Method>();
			Map<Method,Method> fontAnno2InfoMap = new HashMap<Method,Method>();
			Class<ExportCellStyleInfo> cellStyleInfoClass = ExportCellStyleInfo.class;
			Class<CellStyle> styleClass = CellStyle.class; 
			Class<ExportCellStyle> styleAnnoClass = ExportCellStyle.class;
			
			Class<ExportFontStyleInfo> fontStyleInfoClass = ExportFontStyleInfo.class;
			Class<ExportFontStyle> fontAnnoClass = ExportFontStyle.class;
			Class<Font> fontClass = Font.class;
			
			initCacheMap(cellStyleInfoClass,styleClass, styleAnnoClass,cellStyleMethodMap,
							cellAnno2InfoMap);
			
			initCacheMap(fontStyleInfoClass, fontClass, fontAnnoClass, fontStyleMethodMap, 
							fontAnno2InfoMap);
			Map<String,Map<Method,Method>> cellNameMap = new HashMap<String,Map<Method,Method>>(4);
			cellNameMap.put(ANNO_KEY, cellAnno2InfoMap);
			cellNameMap.put(STYLE_KEY, cellStyleMethodMap);
			Map<String,Map<Method,Method>> fontNameMap = new HashMap<String,Map<Method,Method>>(4);
			fontNameMap.put(ANNO_KEY, fontAnno2InfoMap);
			fontNameMap.put(STYLE_KEY, fontStyleMethodMap);
			classInfoCache = new HashMap<Class<?>,Map<String,Map<Method,Method>>>();
			classInfoCache.put(cellStyleInfoClass, cellNameMap);
			classInfoCache.put(fontStyleInfoClass, fontNameMap);
			
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		}		
	}
	
	private static void initCacheMap(Class<?> infoClass,
			Class<?> styleClass, Class<?> annoClass,
			Map<Method,Method> styleMethodMap, Map<Method,Method> anno2InfoMap)
			throws NoSuchMethodException {
		Method[] methods = styleClass.getMethods();			
		Map<String,Method> map = new HashMap<String,Method>();
		for(Method method : methods){
			map.put(method.getName(), method);
		}
		
		Field[] fields = infoClass.getDeclaredFields();
		for(Field field : fields){
			String fieldName = field.getName();
			if("fontStyleInfo".equals(fieldName)){
				continue;
			}
			String getMethodName = GET_METHOD_PREFIX + fieldName.substring(0, 1).toUpperCase() 
										+ fieldName.substring(1, fieldName.length());
			
			String setMethodName = SET_METHOD_PREFIX + fieldName.substring(0, 1).toUpperCase() 
					+ fieldName.substring(1, fieldName.length());
			Method infoGetMethod = infoClass.getMethod(getMethodName);
			Method infoSetMethod = infoClass.getMethod(setMethodName, field.getType());
			Method annoMethod = annoClass.getMethod(fieldName);
			styleMethodMap.put(infoGetMethod, map.get(setMethodName));
			anno2InfoMap.put(annoMethod, infoSetMethod);
		}
	}
	
	public static <T> Workbook exportXlsxExcel(Class<T> clazz, List<T> list) throws IntrospectionException, IllegalAccessException, IllegalArgumentException, InvocationTargetException{
		return exportExcel(clazz, list, ExcelFileType.XLSX, null);
	}
	
	public static <T> Workbook exportXlsExcel(Class<T> clazz, List<T> list) throws IntrospectionException, IllegalAccessException, IllegalArgumentException, InvocationTargetException{
		return exportExcel(clazz, list, ExcelFileType.XLS, null);
	}
	
	public static <T> Workbook exportExcel(Class<T> clazz, List<T> list ,ExcelFileType type) throws IntrospectionException, IllegalAccessException, IllegalArgumentException, InvocationTargetException{
		return exportExcel(clazz, list, type, null);
	}
	
	public static <T> Workbook exportExcel(Class<T> clazz, List<T> list , ExcelFileType type,Map<String,String> staticRowData ) throws IntrospectionException, IllegalAccessException, IllegalArgumentException, InvocationTargetException{	
		return exportExcel(clazz, list, type, staticRowData, null,false);
	}
	
/*	public static <T> InputStream exportExcelForStrem(Class<T> clazz, List<T> list , ExcelFileType type,Map<String,String> staticRowData ,boolean xlsxQuickMode) throws IntrospectionException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, IOException{	
		Workbook workbook = exportExcel(clazz, list, type, staticRowData, null, xlsxQuickMode);
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		ByteArrayInputStream in = new ByteArrayInputStream(outputStream.toByteArray());
		if(SXSSFWorkbook.class.equals(workbook.getClass())){
			SXSSFWorkbook wb = (SXSSFWorkbook)workbook;
			wb.dispose();
		}
		return in;
	}*/
	
	public static <T> Workbook exportExcel(Class<T> clazz, List<T> list , ExcelFileType type,Map<String,String> staticRowData , Workbook workbook,boolean xlsxQuickMode) throws IntrospectionException, IllegalAccessException, IllegalArgumentException, InvocationTargetException{	
		if(clazz == null){
			return null;
		}
		if(infoMap == null ){
			infoMap = new HashMap<Class<?>, ExportInfo>();
		}
		ExportInfo exportInfo = infoMap.get(clazz) ;
		if(exportInfo == null ){	
			exportInfo = initInfoForTargetClass(clazz);
		}
		
		int maxSheetSize = exportInfo.getMaxSheetSize(); 
		if(workbook == null){
			if(ExcelFileType.XLS == type){
				workbook = new HSSFWorkbook();
				if(maxSheetSize > XLX_MAX_SHEET_SIZE){
					maxSheetSize= XLX_MAX_SHEET_SIZE;					
				}
			}else if (ExcelFileType.XLSX == type){
				if(maxSheetSize > XLXS_MAX_SHEET_SIZE){
					maxSheetSize= XLXS_MAX_SHEET_SIZE;					
				}
				if(xlsxQuickMode){
					SXSSFWorkbook wb = new SXSSFWorkbook();
					//压缩临时文件
					wb.setCompressTempFiles(true);
					workbook = 	wb;
				}else{
					workbook = new XSSFWorkbook();
				}
			}		
		}
		Sheet sheet = null ;
		int dataRowNum = exportInfo.getDataRow();
		dataRowNum = dataRowNum > exportInfo.getHeadRow() + exportInfo.getHeadRowCount() ? 
						dataRowNum : exportInfo.getHeadRow() + exportInfo.getHeadRowCount();
		Map<Field,CellStyle> headCellStyleMap = createStyleMap(workbook,exportInfo,true);
		Map<Field,CellStyle> dataCellStyleMap = createStyleMap(workbook,exportInfo,false);

		List<Field> availableFields = new ArrayList<Field>();
		setAvailableFileds(exportInfo.getSortableFields() , availableFields);
		
		int startSheetNum = workbook.getNumberOfSheets();
		
		int usedRows = startSheetNum == 0 ? 0 : workbook.getSheetAt(startSheetNum-1).getPhysicalNumberOfRows() + 1;
		
		int i = startSheetNum == 0 ? 0 : startSheetNum - 1; 
		int page = (list.size() - usedRows + 1) / maxSheetSize + 1;
		if(usedRows == 0 ){
			page = (list.size() + 1) / maxSheetSize + 1;
		}else{
			page = (list.size()  - (maxSheetSize - usedRows) + 1)/maxSheetSize + 2 + i;
		}
		
		int usedListNum = 0;
		for( ; i < page ; i++  ){
			int startRow = dataRowNum > usedRows ? dataRowNum : usedRows;
			int dataSize = maxSheetSize - startRow;
			if(startSheetNum > 0){
				sheet = workbook.getSheetAt(i);
				startSheetNum = 0;
				usedRows = 0 ;
			}else{
				sheet = workbook.createSheet(exportInfo.getSheetName() + (i+1) );
				//创建静态行
				setStaticRows(exportInfo, staticRowData, workbook, sheet , availableFields.size());
				//创建表头
				createSheetHeadRow(exportInfo, headCellStyleMap, sheet);
			}
			createSheetDataRow(list, startRow, usedListNum , usedListNum + dataSize ,
					exportInfo,dataCellStyleMap,availableFields, sheet);	
			setColumWidth(type, exportInfo,availableFields, sheet);			
			usedListNum += dataSize;
		}
		
		return workbook;
	}

	private static void setAvailableFileds(List<SortableField> sortableFields, List<Field> availableFields) {
		for (SortableField sortField : sortableFields) {
			if(sortField.getChildFileds() ==  null){
				availableFields.add(sortField.getField());
			}else if(sortField.getChildFileds().isEmpty()){
				availableFields.add(null);
			}else{
				setAvailableFileds(sortField.getChildFileds(), availableFields);				
			}
		}
		
	}

	/** 根据缓存信息创建CellStyle
	 * @param workbook 工作簿
	 * @param exportInfo 缓存信息
	 * @param isHead 是否表头样式
	 * @return
	 */
	private static Map<Field, CellStyle> createStyleMap(Workbook workbook,ExportInfo exportInfo,boolean isHead) {
		Map<Field, CellStyle> styleMap;	
		styleMap = doCreateStyleMap(workbook, exportInfo, isHead);
		return styleMap == null ? null : styleMap;
	}


	/** 实际创建CellStyle Map
	 * @param workbook
	 * @param exportInfo
	 * @param isHead
	 * @return
	 */
	private static Map<Field, CellStyle> doCreateStyleMap(Workbook workbook, ExportInfo exportInfo,boolean isHead) {
		Map<ExportCellStyleInfo, CellStyle> tempCacheMap = new HashMap<ExportCellStyleInfo, CellStyle>();
		Map<Field, CellStyle> styleMap = new HashMap<Field, CellStyle>();
		CellStyle style;
		for(Map.Entry<Field,ExcelFieldInfo> entry : exportInfo.getFieldInfoMap().entrySet()){
			ExportCellStyleInfo styleInfo ;
			if(isHead){
				styleInfo = entry.getValue().getHeadStyle();
			}else{
				styleInfo = entry.getValue().getDataStyle();
			}
			if(!StringUtils.isEmpty(entry.getValue().getDataFormat())){
				//当存在格式化时，即使是来自通用的样式，但是格式不一样，所以需要new专属格式的样式
				//由于格式化属于专属，因此也不需要放到临时缓存map之中
				style = doCreateCellStyle(workbook,styleInfo,entry.getValue().getDataFormat());			
			}else{
				style = tempCacheMap.get(styleInfo);
				if(style == null){
					style = doCreateCellStyle(workbook,styleInfo,null);
					tempCacheMap.put(styleInfo, style);
				}
			}
			if(style != null ){
				styleMap.put(entry.getKey(),style);						
			}				
		}
		tempCacheMap.clear();
		return styleMap.isEmpty() ? null : styleMap;
	}

	/** 创建每个CellStyle，并放到map之中
	 * @param workbook
	 * @param styleInfo
	 * @param dataFormat 
	 * @return
	 */
	private static CellStyle doCreateCellStyle(Workbook workbook,ExportCellStyleInfo styleInfo,String dataFormat) {
		if(styleInfo != null ){
			CellStyle cellStyle = workbook.createCellStyle();
			setStyleValue(ExportCellStyleInfo.class, cellStyle, styleInfo);
			if(styleInfo.getFontStyleInfo() != null){
				Font fontStyle = workbook.createFont();
				setStyleValue(ExportFontStyleInfo.class, fontStyle, styleInfo.getFontStyleInfo());
				cellStyle.setFont(fontStyle);
			}
			if(!StringUtils.isEmpty(dataFormat)){
				short format = workbook.createDataFormat().getFormat(dataFormat);
				cellStyle.setDataFormat(format);
			}
			return cellStyle;
		}
		return null;
	}

	private static <T> void setStaticRows(ExportInfo exportInfo, Map<String, String> staticRowData, Workbook workbook,
											Sheet sheet , int colSize) {
		if(exportInfo.getStaticExcelRowCellInfo() != null){
			ExcelRowCellInfo[] cellInfoArray = exportInfo.getStaticExcelRowCellInfo().getCellInfoArray();
			if(cellInfoArray.length > 0 ){
				
				Cell cell = null;
				Row row;
				for(ExcelRowCellInfo rowCellInfo : cellInfoArray ){
					boolean single = rowCellInfo.getSingle();
					boolean autoCol = rowCellInfo.getAutoCol();
					int startRow = rowCellInfo.getStartRow();
					int endRow = rowCellInfo.getEndRow();
					int startCol = rowCellInfo.getStartCol();
					int endCol = rowCellInfo.getEndCol();
					String value = rowCellInfo.getValue();
					
					if(autoCol && colSize > 0 && startRow > -1){
						endRow = endRow < 0 ? startRow : endRow;
						CellRangeAddress cellRangeAddress = new CellRangeAddress(startRow, endRow,0, colSize - 1);
						sheet.addMergedRegion(cellRangeAddress);
						row = sheet.createRow(startRow);
						cell = row.createCell(0);
					}else if(single){
						if(startRow < 0 || startCol < 0){
							continue;
						}
						row = sheet.createRow(startRow);
						cell = row.createCell(startCol);
					}else{
						if(startRow == -1 || startCol == -1 || endRow == -1 || endCol == -1){
							continue;
						}
						CellRangeAddress cellRangeAddress = new CellRangeAddress(startRow, endRow,startCol, endCol);
						sheet.addMergedRegion(cellRangeAddress);
						row = sheet.getRow(startRow);
						if(row == null ){
							row = sheet.createRow(startRow);							
						}
						cell = row.createCell(startCol);
					}
					row.setHeightInPoints(rowCellInfo.getRowHightInPoint());
					if(rowCellInfo.getCellStyleInfo() != null){
						CellStyle style = doCreateCellStyle(workbook, rowCellInfo.getCellStyleInfo(), null);
						cell.setCellStyle(style);
					}
					if(value.matches(REGEX)){
						value = value.substring(2, value.length()-1);
						if(staticRowData != null ){
							value = staticRowData.get(value);							
						}else{
							value = null;
						}
					}
					if(!StringUtils.isEmpty(value)){
						cell.setCellValue(value);						
					}
				}
			}			
		}
	}

	private static void setColumWidth(ExcelFileType type,ExportInfo exportInfo,
			List<Field> availableFields,Sheet sheet) {
		Map<Field, ExcelFieldInfo> fieldInfoMap = exportInfo.getFieldInfoMap();
		for(int i = 0 ; i < availableFields.size() ; i++){
			Field field = availableFields.get(i);
			ExcelFieldInfo excelFieldInfo = fieldInfoMap.get(field);
			if (excelFieldInfo.getWidth() != null) {
				sheet.setColumnWidth(i, excelFieldInfo.getWidth()*256);	
				continue;
			}
			if(excelFieldInfo.getAutoWidth() != null){
				if(ExcelFileType.XLS == type){
					HSSFSheet hSheet = (HSSFSheet) sheet;
					hSheet.autoSizeColumn(i);
				}else if(ExcelFileType.XLSX == type){
					XSSFSheet xSheet = (XSSFSheet) sheet;
					xSheet.autoSizeColumn(i);
				}
			}
		}
	}

	private static <T> void createSheetDataRow(List<T> dataList, int dataRowNum,
			int listStart, int listEnd,ExportInfo exportInfo,
			Map<Field, CellStyle> dataCellStyleMap, List<Field> availableFields,Sheet sheet)
			throws IllegalAccessException, InvocationTargetException {
		
		Map<Field, ExcelFieldInfo> fieldInfoMap = exportInfo.getFieldInfoMap();
		Cell cell;
		Field field;
		Row row;
		T obj ;
		int dataSize = dataList.size() ;
		listEnd = listEnd > dataSize ? dataSize : listEnd ;
		int fieldListSize = availableFields.size();
		int dataHightInPoint = exportInfo.getDataHightInPoint();
		for(int i = listStart ; i < listEnd ; i++ , dataRowNum ++){
			obj = dataList.get(i);
			row = sheet.createRow(dataRowNum ); 
			row.setHeightInPoints(dataHightInPoint);
			for(int j = 0 ; j < fieldListSize ; j++){
				field = availableFields.get(j);
				if(field == null ){
					continue;
				}
				cell= row.createCell(j);
				Object returnVal = obj;
				List<Method> methods = fieldInfoMap.get(field).getMethodChain();
				for(Method method : methods){
					if (returnVal == null) {
						continue;
					}
					returnVal = method.invoke(returnVal);
				}
				setCellValue(cell, returnVal);
				if(dataCellStyleMap != null && dataCellStyleMap.get(field) != null){
					cell.setCellStyle(dataCellStyleMap.get(field));					
				}
			}
		}
	}

	private static void createSheetHeadRow(ExportInfo exportInfo,Map<Field,CellStyle> headCellStyleMap, Sheet sheet) {
		doCreateSheetHeadRow(exportInfo.getFieldInfoMap(),headCellStyleMap,sheet,exportInfo.getSortableFields(),
				exportInfo.getHeadRow(),0,exportInfo.getHeadRowCount(), exportInfo.getHeadHightInPoint());
	}




	private static void doCreateSheetHeadRow(Map<Field, ExcelFieldInfo> fieldInfoMap,
			Map<Field, CellStyle> headCellStyleMap, Sheet sheet, List<SortableField> fieldList, int rowNum, int cellNum,
			int headRowCount , short hightInPonit) {
		if(headRowCount == 0 ){
			Row row = sheet.getRow(rowNum) != null ? sheet.getRow(rowNum) : sheet.createRow(rowNum);
			row.setHeightInPoints(hightInPonit);
			doCreateSheetSingleHeadRow(fieldInfoMap,headCellStyleMap,row,fieldList,cellNum);
		}else{
			Cell cell;
			String value;
			Row row = null;
			for(int i = 0 ; i < fieldList.size() ; i++){
				int endRow = rowNum;
				int endCell = cellNum;
				List<SortableField> childFileds = fieldList.get(i).getChildFileds();
				if(childFileds == null || childFileds.isEmpty()){//
					endRow = rowNum + headRowCount;
				}else{//递归创建复合表头
					doCreateSheetHeadRow(fieldInfoMap, headCellStyleMap, sheet, childFileds, rowNum+1, cellNum, headRowCount-1,hightInPonit);
					//获取子元素数量，并以此创建合并单元格
					int count = countChildSize(fieldList.get(i));
					endCell = cellNum + count - 1;
				}		
				CellRangeAddress address = new CellRangeAddress(rowNum, endRow, cellNum, endCell);
				sheet.addMergedRegion(address);
				row = sheet.getRow(rowNum) != null ? sheet.getRow(rowNum) : sheet.createRow(rowNum);
				row.setHeightInPoints(hightInPonit);
				Field field = fieldList.get(i).getField();
				value = fieldInfoMap.get(field).getHeadName();
				cell = row.createCell(cellNum);
				cell.setCellValue(value);
				if( headCellStyleMap != null && headCellStyleMap.get(field) != null ){
					cell.setCellStyle(headCellStyleMap.get(field));
				}
				cellNum = endCell+1;
			}		
		}
	}

	private static void doCreateSheetSingleHeadRow(Map<Field, ExcelFieldInfo> fieldInfoMap,
			Map<Field, CellStyle> headCellStyleMap, Row row, List<SortableField> fieldList, int cellNum) {
		
		Cell cell;
		String value;
		for(int i = 0 ; i < fieldList.size() ; i++ , cellNum++){
			Field field = fieldList.get(i).getField();
			value = fieldInfoMap.get(field).getHeadName();
			cell = row.createCell(cellNum);
			cell.setCellValue(value);
			if( headCellStyleMap != null && headCellStyleMap.get(field) != null ){
				cell.setCellStyle(headCellStyleMap.get(field));
			}
		}
	}

	private static int countChildSize(SortableField sortableField) {
		if(sortableField.getChildFileds() == null || sortableField.getChildFileds().isEmpty()){
			return 0;			
		}
		int count = sortableField.getChildFileds().size() ;
		for(SortableField sf : sortableField.getChildFileds()){
			count += countChildSize(sf);
			if(!ClassUtils.isSimpleType(sf.getField().getType())){
				count-- ;
			}
		}
		return count;
	}
	private static ExportInfo initInfoForTargetClass(Class<?> clazz)
			throws IntrospectionException {
		ExportInfo exportInfo ;
		int headRowNum = DEFAULT_HEAD_ROW ; 
		int dataRowNum = DEFAULT_DATA_ROW ;
		String sheetName = DEFAULT_SHEET_NAME;
		int maxSheetSize = XLXS_MAX_SHEET_SIZE;
		
		short dataHightInPoint = DEFAULT_HIGHT_IN_POINT;
		short headHightInPoint = DEFAULT_HIGHT_IN_POINT;
		
		if(clazz.isAnnotationPresent(Excel.class)){
			Excel excel = clazz.getAnnotation(Excel.class);
			dataRowNum = excel.dataRow();
			headRowNum = excel.headRow();
			sheetName = excel.sheetName();
			maxSheetSize = excel.sheetSize();
			sheetName = StringUtils.isEmpty(sheetName) ? DEFAULT_SHEET_NAME : sheetName;
		}
		
		ExportCellStyleInfo globalHeadStyleInfo = null;
		ExportCellStyleInfo globalDataStyleInfo = null;
		//
		if(clazz.isAnnotationPresent(ExportStyle.class)){
			ExportStyle exportStyle = clazz.getAnnotation(ExportStyle.class);
			globalDataStyleInfo = createExportCellStyleInfo(exportStyle.dataStyle());
			globalHeadStyleInfo = createExportCellStyleInfo(exportStyle.headStyle());
			if(globalDataStyleInfo == null && exportStyle.dataEqHead()){
				globalDataStyleInfo = globalHeadStyleInfo;
			}
			dataHightInPoint = exportStyle.dataHightInPoint();
			headHightInPoint = exportStyle.headHightInPoint();
		}
		Map<Field,ExcelFieldInfo> fieldInfoMap = new HashMap<Field,ExcelFieldInfo>();
		//获取所有字段，包括父类字段
		List<Field> fieldList = getAllFieldFromClass(clazz);
		List<SortableField> sortFieldList = new ArrayList<SortableField>(fieldList.size());
		//递归创建Filed对应信息，并统计表头层数
		int count = createFieldInfo(clazz, globalHeadStyleInfo, globalDataStyleInfo, fieldInfoMap, fieldList,
				sortFieldList,null,0);
		//
		exportInfo = new ExportInfo(sheetName, headRowNum, dataRowNum, 
							fieldInfoMap,sortFieldList);
		exportInfo.setHeadRowCount(count);
		exportInfo.setMaxSheetSize(maxSheetSize);
		exportInfo.setDataHightInPoint(dataHightInPoint);
		exportInfo.setHeadHightInPoint(headHightInPoint);
		setStaticRowCellInfo(clazz,exportInfo);
		synchronized (infoMap) {
			infoMap.put(clazz, exportInfo);				
		}
		return exportInfo;
	}

	private static void setStaticRowCellInfo(Class<?> clazz, ExportInfo exportInfo) {
		if(clazz.isAnnotationPresent(StaticExcelRow.class)){
			StaticExcelRow staticRows = clazz.getAnnotation(StaticExcelRow.class);
			StaticExcelRowCellInfo staticExcelRowCellInfo = null;
			if(staticRows.cells().length > 0 ){
				ExcelRowCell[] cells = staticRows.cells();
				ExcelRowCellInfo[] rowCellInfos = new ExcelRowCellInfo[cells.length];
				int i = 0;
				for(ExcelRowCell rowCell : cells ){
					boolean single = rowCell.isSingle();
					boolean autoCol = rowCell.autoCol();
					int startRow = rowCell.startRow();
					int endRow = rowCell.endRow();
					int startCol = rowCell.startCol();
					int endCol = rowCell.endCol();
					String value = rowCell.value();
					short rowHightInPoint = rowCell.rowHightInPoint();
					ExportCellStyleInfo cellStyleInfo = createExportCellStyleInfo(rowCell.cellStyle());
					ExcelRowCellInfo rowCellInfo = new ExcelRowCellInfo(value, single, autoCol, startRow, endRow, 
														startCol, endCol, rowHightInPoint, cellStyleInfo);
					rowCellInfos[i++] = rowCellInfo;
				}
				staticExcelRowCellInfo = new StaticExcelRowCellInfo(rowCellInfos);
			}
			exportInfo.setStaticExcelRowCellInfo(staticExcelRowCellInfo);
		}
		
	}

	private static List<Field> getAllFieldFromClass(Class<?> clazz) {
		List<Field> fieldList = new ArrayList<Field>();
		Class<?> c = clazz;
		while(!Object.class.equals(c)){
			fieldList.addAll(Arrays.asList(c.getDeclaredFields()));
			c = c.getSuperclass();
		}
		return fieldList;
	}

	private static int createFieldInfo(Class<?> clazz, ExportCellStyleInfo parentHeadStyle,
			ExportCellStyleInfo parentDataStyle, Map<Field, ExcelFieldInfo> fieldInfoMap,
			List<Field> fieldList, List<SortableField> sortFieldList,
			List<Method> parentMethods,int count)
			throws IntrospectionException {
		for(int i = 0 ; i < fieldList.size() ; i++){
			Field field = fieldList.get(i);
			if(field.isAnnotationPresent(IgnoreField.class)){
				continue;
			}
			List<Method> methodChain = new ArrayList<Method>(1);
			if(parentMethods != null){
				methodChain.addAll(parentMethods);			
			}
			String name = field.getName();
			PropertyDescriptor pd = new PropertyDescriptor(name, clazz);
			Method getMethod = pd.getReadMethod();
			methodChain.add(getMethod);
			
			//获取字段的注解信息
			int sort = DEFAULT_SORT;
			ExcelFieldInfo fieldInfo = new ExcelFieldInfo(name);
			if(field.isAnnotationPresent(ExcelField.class)){
				ExcelField excelField = field.getAnnotation(ExcelField.class);
				sort = excelField.sort();
				setExcelFieldInfo(fieldInfo,excelField);
			}
			
			//读取style相关信息
			ExportCellStyleInfo fieldHeadStyleInfo = null;
			ExportCellStyleInfo fieldDataStyleInfo = null;		
			boolean required = StringUtils.isEmpty(fieldInfo.getDataFormat()) ? false : true;
			if(field.isAnnotationPresent(ExportStyle.class)){
				ExportStyle fieldExportStyle = field.getAnnotation(ExportStyle.class);
				fieldHeadStyleInfo = createExportCellStyleInfo(fieldExportStyle.headStyle());
				fieldDataStyleInfo = createExportCellStyleInfo(fieldExportStyle.dataStyle());
				if(fieldDataStyleInfo == null && fieldExportStyle.dataEqHead()){
					fieldDataStyleInfo = fieldHeadStyleInfo;
				}
			}
			fieldDataStyleInfo = fieldDataStyleInfo != null ? fieldDataStyleInfo : parentDataStyle;
			fieldHeadStyleInfo = fieldHeadStyleInfo != null ? fieldHeadStyleInfo :parentHeadStyle;
			//存在dataFormat时，必须创建样式
			if(fieldDataStyleInfo == null  && required){
				fieldDataStyleInfo = new ExportCellStyleInfo();
			}
			
			fieldInfo.setMethodChain(methodChain);
			fieldInfo.setHeadStyle(fieldHeadStyleInfo);
			fieldInfo.setDataStyle(fieldDataStyleInfo);
			fieldInfoMap.put(field, fieldInfo);
			
			List<SortableField> list = null;
			if(!ClassUtils.isSimpleType(field.getType())){
				list = new ArrayList<SortableField>();
				int temp = createFieldInfo(field.getType(), fieldHeadStyleInfo, fieldDataStyleInfo, 
						fieldInfoMap, getAllFieldFromClass(field.getType()), 
						list, methodChain,count+1);
				count = count > temp ? count : temp;
			}
			//每个类的字段
			sortFieldList.add(new SortableField(sort, field,list));		
		}
		//对每个类的字段进行排序
		Collections.sort(sortFieldList, new SortComparator());
		return count;
	}

	private static void setExcelFieldInfo(ExcelFieldInfo fieldInfo,
			ExcelField excelField) {
		if(!"".equals(excelField.headName())){
			fieldInfo.setHeadName(excelField.headName());
		}
		
		if(!"".equals(excelField.dataFormat())){
			fieldInfo.setDataFormat(excelField.dataFormat());
		}
		
		if(!"".equals(excelField.dateFormat())){
			fieldInfo.setDateFormat(excelField.dateFormat());
		}
		
		if(DataType.None !=excelField.dataType()){
			fieldInfo.setDataType(excelField.dataType());
		}
		
		if(-1 != excelField.width()){
			fieldInfo.setWidth(excelField.width());
		}
		if(!excelField.autoWidth()){
			fieldInfo.setAutoWidth(excelField.autoWidth());
		}
	}

	/** 给cell设置值ֵ
	 * @param cell
	 * @param returnVal
	 */
	private static void setCellValue(Cell cell, Object returnVal) {//以object来接收，会将基本数据类型自动装箱为包装类型
		if(returnVal == null ){
			return ;
		}
		if(String.class.equals(returnVal.getClass())){
			cell.setCellValue((String)returnVal);
		}else if(Boolean.class.equals(returnVal.getClass())){
			cell.setCellValue((Boolean)returnVal);
		}else if(Date.class.equals(returnVal.getClass())){
			cell.setCellValue((Date)returnVal);
		}else if (Number.class.isAssignableFrom(returnVal.getClass())){
			cell.setCellValue(((Number)returnVal).doubleValue());
		}else if(Character.class == returnVal.getClass()){
			cell.setCellValue(returnVal.toString());;
		}
	}
	
	
	private static ExportCellStyleInfo createExportCellStyleInfo(ExportCellStyle annoCellStyle){
		ExportFontStyle annoFontStyle = annoCellStyle.fontStyle();
		ExportFontStyleInfo fontStyleInfo = getExportInfo(ExportFontStyleInfo.class, annoFontStyle);
		ExportCellStyleInfo cellStyleInfo;
		if(fontStyleInfo  != null){
			cellStyleInfo = getExportInfo(ExportCellStyleInfo.class, annoCellStyle);
			cellStyleInfo.setFontStyleInfo(fontStyleInfo);
		}else{
			cellStyleInfo = getExportInfo(ExportCellStyleInfo.class, annoCellStyle);;			
		}
		return cellStyleInfo;

	}
	
	public static <T> T getExportInfo(Class<T> clazz,Object annotation){
		T t = null;
		try {
			t = clazz.newInstance();

			Map<Method, Method> anno2InfoMap = classInfoCache.get(clazz).get(ANNO_KEY);
			for (Map.Entry<Method, Method> entry : anno2InfoMap.entrySet()) {
				Method annoMethod = entry.getKey();
				Object value = annoMethod.invoke(annotation);
				entry.getValue().invoke(t, value);
			}
		} catch (Exception e) {
			LOG.error("创建ExportCellStyleInfo出错");
			e.printStackTrace();
		}
		return t;
	}

	public static <T> void setStyleValue(Class<T> clazz,Object style, Object info) {
		try {
			Map<Method, Method> styleMethodMap = classInfoCache.get(clazz).get(STYLE_KEY);
			for(Map.Entry<Method, Method> entry : styleMethodMap.entrySet()){
				Object value = entry.getKey().invoke(info);
				if(value != null ){
					entry.getValue().invoke(style, value);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}
}

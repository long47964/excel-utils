package me.qinmian.util;

import java.beans.IntrospectionException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import me.qinmian.bean.ExcelRowCellInfo;
import me.qinmian.bean.ExportFieldInfo;
import me.qinmian.bean.ExportInfo;
import me.qinmian.bean.SortableField;
import me.qinmian.bean.inter.ExportProcessor;
import me.qinmian.bean.inter.LinkProcessor;
import me.qinmian.emun.DataType;
import me.qinmian.emun.ExcelFileType;

/**
 *  pojo+注解 导出excel工具
 * @author woshi07948@163.com
 *
 */
public class ExcelExportUtil {
		
	private final static String REGEX = "\\$\\{.*\\}";
	
	private final static int XLX_MAX_SHEET_SIZE = 65536;
	
	private final static int XLXS_MAX_SHEET_SIZE = 1048576;
	
	private final static Map<Class<?> , ExportInfo> infoMap = new HashMap<Class<?>, ExportInfo>(8);;	

	
	//-------------------------07
	
	/**
	 * @param clazz 要导出的类
	 * @param list 导出的数据
	 * @return
	 */
	public static <T> Workbook exportExcel07(Class<T> clazz, List<T> list) {
		return exportExcel07(clazz, list, false);
	}
	
	/**
	 * @param clazz 要导出的类
	 * @param list 导出的数据
	 * @param quickMode 导出xlsx时是否开启快速模式，即使用SXSSFWorkbook类进行导出，快速模式会产生临时文件，
	 * 需要手动删除，默认为false 
	 * @return
	 */
	public static <T> Workbook exportExcel07(Class<T> clazz, List<T> list, boolean xlsxQuickMode){
		
		return exportExcel07(clazz, list,null, xlsxQuickMode);
	}
	
	public static <T> Workbook exportExcel07(Class<T> clazz, List<T> list , 
			Map<String,String> staticRowData){
		
		return exportExcel07(clazz, list, staticRowData, false);
	}
	
	public static <T> Workbook exportExcel07(Class<T> clazz, List<T> list , 
			Map<String,String> staticRowData, boolean xlsxQuickMode){
		
		return exportExcel07(clazz, list, staticRowData, null, xlsxQuickMode);
	}
	
	public static <T> Workbook exportExcel07(Class<T> clazz, List<T> list ,
			Map<String,String> staticRowData, Workbook workbook){
		return exportExcel07(clazz, list,staticRowData, workbook, false);
	}
	
	public static <T> Workbook exportExcel07(Class<T> clazz, List<T> list ,
					Map<String,String> staticRowData, Workbook workbook, boolean xlsxQuickMode){
		return exportExcel(clazz, list, ExcelFileType.XLSX, staticRowData, workbook, xlsxQuickMode);
	}
	
	//----------------------------07
	
	
	//----------------------------03
	/**
	 * @param clazz 要导出的类
	 * @param list 导出的数据
	 * @param staticRowData 静态行数据
	 * @param workbook 工作簿，分批写入时需要传入
	 * @return
	 */
	public static <T> Workbook exportExcel03(Class<T> clazz, List<T> list, Map<String,String> staticRowData , Workbook workbook){
		return exportExcel(clazz, list, ExcelFileType.XLS, staticRowData, workbook, false);
	}
	
	/**
	 * @param clazz 要导出的类
	 * @param list 导出的数据
	 * @param staticRowData 静态行数据
	 * @return
	 */
	public static <T> Workbook exportExcel03(Class<T> clazz, List<T> list, Map<String,String> staticRowData ){
		return exportExcel(clazz, list, ExcelFileType.XLS, staticRowData);
	}
	
	/**
	 * @param clazz 要导出的类
	 * @param list 导出的数据
	 * @return
	 */
	public static <T> Workbook exportExcel03(Class<T> clazz, List<T> list){
		return exportExcel(clazz, list, ExcelFileType.XLS);
	}
	//---------------------------03
	
	
	
	private static <T> Workbook exportExcel(Class<T> clazz, List<T> list, ExcelFileType type) {
		return exportExcel(clazz, list, type, null);
	}
	
	private static <T> Workbook exportExcel(Class<T> clazz, List<T> list, ExcelFileType type, 
			Map<String, String> staticRowData ){	
		return exportExcel(clazz, list, type, staticRowData, null,false);
	}
	
	/**
	 * @param clazz 要导出的类
	 * @param list 导出的数据
	 * @param type 导出表格类型
	 * @param staticRowData 静态行数据
	 * @param workbook 工作簿，分批写入时需要传入
	 * @param xlsxQuickMode 导出xlsx时是否开启快速模式，即使用SXSSFWorkbook类进行导出，快速模式会产生临时文件，
	 * 需要手动删除，默认为false 
	 * @return 工作簿
	 */
	private static <T> Workbook exportExcel(Class<T> clazz, List<T> list , ExcelFileType type, 
					Map<String, String> staticRowData , Workbook workbook, boolean xlsxQuickMode) {	
		if(clazz == null){
			return null;
		}
		ExportInfo exportInfo = infoMap.get(clazz) ;
		if(exportInfo == null ){	
			try {
				exportInfo = initTargetClass(clazz);
			} catch (IntrospectionException e) {
				throw new RuntimeException(e);
			}
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
		Map<Field,CellStyle> headCellStyleMap = CellStyleUtils.createStyleMap(workbook,exportInfo,true);
		Map<Field,CellStyle> dataCellStyleMap = CellStyleUtils.createStyleMap(workbook,exportInfo,false);

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
		try {
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
		} catch (Exception e) {
			throw new RuntimeException(e);
		}finally {
			//释放dateformate
			DateFormatHolder.remove();
		}
		return workbook;
	}
	
	
	private static ExportInfo initTargetClass(Class<?> clazz)
			throws IntrospectionException {
		ExportInfo exportInfo = ExcelUtils.getExportInfo(clazz);		
		synchronized (infoMap) {
			infoMap.put(clazz, exportInfo);				
		}
		return exportInfo;
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
						CellStyle style = CellStyleUtils.doCreateCellStyle(workbook, rowCellInfo.getCellStyleInfo(), null);
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
		Map<Field, ExportFieldInfo> fieldInfoMap = exportInfo.getFieldInfoMap();
		for(int i = 0 ; i < availableFields.size() ; i++){
			Field field = availableFields.get(i);
			ExportFieldInfo excelFieldInfo = fieldInfoMap.get(field);
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
		
		Map<Field, ExportFieldInfo> fieldInfoMap = exportInfo.getFieldInfoMap();
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
				
				setCellValue(cell, fieldInfoMap.get(field),  returnVal, obj);
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


	private static void doCreateSheetHeadRow(Map<Field, ExportFieldInfo> fieldInfoMap,
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

	/** 创建不需要合并单元格的表头
	 * @param fieldInfoMap
	 * @param headCellStyleMap
	 * @param row
	 * @param fieldList
	 * @param cellNum
	 */
	private static void doCreateSheetSingleHeadRow(Map<Field, ExportFieldInfo> fieldInfoMap,
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

	/** 给cell设置值
	 * @param cell
	 * @param returnVal
	 */
	private static void setCellValue(Cell cell, ExportFieldInfo fieldInfo, Object returnVal, Object current) {//以object来接收，会将基本数据类型自动装箱为包装类型
		
		ExportProcessor processor = fieldInfo.getExportProcessor();
		if(processor != null){
			if(LinkProcessor.class.isAssignableFrom(processor.getClass())){
				//处理链接类型
				setHyperlink(cell, (LinkProcessor) processor, returnVal, current);
			}
			returnVal = processor.process(returnVal, current);
		}else if(returnVal == null ){
			return ;
		}else if(fieldInfo.getDataType() != null && fieldInfo.getDataType() != DataType.None){
			try {
				switch (fieldInfo.getDataType()) {
				case String:
					returnVal = ConvertUtils.convertIfNeccesary(returnVal, String.class, fieldInfo.getDateFormat());
					break;
				case Number :
					returnVal = ConvertUtils.convertIfNeccesary(returnVal, Double.class, null);
					break;
				case Boolean :
					returnVal = ConvertUtils.convertIfNeccesary(returnVal, Boolean.class, null);
					break;
				case Date :
					returnVal = ConvertUtils.convertIfNeccesary(returnVal, Date.class, null);
					break;
				default:
					break;
				}
			} catch (ParseException e) {
				throw new RuntimeException(e);
			}
			
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
	
	
	/** 设置超链接属性
	 * @param cell
	 * @param processor
	 * @param returnVal
	 * @param current
	 */
	private static void setHyperlink(Cell cell, LinkProcessor processor, Object returnVal, Object current) {
		int linkType = 0 ;		
		String prefix = "";
		switch (processor.getLinkType()) {
		case Url:
			linkType = Hyperlink.LINK_URL;
			prefix = "http";
			break;
		case Document:
			linkType = Hyperlink.LINK_DOCUMENT;
			break;
		case Email:
			linkType = Hyperlink.LINK_EMAIL;
			prefix = "mailto:";
			break;
		case File:
			linkType = Hyperlink.LINK_FILE;
			break;
		}
		CreationHelper creationHelper = cell.getSheet().getWorkbook().getCreationHelper();
		org.apache.poi.ss.usermodel.Hyperlink hyperlink = creationHelper.createHyperlink(linkType);
		String address = processor.getLinkAddress(returnVal, current);
		if(!address.startsWith(prefix)){
			if(linkType == Hyperlink.LINK_EMAIL){
				address = prefix +  address;				
			}else{
				address = "http://" + address;
			}
		}
		hyperlink.setAddress(address);
		cell.setHyperlink(hyperlink);
	}

}

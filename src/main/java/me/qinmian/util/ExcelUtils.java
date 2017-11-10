package me.qinmian.util;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.util.StringUtils;

import me.qinmian.annotation.DataStyle;
import me.qinmian.annotation.Excel;
import me.qinmian.annotation.ExcelField;
import me.qinmian.annotation.ExcelRowCell;
import me.qinmian.annotation.ExportStyle;
import me.qinmian.annotation.HeadStyle;
import me.qinmian.annotation.IgnoreField;
import me.qinmian.annotation.StaticExcelRow;
import me.qinmian.bean.ExcelRowCellInfo;
import me.qinmian.bean.ExportCellStyleInfo;
import me.qinmian.bean.ExportFieldInfo;
import me.qinmian.bean.ExportInfo;
import me.qinmian.bean.ImportFieldInfo;
import me.qinmian.bean.ImportInfo;
import me.qinmian.bean.SortComparator;
import me.qinmian.bean.SortableField;
import me.qinmian.bean.StaticExcelRowCellInfo;
import me.qinmian.bean.inter.ExportProcessor;
import me.qinmian.bean.inter.ImportProcessor;
import me.qinmian.emun.DataType;

public class ExcelUtils {

	private final static int DEFAULT_HEAD_ROW = 0;
	
	private final static int DEFAULT_DATA_ROW = 1;
	
	private final static int DEFAULT_SORT = 100;
	
	private final static String DEFAULT_SHEET_NAME = "Sheet";
	
	private final static short DEFAULT_HIGHT_IN_POINT = 25;
	
	
	public static ExportInfo getExportInfo(Class<?> clazz) throws IntrospectionException{
		ExportInfo exportInfo ;
		int headRowNum = DEFAULT_HEAD_ROW ; 
		int dataRowNum = DEFAULT_DATA_ROW ;
		String sheetName = DEFAULT_SHEET_NAME;
		int maxSheetSize = Integer.MAX_VALUE - 1 ;
		
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
			globalDataStyleInfo = CellStyleUtils.createExportCellStyleInfo(exportStyle.dataStyle());
			globalHeadStyleInfo = CellStyleUtils.createExportCellStyleInfo(exportStyle.headStyle());
			if(globalDataStyleInfo == null && exportStyle.dataEqHead()){
				globalDataStyleInfo = globalHeadStyleInfo;
			}
			dataHightInPoint = exportStyle.dataHightInPoint();
			headHightInPoint = exportStyle.headHightInPoint();
		}
		if(clazz.isAnnotationPresent(HeadStyle.class)){
			HeadStyle headStyle = clazz.getAnnotation(HeadStyle.class);
			globalHeadStyleInfo = CellStyleUtils.createExportCellStyleInfo(headStyle.value());
			headHightInPoint = headStyle.hightInPoint();
		}
		
		if(clazz.isAnnotationPresent(DataStyle.class)){
			DataStyle dataStyle = clazz.getAnnotation(DataStyle.class);
			globalDataStyleInfo = CellStyleUtils.createExportCellStyleInfo(dataStyle.value());
			dataHightInPoint = dataStyle.hightInPoint();
		}
		
		Map<Field,ExportFieldInfo> fieldInfoMap = new HashMap<Field,ExportFieldInfo>();
		//获取所有字段，包括父类字段
		List<Field> fieldList = ClassUtils.getAllFieldFromClass(clazz);
		List<SortableField> sortFieldList = new ArrayList<SortableField>(fieldList.size());
		//递归创建Filed对应信息，并统计表头层数
		int count = createExportFieldInfo(clazz, globalHeadStyleInfo, globalDataStyleInfo, fieldInfoMap, fieldList,
				sortFieldList,null,0);
		//
		exportInfo = new ExportInfo(sheetName, headRowNum, dataRowNum, 
							fieldInfoMap,sortFieldList);
		exportInfo.setHeadRowCount(count);
		exportInfo.setMaxSheetSize(maxSheetSize);
		exportInfo.setDataHightInPoint(dataHightInPoint);
		exportInfo.setHeadHightInPoint(headHightInPoint);
		setStaticRowCellInfo(clazz,exportInfo);
		
		return exportInfo;
	}
	
	private static int createExportFieldInfo(Class<?> clazz, ExportCellStyleInfo parentHeadStyle,
			ExportCellStyleInfo parentDataStyle, Map<Field, ExportFieldInfo> fieldInfoMap,
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
			ExportFieldInfo fieldInfo = new ExportFieldInfo(name);
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
				fieldHeadStyleInfo = CellStyleUtils.createExportCellStyleInfo(fieldExportStyle.headStyle());
				fieldDataStyleInfo = CellStyleUtils.createExportCellStyleInfo(fieldExportStyle.dataStyle());
				if(fieldDataStyleInfo == null && fieldExportStyle.dataEqHead()){
					fieldDataStyleInfo = fieldHeadStyleInfo;
				}
			}
			if(field.isAnnotationPresent(HeadStyle.class)){
				HeadStyle headStyle = field.getAnnotation(HeadStyle.class);
				fieldHeadStyleInfo = CellStyleUtils.createExportCellStyleInfo(headStyle.value());
			}
			
			if(field.isAnnotationPresent(DataStyle.class)){
				DataStyle dataStyle = field.getAnnotation(DataStyle.class);
				fieldDataStyleInfo = CellStyleUtils.createExportCellStyleInfo(dataStyle.value());
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
				int temp = createExportFieldInfo(field.getType(), fieldHeadStyleInfo, fieldDataStyleInfo, 
						fieldInfoMap, ClassUtils.getAllFieldFromClass(field.getType()), 
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
	
	private static void setExcelFieldInfo(ExportFieldInfo fieldInfo,
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
		if(excelField.exportProcessor() != Void.class && 
				ExportProcessor.class.isAssignableFrom(excelField.exportProcessor())){
			Class<?> clazz = excelField.exportProcessor();
			try {
				ExportProcessor exportProcessor = (ExportProcessor) clazz.newInstance();
				fieldInfo.setExportProcessor(exportProcessor);
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}
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
					ExportCellStyleInfo cellStyleInfo = CellStyleUtils.createExportCellStyleInfo(rowCell.cellStyle());
					ExcelRowCellInfo rowCellInfo = new ExcelRowCellInfo(value, single, autoCol, startRow, endRow, 
														startCol, endCol, rowHightInPoint, cellStyleInfo);
					rowCellInfos[i++] = rowCellInfo;
				}
				staticExcelRowCellInfo = new StaticExcelRowCellInfo(rowCellInfos);
			}
			exportInfo.setStaticExcelRowCellInfo(staticExcelRowCellInfo);
		}	
	}
	
	public static <T> ImportInfo getImportInfo(Class<T> clazz) throws IntrospectionException{
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
		List<Field> fields = ClassUtils.getAllFieldFromClass(clazz);
		fieldInfoMap = new HashMap<String, ImportFieldInfo>();
		List<Class<?>> classChain = new ArrayList<Class<?>>();
		classChain.add(clazz);
		createImportFieldInfo(classChain, null , null ,fieldInfoMap, fields);			
		return new ImportInfo(dataNum,headNum,fieldInfoMap);			
		
	}
	
	private static  void createImportFieldInfo(List<Class<?>> clazzChain, List<Method> setMethodChain, List<Method> getMethodChain ,
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
				ImportProcessor processor = null;
				if(field.isAnnotationPresent(ExcelField.class)){
					ExcelField excelField = field.getAnnotation(ExcelField.class);
					if(!"".equals(excelField.headName())){
						name = excelField.headName();
					}
					required = excelField.required();
					dateFormat = excelField.dateFormat();
					Class<?> processorClazz = excelField.importProcessor();
					if(processorClazz != Void.class && 
							ImportProcessor.class.isAssignableFrom(processorClazz)){
						try {
							processor = (ImportProcessor) processorClazz.newInstance();
						} catch (Exception e) {
							throw new RuntimeException(e);
						}
					}
				}
				fieldInfoMap.put(name, new ImportFieldInfo(currentClassChain,currentSetMethodChain
						,getMethodChain,required,dateFormat,processor));				
			}else{
				
				List<Method> currentGetMehtodChain = new ArrayList<Method>();
				if(getMethodChain != null && !getMethodChain.isEmpty()){
					currentGetMehtodChain.addAll(getMethodChain);
				}
				currentGetMehtodChain.add(pd.getReadMethod());
				createImportFieldInfo(currentClassChain, currentSetMethodChain, currentGetMehtodChain, fieldInfoMap, 
								ClassUtils.getAllFieldFromClass(field.getType()));
			}
		}
	}//结束方法
	
	
}

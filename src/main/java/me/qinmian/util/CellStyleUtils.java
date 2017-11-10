package me.qinmian.util;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import me.qinmian.annotation.ExportCellStyle;
import me.qinmian.annotation.ExportFontStyle;
import me.qinmian.bean.ExportCellStyleInfo;
import me.qinmian.bean.ExportFieldInfo;
import me.qinmian.bean.ExportFontStyleInfo;
import me.qinmian.bean.ExportInfo;

/**
 * 样式工具，用于将注解转换为对应的InfoBean和创建CellStyle
 * @author woshi07948@163.com
 *
 */
public class CellStyleUtils {

	private static final Logger LOG = LoggerFactory.getLogger(CellStyleUtils.class);
	
	private final static String GET_METHOD_PREFIX = "get";
	
	private final static String SET_METHOD_PREFIX = "set";
	
	private final static String ANNO_KEY = "anno";
	
	private final static String STYLE_KEY = "style";
	/**
	 *  key
	 *  value
	 */
	private static Map<Class<?>,Map<String,Map<Method,Method>>> classInfoCache;
	
	private static List<String> excludeNames ;
	

	//初始化缓存信息
	static
	{
		try {
			excludeNames = new ArrayList<String>();
			excludeNames.add("fontStyleInfo");
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
			if(excludeNames.contains(fieldName)){
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
	
	
	
	/** 根据缓存信息创建CellStyle
	 * @param workbook 工作簿
	 * @param exportInfo 缓存信息
	 * @param isHead 是否表头样式
	 * @return
	 */
	public static Map<Field, CellStyle> createStyleMap(Workbook workbook,ExportInfo exportInfo,boolean isHead) {
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
		for(Map.Entry<Field,ExportFieldInfo> entry : exportInfo.getFieldInfoMap().entrySet()){
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
	public static CellStyle doCreateCellStyle(Workbook workbook,ExportCellStyleInfo styleInfo,String dataFormat) {
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
	
	private static <T> void setStyleValue(Class<T> clazz,Object style, Object info) {
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
			LOG.error("创建ExportCellStyleInfo出错",e);
			throw new RuntimeException(e);
		}
		return t;
	}
	
	
	/** 将注解信息转为对应的bean
	 * @param annoCellStyle
	 * @return
	 */
	public static ExportCellStyleInfo createExportCellStyleInfo(ExportCellStyle annoCellStyle){
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
}

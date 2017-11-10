package me.qinmian.util;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ClassUtils {

	private static final Map<Class<?>,Class<? extends Number>> WRAPCLASS_MAP = new HashMap<Class<?>,Class<? extends Number>>();
	
	static{
		WRAPCLASS_MAP.put(double.class, Double.class);
		WRAPCLASS_MAP.put(int.class, Integer.class);
		WRAPCLASS_MAP.put(short.class, Short.class);
		WRAPCLASS_MAP.put(long.class, Long.class);
		WRAPCLASS_MAP.put(float.class, Float.class);
		WRAPCLASS_MAP.put(byte.class, Byte.class);
	}
	
	public static boolean isSimpleType(Class<?> clazz){
		return CharSequence.class.isAssignableFrom(clazz) || Boolean.class == clazz 
				||Number.class.isAssignableFrom(clazz)
				|| Date.class == clazz || clazz.isPrimitive();
	}
	
	
	/** 判断是否是基本数据类型之中的数字类型
	 * @return
	 */
	public static boolean isBaseNumberType(Class<?> clazz){
		return int.class == clazz || short.class == clazz || float.class == clazz
				|| long.class == clazz || double.class == clazz || byte.class == clazz; 
	}
	
	@SuppressWarnings("unchecked")
	public static Class<? extends Number> getNumberWrapClass(Class<?> clazz){
		if(Number.class.isAssignableFrom(clazz)){
			return (Class<? extends Number>) clazz;
		}
		return WRAPCLASS_MAP.get(clazz);
	}
	
	public static List<Field> getAllFieldFromClass(Class<?> clazz) {
		List<Field> fieldList = new ArrayList<Field>();
		Class<?> c = clazz;
		while(!Object.class.equals(c)){
			fieldList.addAll(Arrays.asList(c.getDeclaredFields()));
			c = c.getSuperclass();
		}
		return fieldList;
	}
}

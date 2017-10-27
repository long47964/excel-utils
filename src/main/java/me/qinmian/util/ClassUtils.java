package me.qinmian.util;

import java.util.Date;

public class ClassUtils {

	public static boolean isSimpleType(Class<?> clazz){
		return CharSequence.class.isAssignableFrom(clazz) || Boolean.class == clazz 
				||Number.class.isAssignableFrom(clazz)
				|| Date.class == clazz || int.class == clazz|| double.class == clazz
				|| short.class == clazz|| byte.class == clazz|| long.class == clazz
				|| float.class == clazz|| char.class == clazz|| boolean.class == clazz;
	}
}

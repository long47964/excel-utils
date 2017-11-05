package me.qinmian.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.springframework.util.NumberUtils;

public class ConvertUtils {

	private final static String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";
	
	@SuppressWarnings("unchecked")
	public static Object convertIfNeccesary(Object obj, Class<?> type,String dateFormat) throws ParseException {
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
		if(ClassUtils.isBaseNumberType(type)){//基本类型之中的数字类型
			return NumberUtils.convertNumberToTargetClass((Number)obj, ClassUtils.getNumberWrapClass(type));
		}
		if(char.class.equals(type) || Character.class.equals(type)){
			if(obj.toString().length() > 0 ){
				return new Character(obj.toString().charAt(0));				
			}
			return null;
		}
		if(String.class.equals(obj.getClass())){
			if(Date.class.equals(type)){	
				SimpleDateFormat simpleDateFormat ;
				if(dateFormat != null){
					simpleDateFormat = new SimpleDateFormat(dateFormat);
				}else{
					simpleDateFormat = new SimpleDateFormat(DEFAULT_DATE_FORMAT);

				}
				return simpleDateFormat.parse(obj.toString());
			}
			if(Boolean.class.equals(type)){
				return Boolean.parseBoolean(obj.toString());
			}
		}
		return null;
	}
}

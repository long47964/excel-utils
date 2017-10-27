package me.qinmian.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import me.qinmian.emun.DataType;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelField {

	String headName() default "";
	
	String dataFormat() default "";
	
	String dateFormat() default "ss";
	
	DataType dataType() default DataType.None;
	
	boolean required() default false;
	
	int sort() default 100;
	
	short width() default -1;
	
	boolean autoWidth() default true;
		
}

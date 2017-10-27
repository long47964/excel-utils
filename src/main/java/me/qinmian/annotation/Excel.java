package me.qinmian.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Excel {
	
	int headRow() default 0;
	
	int dataRow() default 1 ;
	
	int sheetSize() default Integer.MAX_VALUE-1;
	
	String sheetName() default "";
	
}

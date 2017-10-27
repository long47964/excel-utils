package me.qinmian.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD,ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface ExportStyle {

	/** 是否数据行样式与表头行样式一致
	 * @return
	 */
	boolean dataEqHead() default false;
	
	ExportCellStyle headStyle() default @ExportCellStyle;
	
	ExportCellStyle dataStyle() default @ExportCellStyle;
	
	short dataHightInPoint() default 25;
	
	short headHightInPoint() default 25;
	
}

package me.qinmian.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelRowCell {
	
	String value() default "";
	/**是否是单个单元格，false代表为合并单元格
	 * @return
	 */
	boolean isSingle() default false;
	
	boolean autoCol() default false;
	
	int startRow() default -1;
	
	int endRow() default -1;
	
	int startCol() default -1;
	
	int endCol() default -1 ;
	
	short rowHightInPoint() default 25;
	
	ExportCellStyle cellStyle() default @ExportCellStyle();
}
